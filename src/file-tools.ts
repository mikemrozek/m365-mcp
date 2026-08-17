import { z } from 'zod';
import crypto from 'node:crypto';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type GraphClient from './graph-client.js';
import logger from './logger.js';
import { withUsageLog } from './usage-log.js';
import {
  extractText,
  UnsupportedFormatError,
  DEFAULT_MAX_CHARS,
  MAX_CHARS_LIMIT,
} from './lib/extract-text.js';

/**
 * Unified file handling: one tool per direction.
 *
 * Why one tool. Previously the model had to choose between get-mail-attachment
 * (inline base64), download-mail-attachment (stage + URL) and
 * read-mail-attachment-text (extracted text) — and it chose badly. Observed in
 * production: it anchored on whichever tool a previous message had named, hit
 * the sandbox egress wall, and concluded attachment reading was broken hours
 * after the working tool had shipped.
 *
 * WHAT THE SERVER DECIDES, AND WHAT IT DOES NOT (revised 2026-08-17)
 *
 * The server decides TRANSPORT — whether bytes come back inline or as a link,
 * based on size. It does NOT decide INTERPRETATION.
 *
 * The first version conflated the two: it extracted text and returned that,
 * which silently assumed the caller wanted prose. Scott's objection was correct
 * — "how do you know the caller even wanted a summary? What if it just had to
 * scan the document? Maybe it's a CSV and it's just looking for a specific
 * value." Our own test demonstrated the cost: a 51-page PDF came back truncated
 * at 50,000 of 74,747 characters, so anything in the last third was silently
 * unreachable.
 *
 * So the default is now a link, and the agent decides what to do with the file
 * — parse it, search it, run code over it, or read it. Extraction is still
 * available, but only when explicitly asked for via `as: 'text'`.
 *
 *   default            -> a short-lived pre-authenticated LINK
 *   tiny text files    -> inline, because a link would cost more than the content
 *   as: 'text'         -> extracted text, because the caller said so
 *
 * Bytes still never traverse the conversation unasked. That is a cost control:
 * a 2.9 MB attachment as base64 would consume a large share of a user's token
 * budget, which on a small plan can exhaust it in one call.
 */

/** Ceiling for explicit text extraction; above this a link is the only sane answer. */
const MAX_EXTRACT_BYTES = 40 * 1024 * 1024;

/**
 * Below this, a file is returned inline rather than as a link — the link,
 * metadata and an extra round trip would cost more than the content itself.
 * Deliberately small: this is the exception, not the rule.
 */
const INLINE_MAX_BYTES = 32 * 1024;

/** Simple PUT covers a single file to this size; beyond it Graph needs a session. */
const MAX_SIMPLE_UPLOAD_BYTES = 250 * 1024 * 1024;

/**
 * Graph rejects a direct attachment at or above 3 MB — the documented boundary
 * between `POST /attachments` and `createUploadSession`. Picking the wrong side
 * of this is the leading explanation for add-mail-attachment's ~71% failure
 * rate. Backed off slightly so encoding overhead cannot push a payload over.
 */
const DIRECT_ATTACH_LIMIT = 3 * 1024 * 1024 - 64 * 1024;

/** Where staged copies live, so they are identifiable and prunable later. */
const STAGING_FOLDER = 'Apps/TSQ-M365-MCP-staging';

function jsonResult(payload: unknown, isError = false) {
  return {
    content: [{ type: 'text' as const, text: JSON.stringify(payload) }],
    ...(isError ? { isError: true } : {}),
  };
}

/** Strips characters OneDrive rejects and prefixes for collision safety. */
function safeStagedName(name: string): string {
  const cleaned = name.replace(/[\\/:*?"<>|]/g, '_');
  return `${Date.now()}-${crypto.randomBytes(4).toString('hex')}-${cleaned}`;
}

interface SourceRef {
  /** Graph path to the item's metadata. */
  metaPath: string;
  /** Graph path to the item's raw bytes. */
  contentPath: string;
  label: string;
}

/** Resolves the caller's parameters into Graph paths, or explains why it can't. */
function resolveSource(params: {
  messageId?: string;
  attachmentId?: string;
  itemId?: string;
  driveId?: string;
  userId?: string;
}): SourceRef | { error: string } {
  const { messageId, attachmentId, itemId, driveId, userId } = params;
  const owner = userId?.trim();
  const mailBase = owner && owner !== 'me' ? `/users/${encodeURIComponent(owner)}` : '/me';

  if (messageId?.trim() && attachmentId?.trim()) {
    const m = encodeURIComponent(messageId.trim());
    const a = encodeURIComponent(attachmentId.trim());
    return {
      metaPath: `${mailBase}/messages/${m}/attachments/${a}?$select=id,name,contentType,size,isInline`,
      contentPath: `${mailBase}/messages/${m}/attachments/${a}/$value`,
      label: 'mail attachment',
    };
  }

  if (itemId?.trim()) {
    const base = driveId?.trim()
      ? `/drives/${encodeURIComponent(driveId.trim())}`
      : `${mailBase === '/me' ? '/me/drive' : `${mailBase}/drive`}`;
    const i = encodeURIComponent(itemId.trim());
    return {
      metaPath: `${base}/items/${i}?$select=id,name,size,file,folder`,
      contentPath: `${base}/items/${i}/content`,
      label: 'drive item',
    };
  }

  return {
    error:
      'Specify either messageId + attachmentId (for a mail attachment) or itemId (for a ' +
      'OneDrive / SharePoint file). Get these from list-mail-attachments, ' +
      'search-onedrive-files, or list-folder-files.',
  };
}

export interface RegisterHooks {
  isToolEnabled: (name: string) => boolean;
  readOnly: boolean;
  push: (name: string) => void;
  fail: (name: string, error: Error) => void;
  skip: (name: string) => void;
}

export function registerFileTools(
  server: McpServer,
  graphClient: GraphClient,
  hooks: RegisterHooks
): void {
  const { isToolEnabled, readOnly, push, fail, skip } = hooks;

  const register = (name: string, isWrite: boolean, registerFn: () => void) => {
    if (isWrite && readOnly) {
      skip(name);
      return;
    }
    if (!isToolEnabled(name)) return;
    try {
      registerFn();
      push(name);
    } catch (error) {
      fail(name, error as Error);
    }
  };

  // --- get-file --------------------------------------------------------------
  register('get-file', true, () => {
    server.tool(
      'get-file',
      'THE tool for getting at a file, whether it is a mail attachment or a OneDrive / ' +
        'SharePoint document. Use this instead of get-mail-attachment, ' +
        'download-mail-attachment, read-mail-attachment-text, download-onedrive-file-content ' +
        'or read-onedrive-file-text — it replaces all of them.\n\n' +
        'By default you get a short-lived pre-authenticated DOWNLOAD LINK, and you decide what ' +
        'to do with the file: fetch and parse it, search it, run code over it, or read it. The ' +
        'server does not interpret the file or assume what you wanted.\n\n' +
        'The `delivery` field tells you what came back:\n' +
        '  • `url` — a download link, valid about an hour, no auth header needed. The normal case.\n' +
        '  • `inline` — the file was tiny (under 32KB) and is included directly, because a link ' +
        'would have cost more than the content.\n' +
        '  • `text` — you asked for text with `as: "text"` and the document was readable.\n\n' +
        "Pass `as: 'text'` ONLY when you actually want the document's prose — a summary, a " +
        'question answered from it. Do not use it when you need the file itself, exact ' +
        'structure, or data you intend to compute over: extraction flattens layout, drops ' +
        'anything without a text layer, and is capped, so a value in a long document can be cut ' +
        'off without you knowing.\n\n' +
        'Identify the file by messageId + attachmentId, or by itemId (with driveId for a ' +
        'SharePoint library).',
      {
        messageId: z
          .string()
          .optional()
          .describe('For a mail attachment: the message that owns it.'),
        attachmentId: z
          .string()
          .optional()
          .describe('For a mail attachment: the attachment id from list-mail-attachments.'),
        itemId: z
          .string()
          .optional()
          .describe('For a OneDrive / SharePoint file: the driveItem id.'),
        driveId: z
          .string()
          .optional()
          .describe(
            "The drive holding the item. Omit for the signed-in user's own OneDrive; for a " +
              'SharePoint library get it from list-sharepoint-site-drives.'
          ),
        userId: z
          .string()
          .optional()
          .describe(
            "For a shared/other mailbox: the user id or UPN. Omit (or 'me') for the signed-in user."
          ),
        as: z
          .enum(['link', 'text'])
          .optional()
          .describe(
            "How you want the file. 'link' (default) returns a download link and leaves the " +
              "file intact for you to handle. 'text' extracts the document's prose server-side " +
              '— only ask for this when prose is genuinely what you want.'
          ),
        maxChars: z
          .number()
          .int()
          .min(500)
          .max(MAX_CHARS_LIMIT)
          .optional()
          .describe(
            `Only applies with as: 'text'. Cap on returned text; default ${DEFAULT_MAX_CHARS}, ` +
              `ceiling ${MAX_CHARS_LIMIT}.`
          ),
      },
      {
        title: 'get-file',
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: true,
      },
      async (params) =>
        withUsageLog('get-file', async () => {
          const source = resolveSource(params);
          if ('error' in source) return jsonResult({ error: source.error }, true);

          const meta = (await graphClient.makeRequest(source.metaPath)) as {
            '@odata.type'?: string;
            name?: string;
            contentType?: string;
            size?: number;
            file?: { mimeType?: string };
            folder?: unknown;
          };

          if (meta?.folder) {
            return jsonResult(
              {
                error: `'${meta.name}' is a folder. Use list-folder-files to see what is inside.`,
              },
              true
            );
          }

          const odataType = meta?.['@odata.type'] ?? '';
          if (odataType && !odataType.includes('fileAttachment')) {
            return jsonResult(
              {
                error:
                  `This is a ${odataType}, which has no file content. Use get-mail-attachment ` +
                  `to inspect it.`,
              },
              true
            );
          }

          const filename = meta?.name ?? 'file';
          const contentType = meta?.contentType ?? meta?.file?.mimeType ?? '';
          const size = meta?.size ?? 0;

          // --- INTERPRETATION: only when the caller explicitly asked for it ----
          if (params.as === 'text') {
            if (size > MAX_EXTRACT_BYTES) {
              return jsonResult(
                {
                  error:
                    `File is ${size} bytes, too large to extract text from (limit ` +
                    `${MAX_EXTRACT_BYTES}). Call again without as:'text' to get a download link.`,
                },
                true
              );
            }
            try {
              const { buffer } = await graphClient.fetchBinary(source.contentPath);
              const result = await extractText(
                buffer,
                filename,
                contentType,
                params.maxChars ?? DEFAULT_MAX_CHARS
              );
              logger.info(`get-file: delivering ${filename} as text (${result.totalChars} chars)`);
              return jsonResult({
                delivery: 'text',
                name: filename,
                size,
                format: result.format,
                truncated: result.truncated,
                totalChars: result.totalChars,
                ...(result.detail ?? {}),
                text: result.text,
                ...(result.truncated
                  ? {
                      note:
                        `Text truncated at ${result.text.length} of ${result.totalChars} ` +
                        `characters — the rest is NOT included. Call again with a larger ` +
                        `maxChars, or without as:'text' to get the whole file as a link.`,
                    }
                  : {}),
              });
            } catch (error) {
              if (!(error instanceof UnsupportedFormatError)) throw error;
              // Asked for text, but this file has none. Say so and give the link
              // rather than silently substituting a different answer.
              logger.info(
                `get-file: ${filename} has no text layer (${error.format}), falling back to link`
              );
              const url = await deliverUrl(graphClient, source, filename, contentType, size);
              return jsonResult({
                delivery: 'url',
                name: filename,
                size,
                ...url,
                note:
                  `You asked for text, but this file has none to extract (${error.format}). ` +
                  `Here is a download link instead.`,
              });
            }
          }

          // --- TRANSPORT: the only decision the server makes by default --------
          // Tiny files come back inline because a link, its metadata and a second
          // round trip would cost more than the content itself.
          if (size > 0 && size <= INLINE_MAX_BYTES) {
            const { buffer, contentType: fetched } = await graphClient.fetchBinary(
              source.contentPath
            );
            const effectiveType = contentType || fetched || '';
            // Deliberately strict. Substring matching is a trap here: a .docx
            // reports application/vnd.openxmlformats-…, which CONTAINS "xml"
            // but is a ZIP archive — decoding it as UTF-8 returns garbage.
            // Extension is the reliable signal; mime type only via exact match.
            const TEXT_TYPES = new Set([
              'application/json',
              'application/xml',
              'application/x-yaml',
              'application/yaml',
              'application/javascript',
            ]);
            const looksTextual =
              /\.(txt|csv|tsv|json|xml|md|log|ya?ml|ini|conf)$/i.test(filename) ||
              /^text\//i.test(effectiveType) ||
              TEXT_TYPES.has(effectiveType.split(';')[0].trim().toLowerCase());

            if (looksTextual) {
              logger.info(`get-file: delivering ${filename} inline (${buffer.byteLength}b)`);
              return jsonResult({
                delivery: 'inline',
                name: filename,
                size: buffer.byteLength,
                contentType: effectiveType,
                content: buffer.toString('utf8'),
                note:
                  'Small text file returned directly. This is the raw file content, not an ' +
                  'interpretation of it.',
              });
            }
            // Small but binary — a link is still the right answer; base64 through
            // the conversation helps nobody.
          }

          const url = await deliverUrl(graphClient, source, filename, contentType, size);
          return jsonResult({
            delivery: 'url',
            name: filename,
            size,
            contentType,
            ...url,
            note:
              'Download link, valid roughly one hour, no auth header needed. Fetch it and do ' +
              "whatever you need with the file. If you want the document's prose instead, call " +
              "again with as:'text'. Note some sandboxes cannot reach the tenant SharePoint " +
              'host — if the fetch is refused, that is an environment restriction, not a bad link.',
          });
        })
    );
  });

  // --- attach-file -----------------------------------------------------------
  register('attach-file', true, () => {
    server.tool(
      'attach-file',
      'Attaches a file to an existing draft message, choosing the correct upload route ' +
        'automatically. Small files are attached directly; larger ones go through a resumable ' +
        'upload session — Microsoft rejects direct attachment at 3 MB and above, which is why ' +
        'attaching used to fail on bigger files.\n\n' +
        'Source the file by itemId (a OneDrive / SharePoint document) or by messageId + ' +
        'attachmentId (to copy an attachment from another message). The bytes move ' +
        'server-side and never pass through the conversation.\n\n' +
        'Flow: create-draft-email (or create-reply-draft) → attach-file → send-draft-message.',
      {
        draftMessageId: z
          .string()
          .min(1)
          .describe('The draft to attach to, from create-draft-email or create-reply-draft.'),
        itemId: z.string().optional().describe('Source: a OneDrive / SharePoint driveItem id.'),
        driveId: z
          .string()
          .optional()
          .describe("Drive holding the source item. Omit for the signed-in user's OneDrive."),
        messageId: z
          .string()
          .optional()
          .describe('Source: the message holding an attachment to copy.'),
        attachmentId: z
          .string()
          .optional()
          .describe('Source: the attachment id to copy.'),
        name: z
          .string()
          .optional()
          .describe("Override the attachment filename. Defaults to the source file's name."),
      },
      {
        title: 'attach-file',
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: true,
      },
      async (params) =>
        withUsageLog('attach-file', async () => {
          const source = resolveSource(params);
          if ('error' in source) {
            return jsonResult(
              {
                error:
                  'Specify the source file: itemId (a OneDrive/SharePoint file) or ' +
                  'messageId + attachmentId (an attachment on another message).',
              },
              true
            );
          }

          const meta = (await graphClient.makeRequest(source.metaPath)) as {
            name?: string;
            size?: number;
            contentType?: string;
            file?: { mimeType?: string };
            folder?: unknown;
          };
          if (meta?.folder) {
            return jsonResult({ error: `'${meta.name}' is a folder, not a file.` }, true);
          }

          const filename = params.name?.trim() || meta?.name || 'attachment';
          const contentType =
            meta?.contentType ?? meta?.file?.mimeType ?? 'application/octet-stream';
          const draft = encodeURIComponent(params.draftMessageId.trim());

          const { buffer } = await graphClient.fetchBinary(source.contentPath);
          const size = buffer.byteLength;
          // Hash the bytes we actually send. Graph's reported attachment size
          // includes MIME envelope overhead (~165 bytes observed), so size is
          // NOT usable as an integrity check — a corrupt upload can report a
          // plausible size. Corruption was the originally reported symptom, so
          // the caller needs something it can actually verify against.
          const sha256 = crypto.createHash('sha256').update(buffer).digest('hex');

          if (size > MAX_SIMPLE_UPLOAD_BYTES) {
            return jsonResult(
              {
                error: `File is ${size} bytes, beyond the ${MAX_SIMPLE_UPLOAD_BYTES}-byte limit.`,
              },
              true
            );
          }

          // The routing decision this tool exists for.
          if (size < DIRECT_ATTACH_LIMIT) {
            await graphClient.makeRequest(`/me/messages/${draft}/attachments`, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                '@odata.type': '#microsoft.graph.fileAttachment',
                name: filename,
                contentType,
                contentBytes: buffer.toString('base64'),
              }),
            });
            logger.info(`attach-file: ${filename} attached directly (${size} bytes)`);
            return jsonResult({
              attached: true,
              route: 'direct',
              name: filename,
              size,
              sha256,
              draftMessageId: params.draftMessageId.trim(),
              note:
                'sha256 is of the bytes uploaded. Verify a downloaded copy against it — the ' +
                "attachment's reported size includes MIME overhead and will not match.",
            });
          }

          await graphClient.uploadViaSession(
            `/me/messages/${draft}/attachments/createUploadSession`,
            {
              AttachmentItem: {
                attachmentType: 'file',
                name: filename,
                size,
                contentType,
              },
            },
            buffer
          );
          logger.info(`attach-file: ${filename} attached via upload session (${size} bytes)`);
          return jsonResult({
            attached: true,
            route: 'uploadSession',
            name: filename,
            size,
            sha256,
            draftMessageId: params.draftMessageId.trim(),
            note:
              'sha256 is of the bytes uploaded. Verify a downloaded copy against it — the ' +
              "attachment's reported size includes MIME overhead and will not match.",
          });
        })
    );
  });
}

/**
 * Produces a pre-authenticated link for a file we could not deliver as text.
 *
 * Drive items already have one. Mail attachments do not, so a copy is staged
 * into the user's own OneDrive to borrow its download URL — the same approach
 * download-mail-attachment uses.
 */
async function deliverUrl(
  graphClient: GraphClient,
  source: SourceRef,
  filename: string,
  contentType: string,
  size: number
): Promise<Record<string, unknown>> {
  if (source.label === 'drive item') {
    const item = (await graphClient.makeRequest(
      source.metaPath.replace(/\?.*$/, '') + '?$select=id,@microsoft.graph.downloadUrl'
    )) as Record<string, unknown>;
    return {
      downloadUrl: item['@microsoft.graph.downloadUrl'],
      expiresInSeconds: 3600,
    };
  }

  if (size > MAX_SIMPLE_UPLOAD_BYTES) {
    throw new Error(`File is ${size} bytes, too large to stage for a download link.`);
  }

  const { buffer } = await graphClient.fetchBinary(source.contentPath);
  const sha256 = crypto.createHash('sha256').update(buffer).digest('hex');
  const uploadPath =
    `/me/drive/root:/${STAGING_FOLDER}/${encodeURIComponent(safeStagedName(filename))}:/content`;
  const staged = (await graphClient.putBinary(uploadPath, buffer, contentType)) as {
    id?: string;
    '@microsoft.graph.downloadUrl'?: string;
  };

  let downloadUrl = staged['@microsoft.graph.downloadUrl'];
  if (!downloadUrl && staged.id) {
    // The PUT response often omits the URL; fetch the staged item to get it.
    const fetched = (await graphClient.makeRequest(
      `/me/drive/items/${staged.id}?$select=id,@microsoft.graph.downloadUrl`
    )) as { '@microsoft.graph.downloadUrl'?: string };
    downloadUrl = fetched['@microsoft.graph.downloadUrl'];
  }

  return {
    downloadUrl,
    sha256,
    expiresInSeconds: 3600,
    stagedDriveItemId: staged.id,
  };
}
