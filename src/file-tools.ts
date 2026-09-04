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

/**
 * Ceiling on bytes accepted INLINE on the write side, deliberately the same as the
 * read side's inline threshold.
 *
 * Scott's framing on 2026-09-02: "if it's 20K, who gives a ****" — small files
 * should just move through the conversation, and only larger ones need a link. The
 * symmetry is the point: a caller who knows get-file returns bytes under 32 KB can
 * assume put-file accepts bytes under 32 KB. Above it, base64 costs more context
 * than the file is worth: 1 MB of base64 is roughly 350,000 tokens, which would
 * exhaust a small plan's budget in a single call.
 */
const INLINE_UPLOAD_MAX_BYTES = 32 * 1024;

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

/** Strips characters OneDrive rejects, without the collision prefix staging adds. */
function safeUploadName(name: string): string {
  return (
    name
      .replace(/[\\/:*?"<>|]/g, '_')
      .replace(/^\.+/, '')
      .trim() || 'upload'
  );
}

/** Builds the `root:/folder/file:` addressing Graph uses for path-based drive writes. */
function drivePath(driveId: string | undefined, folder: string, name: string): string {
  const base = driveId?.trim() ? `/drives/${encodeURIComponent(driveId.trim())}` : '/me/drive';
  const segments = folder
    .split('/')
    .map((part) => part.trim())
    .filter(Boolean)
    .map((part) => encodeURIComponent(part));
  segments.push(encodeURIComponent(name));
  return `${base}/root:/${segments.join('/')}:`;
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
      // The opening sentence is written for a tool SEARCH, not for a reader. Claude
      // discovers capabilities by matching a query against names and descriptions,
      // and this tool lost that race for a month: `download-mail-attachment` carries
      // the words people type — download, mail, attachment — while `get-file` carries
      // none of them, so callers kept landing on the superseded tool even though its
      // own text says to prefer this one. Usage 25 Aug-1 Sep: get-file 11 calls, the
      // five it replaced 21. Hence the vocabulary below, deliberately front-loaded.
      'Download, read or open a file — an email attachment, or a document in OneDrive ' +
        'or SharePoint. Handles PDF, Word, Excel, PowerPoint, images, CSV and text.\n\n' +
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
          // `bytes` is the truth and is always present. Graph's own figure is
          // surfaced ONLY when it disagrees — for mail attachments it includes
          // MIME overhead (~165 bytes) and quietly differs from the real file,
          // which reads as corruption to anyone comparing sizes.
          const actualBytes = (url as { bytes?: number }).bytes;
          const graphDisagrees = typeof actualBytes === 'number' && actualBytes !== size;
          return jsonResult({
            delivery: 'url',
            name: filename,
            contentType,
            ...url,
            ...(graphDisagrees
              ? {
                  graphReportedSize: size,
                  sizeNote:
                    `Graph reports ${size} bytes, the file is ${actualBytes}. The difference is ` +
                    `MIME envelope overhead in Graph's metadata, not corruption. Verify against ` +
                    `bytes and sha256.`,
                }
              : {}),
            note:
              'Download link, valid roughly one hour, no auth header needed. QUOTE THE URL in ' +
              'any shell command — it contains & characters, and an unquoted & truncates the ' +
              'URL so the auth token never arrives (symptom: 401 generalException; fix the ' +
              'quoting, do not blame the link). Fetch it and do whatever you need with the ' +
              "file. If you want the document's prose instead, call again with as:'text'. A " +
              '403 with an x-deny-reason header means your environment blocks the tenant ' +
              'SharePoint host — an egress restriction, not a bad link.',
          });
        })
    );
  });

  // --- put-file --------------------------------------------------------------
  register('put-file', true, () => {
    server.tool(
      'put-file',
      // Written for tool search: upload/send/attach/save are the words a person
      // types, and the first sentence is what discovery matches on.
      'Upload or save a file into OneDrive or SharePoint — including a file that exists ' +
        'only on the local machine and is not yet in Microsoft 365. This is the write-side ' +
        'counterpart to get-file, and the missing step when someone asks to email a file ' +
        'they have locally.\n\n' +
        'Two routes, chosen the same way get-file chooses:\n' +
        '  • SMALL (under 32KB) — pass the bytes directly as contentBase64. Done in one call.\n' +
        '  • LARGER — omit contentBase64 and pass sizeBytes. You get back a short-lived ' +
        'uploadUrl to PUT the bytes to, so the file never passes through this conversation. ' +
        'The response to the final PUT contains the new item, whose id attach-file accepts.\n\n' +
        'To email the file: put-file → create-draft-email → attach-file → send-draft-message.',
      {
        name: z.string().min(1).describe('Filename including extension, e.g. "Q3 report.pdf".'),
        contentBase64: z
          .string()
          .optional()
          .describe(
            'The file itself, base64-encoded. Only for files under 32KB — above that omit ' +
              'this and pass sizeBytes to get an upload URL instead.'
          ),
        sizeBytes: z
          .number()
          .optional()
          .describe('Exact size in bytes. Required when contentBase64 is omitted.'),
        folderPath: z
          .string()
          .optional()
          .describe(
            `Destination folder, e.g. "Documents/Reports". Defaults to ${STAGING_FOLDER}, ` +
              'which is cleaned up automatically after 24 hours — pass a real folder for a ' +
              'file the user wants to keep.'
          ),
        driveId: z
          .string()
          .optional()
          .describe("Target a SharePoint library. Omit for the user's own OneDrive."),
        conflictBehavior: z
          .enum(['rename', 'replace', 'fail'])
          .optional()
          .describe('What to do if the name is taken. Default: rename.'),
      },
      {
        title: 'put-file',
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: true,
      },
      async (params) =>
        withUsageLog('put-file', async () => {
          const name = safeUploadName(params.name);
          const folder = params.folderPath?.trim() || STAGING_FOLDER;
          const conflict = params.conflictBehavior ?? 'rename';
          const target = drivePath(params.driveId, folder, name);

          if (params.contentBase64) {
            let buffer: Buffer;
            try {
              buffer = Buffer.from(params.contentBase64, 'base64');
            } catch {
              return jsonResult({ error: 'contentBase64 is not valid base64.' }, true);
            }
            if (buffer.byteLength === 0) {
              return jsonResult({ error: 'contentBase64 decoded to zero bytes.' }, true);
            }
            if (buffer.byteLength > INLINE_UPLOAD_MAX_BYTES) {
              return jsonResult(
                {
                  error:
                    `File is ${buffer.byteLength} bytes, over the ${INLINE_UPLOAD_MAX_BYTES}-byte ` +
                    'inline limit. Call again without contentBase64, passing sizeBytes, and PUT ' +
                    'the bytes to the uploadUrl you get back.',
                },
                true
              );
            }

            const item = (await graphClient.makeRequest(
              `${target}/content?@microsoft.graph.conflictBehavior=${conflict}`,
              {
                method: 'PUT',
                headers: { 'Content-Type': 'application/octet-stream' },
                body: buffer,
              }
            )) as { id?: string; name?: string; size?: number; webUrl?: string };

            const sha256 = crypto.createHash('sha256').update(buffer).digest('hex');
            logger.info(`put-file: ${name} uploaded inline (${buffer.byteLength} bytes)`);
            return jsonResult({
              uploaded: true,
              route: 'inline',
              itemId: item?.id,
              name: item?.name ?? name,
              bytes: buffer.byteLength,
              sha256,
              webUrl: item?.webUrl,
              folder,
              next: 'Pass itemId to attach-file to put this on a draft email.',
            });
          }

          const size = params.sizeBytes;
          if (typeof size !== 'number' || size <= 0) {
            return jsonResult(
              {
                error:
                  'Pass either contentBase64 (files under 32KB) or sizeBytes (anything larger, ' +
                  'to get an upload URL).',
              },
              true
            );
          }
          if (size > MAX_SIMPLE_UPLOAD_BYTES) {
            return jsonResult(
              { error: `File is ${size} bytes, beyond the ${MAX_SIMPLE_UPLOAD_BYTES}-byte limit.` },
              true
            );
          }

          const session = (await graphClient.makeRequest(`${target}/createUploadSession`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              item: { '@microsoft.graph.conflictBehavior': conflict, name },
            }),
          })) as { uploadUrl?: string; expirationDateTime?: string };

          if (!session?.uploadUrl) {
            return jsonResult({ error: 'Microsoft did not return an upload URL.' }, true);
          }

          logger.info(`put-file: upload session opened for ${name} (${size} bytes)`);
          return jsonResult({
            uploaded: false,
            route: 'uploadSession',
            uploadUrl: session.uploadUrl,
            expiresAt: session.expirationDateTime,
            name,
            bytes: size,
            folder,
            howTo:
              'PUT the bytes to uploadUrl with NO Authorization header — the URL carries its ' +
              'own. QUOTE THE URL in any shell command: it contains & and single quotes, and ' +
              'an unquoted & truncates the URL so the auth token never arrives. ' +
              `Send Content-Range: bytes 0-${Math.max(size - 1, 0)}/${size} for a single ` +
              'shot; above ~60MB send sequential chunks that are multiples of 320KB. The final ' +
              'response body is the created item — take its id and pass that to attach-file.',
            note:
              'The URL is short-lived and single-session, and the bytes go straight to Microsoft ' +
              'rather than through this conversation. Reading the error: 403 with an ' +
              'x-deny-reason header means your environment blocks the upload host — an egress ' +
              'restriction, not a bad URL. 401 generalException means the auth token never ' +
              'reached Microsoft — almost always an unquoted URL split at the first &; fix the ' +
              'quoting and mint a fresh session before blaming the environment.',
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
        attachmentId: z.string().optional().describe('Source: the attachment id to copy.'),
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
 * Deletes staged copies older than the retention window.
 *
 * Staging writes a second copy of the file into the user's own OneDrive so a
 * mail attachment can borrow a download URL. Nothing ever removed them, so the
 * folder grew without bound — confirmed in production 2026-08-17, where the same
 * 5.5 MB file appeared twice from two separate calls. That is both a quota
 * problem and a data-exposure one: copies of potentially sensitive attachments
 * accumulate outside the mailbox, under different retention and DLP treatment.
 *
 * Links expire in about an hour, so anything past a day is certainly dead.
 * Best-effort by design: pruning must never fail the caller's actual request,
 * so every error here is swallowed and logged.
 */
const STAGING_RETENTION_MS = 24 * 60 * 60 * 1000;

async function pruneStagingFolder(graphClient: GraphClient): Promise<void> {
  try {
    const listing = (await graphClient.makeRequest(
      `/me/drive/root:/${STAGING_FOLDER}:/children?$select=id,name,createdDateTime&$top=200`
    )) as { value?: Array<{ id?: string; name?: string; createdDateTime?: string }> };

    const cutoff = Date.now() - STAGING_RETENTION_MS;
    const stale = (listing?.value ?? []).filter((item) => {
      const created = Date.parse(item.createdDateTime ?? '');
      return Number.isFinite(created) && created < cutoff && item.id;
    });

    for (const item of stale) {
      try {
        await graphClient.makeRequest(`/me/drive/items/${item.id}`, { method: 'DELETE' });
      } catch (error) {
        logger.warn(`Could not prune staged file ${item.name}: ${(error as Error).message}`);
      }
    }
    if (stale.length) {
      logger.info(`Pruned ${stale.length} staged file(s) older than 24h`);
    }
  } catch (error) {
    // Folder may not exist yet on first use — not worth surfacing.
    logger.info(`Staging prune skipped: ${(error as Error).message}`);
  }
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
    // Do NOT $select here. `@microsoft.graph.downloadUrl` is an OData annotation,
    // not a selectable property: adding it to $select suppresses it, so the call
    // succeeds and returns an item with no download URL at all. That produced a
    // response claiming delivery:'url' with no url in it — reported from
    // production 2026-08-17. Fetching the item plainly returns the annotation.
    const item = (await graphClient.makeRequest(source.metaPath.replace(/\?.*$/, ''))) as Record<
      string,
      unknown
    >;
    const downloadUrl = item['@microsoft.graph.downloadUrl'];
    if (typeof downloadUrl !== 'string' || !downloadUrl) {
      throw new Error(
        'Graph returned no download URL for this item. It may be a folder, a OneNote ' +
          'section, or a file type without downloadable content.'
      );
    }
    // Drive metadata size IS the real byte count (unlike mail attachments, whose
    // size includes MIME overhead), so it can be reported as `bytes` directly.
    return {
      downloadUrl,
      bytes: size,
      expiresInSeconds: 3600,
      // Deliberate asymmetry, stated rather than left to be discovered: this is a
      // pass-through link, so the server never handles the bytes and cannot hash
      // them. Staged mail attachments DO get a sha256 because staging streams
      // them through us anyway. Hashing here would mean downloading the whole
      // file purely to checksum it, which defeats the point of a link.
      integrity: 'none — pass-through link; server did not read the bytes',
    };
  }

  if (size > MAX_SIMPLE_UPLOAD_BYTES) {
    throw new Error(`File is ${size} bytes, too large to stage for a download link.`);
  }

  // Clear out expired copies before adding another. Deliberately before rather
  // than after, so a failure here cannot leave the caller without their file.
  await pruneStagingFolder(graphClient);

  const { buffer } = await graphClient.fetchBinary(source.contentPath);
  const sha256 = crypto.createHash('sha256').update(buffer).digest('hex');
  const uploadPath = `/me/drive/root:/${STAGING_FOLDER}/${encodeURIComponent(safeStagedName(filename))}:/content`;
  const staged = (await graphClient.putBinary(uploadPath, buffer, contentType)) as {
    id?: string;
    '@microsoft.graph.downloadUrl'?: string;
  };

  let downloadUrl = staged['@microsoft.graph.downloadUrl'];
  if (!downloadUrl && staged.id) {
    // The PUT response often omits the URL; fetch the staged item to get it.
    const fetched = (await graphClient.makeRequest(
      // No $select — the downloadUrl annotation is suppressed by it (same trap
      // that broke the drive-item branch).
      `/me/drive/items/${staged.id}`
    )) as { '@microsoft.graph.downloadUrl'?: string };
    downloadUrl = fetched['@microsoft.graph.downloadUrl'];
  }

  return {
    downloadUrl,
    sha256,
    // The byte count the hash was computed over. Graph's attachment metadata
    // reports a LARGER number because it includes MIME envelope overhead (~165
    // bytes observed), so the two must be distinguishable or a caller comparing
    // them concludes the transfer was corrupt when it was not.
    bytes: buffer.byteLength,
    integrity: 'sha256 over the bytes uploaded',
    expiresInSeconds: 3600,
    stagedDriveItemId: staged.id,
  };
}
