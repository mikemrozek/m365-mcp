import { z } from 'zod';
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
 * Tools that return a document's *text* rather than its bytes or a link to it.
 *
 * Why: `download-mail-attachment` stages a file to OneDrive and returns a
 * pre-authed URL, which works for a human clicking it but not for Claude —
 * Claude's sandbox has no egress to the tenant SharePoint host, so fetching
 * that URL fails (`403 host_not_allowed`, verified 2026-08-10). Extraction runs
 * here, server-side, so no egress is involved. It is also cheaper in context
 * than the base64 fallback, since prose compresses better than an encoded
 * binary, and it answers the question people actually ask ("what does this
 * invoice say") instead of handing back a file.
 */

/** Refuse before downloading anything larger than this. */
const MAX_SOURCE_BYTES = 40 * 1024 * 1024;

function jsonResult(payload: unknown, isError = false) {
  return {
    content: [{ type: 'text' as const, text: JSON.stringify(payload) }],
    ...(isError ? { isError: true } : {}),
  };
}

const maxCharsParam = z
  .number()
  .int()
  .min(500)
  .max(MAX_CHARS_LIMIT)
  .optional()
  .describe(
    `Maximum characters of text to return. Default ${DEFAULT_MAX_CHARS}, hard ceiling ` +
      `${MAX_CHARS_LIMIT}. Raise it only when you genuinely need the whole document — large ` +
      `values consume the conversation's working room.`
  );

const SHARED_DESCRIPTION =
  'SUPERSEDED by get-file, which covers both mail attachments and documents and also falls '  +
  'back to a link when a file cannot be read as text. Prefer get-file unless you specifically '  +
  'want extraction to fail loudly on an unreadable file.\n\n' +
  'Returns the extracted TEXT, not the file and not a download link, so it can be read and ' +
  'reasoned over directly. Supports pdf, docx, xlsx, pptx, and plain-text formats (csv, json, ' +
  'xml, html, md, txt). Spreadsheets come back as tab-separated rows per sheet.\n\n' +
  'Prefer this over the download tools whenever the goal is to READ or SUMMARISE a document. ' +
  'Use download-mail-attachment / download-onedrive-file-content only when the actual file is ' +
  'needed, or when this reports the format as unsupported.\n\n' +
  'Scanned documents and images have no text layer and cannot be read here — that is reported ' +
  'explicitly rather than returned as empty text. Output is capped; when `truncated` is true, ' +
  '`totalChars` tells you how much text the document really had.';

export interface RegisterHooks {
  isToolEnabled: (name: string) => boolean;
  push: (name: string) => void;
  fail: (name: string, error: Error) => void;
}

export function registerFileTextTools(
  server: McpServer,
  graphClient: GraphClient,
  hooks: RegisterHooks
): void {
  const { isToolEnabled, push, fail } = hooks;

  const register = (name: string, registerFn: () => void) => {
    if (!isToolEnabled(name)) return;
    try {
      registerFn();
      push(name);
    } catch (error) {
      fail(name, error as Error);
    }
  };

  /** Shared tail: extract, shape the response, turn format problems into guidance. */
  async function respondWithText(
    buffer: Buffer,
    filename: string,
    contentType: string,
    maxChars: number | undefined,
    extra: Record<string, unknown> = {}
  ) {
    try {
      const result = await extractText(buffer, filename, contentType, maxChars ?? DEFAULT_MAX_CHARS);
      return jsonResult({
        ...extra,
        name: filename,
        format: result.format,
        truncated: result.truncated,
        totalChars: result.totalChars,
        ...(result.detail ?? {}),
        text: result.text,
        ...(result.truncated
          ? {
              note: `Text was truncated at ${result.text.length} of ${result.totalChars} characters. ` +
                `Call again with a larger maxChars if you need the rest.`,
            }
          : {}),
      });
    } catch (error) {
      if (error instanceof UnsupportedFormatError) {
        // Not a failure of the tool — a fact about the file. Say which format
        // and what to do instead, so the model stops rather than retrying.
        return jsonResult(
          { name: filename, format: error.format, extractable: false, reason: error.message },
          true
        );
      }
      throw error;
    }
  }

  // --- read-mail-attachment-text -------------------------------------------
  register('read-mail-attachment-text', () => {
    server.tool(
      'read-mail-attachment-text',
      'Reads the text of a mail attachment. ' + SHARED_DESCRIPTION,
      {
        messageId: z.string().min(1).describe('The message that owns the attachment.'),
        attachmentId: z
          .string()
          .min(1)
          .describe('The attachment to read. Get it from list-mail-attachments.'),
        userId: z
          .string()
          .optional()
          .describe(
            "For shared/other mailboxes: the user id or UPN whose message this is. Omit (or " +
              "'me') for the signed-in user's own mailbox."
          ),
        maxChars: maxCharsParam,
      },
      {
        title: 'read-mail-attachment-text',
        readOnlyHint: true,
        openWorldHint: true,
      },
      async (params) =>
        withUsageLog('read-mail-attachment-text', async () => {
          const messageId = params.messageId.trim();
          const attachmentId = params.attachmentId.trim();
          const owner = params.userId?.trim();
          const base =
            owner && owner !== 'me'
              ? `/users/${encodeURIComponent(owner)}`
              : '/me';

          // Metadata first: this lets us reject an oversized or non-file
          // attachment before pulling a single byte.
          const meta = (await graphClient.makeRequest(
            `${base}/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(
              attachmentId
            )}?$select=id,name,contentType,size,isInline`
          )) as {
            '@odata.type'?: string;
            name?: string;
            contentType?: string;
            size?: number;
          };

          const odataType = meta?.['@odata.type'] ?? '';
          if (odataType && !odataType.includes('fileAttachment')) {
            return jsonResult(
              {
                extractable: false,
                reason:
                  `This is a ${odataType} attachment, not a file. Use get-mail-attachment to ` +
                  `inspect it.`,
              },
              true
            );
          }
          if ((meta?.size ?? 0) > MAX_SOURCE_BYTES) {
            return jsonResult(
              {
                extractable: false,
                reason:
                  `Attachment is ${meta.size} bytes, over the ${MAX_SOURCE_BYTES}-byte limit for ` +
                  `text extraction. Use download-mail-attachment instead.`,
              },
              true
            );
          }

          const filename = meta?.name ?? 'attachment';
          const { buffer } = await graphClient.fetchBinary(
            `${base}/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(
              attachmentId
            )}/$value`
          );
          logger.info(`Extracting text from mail attachment ${filename} (${buffer.byteLength}b)`);
          return respondWithText(buffer, filename, meta?.contentType ?? '', params.maxChars, {
            size: buffer.byteLength,
          });
        })
    );
  });

  // --- read-onedrive-file-text ---------------------------------------------
  register('read-onedrive-file-text', () => {
    server.tool(
      'read-onedrive-file-text',
      'Reads the text of a file stored in OneDrive or a SharePoint document library. ' +
        SHARED_DESCRIPTION,
      {
        itemId: z
          .string()
          .min(1)
          .describe('The driveItem id. Get it from search-onedrive-files or list-folder-files.'),
        driveId: z
          .string()
          .optional()
          .describe(
            "The drive containing the item. Omit for the signed-in user's own OneDrive. For a " +
              'SharePoint document library, get it from list-sharepoint-site-drives.'
          ),
        maxChars: maxCharsParam,
      },
      {
        title: 'read-onedrive-file-text',
        readOnlyHint: true,
        openWorldHint: true,
      },
      async (params) =>
        withUsageLog('read-onedrive-file-text', async () => {
          const itemId = params.itemId.trim();
          const driveId = params.driveId?.trim();
          const base = driveId ? `/drives/${encodeURIComponent(driveId)}` : '/me/drive';
          const itemPath = `${base}/items/${encodeURIComponent(itemId)}`;

          const meta = (await graphClient.makeRequest(
            `${itemPath}?$select=id,name,size,file,folder`
          )) as {
            name?: string;
            size?: number;
            file?: { mimeType?: string };
            folder?: unknown;
          };

          if (meta?.folder) {
            return jsonResult(
              {
                extractable: false,
                reason: `'${meta.name}' is a folder, not a file. Use list-folder-files to see what is inside.`,
              },
              true
            );
          }
          if ((meta?.size ?? 0) > MAX_SOURCE_BYTES) {
            return jsonResult(
              {
                extractable: false,
                reason:
                  `File is ${meta.size} bytes, over the ${MAX_SOURCE_BYTES}-byte limit for text ` +
                  `extraction. Use download-onedrive-file-content instead.`,
              },
              true
            );
          }

          const filename = meta?.name ?? 'file';
          const { buffer } = await graphClient.fetchBinary(`${itemPath}/content`);
          logger.info(`Extracting text from drive item ${filename} (${buffer.byteLength}b)`);
          return respondWithText(
            buffer,
            filename,
            meta?.file?.mimeType ?? '',
            params.maxChars,
            { size: buffer.byteLength }
          );
        })
    );
  });
}
