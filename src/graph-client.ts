import logger from './logger.js';
import AuthManager from './auth.js';
import { encode as toonEncode } from '@toon-format/toon';
import type { AppSecrets } from './secrets.js';
import { getCloudEndpoints } from './cloud-config.js';
import { getRequestTokens } from './request-context.js';

/**
 * Returns true if the given HTTP Content-Type header indicates a binary
 * payload that must not be decoded as UTF-8 text. Graph returns binary for
 * endpoints like /me/photo/$value, /chats/.../hostedContents/{id}/$value, and
 * /drives/.../items/{id}/content, among others.
 */
export function isBinaryContentType(contentType: string): boolean {
  if (!contentType) return false;
  const lower = contentType.toLowerCase().split(';')[0].trim();
  if (!lower) return false;
  if (
    lower.startsWith('image/') ||
    lower.startsWith('video/') ||
    lower.startsWith('audio/') ||
    lower.startsWith('font/')
  ) {
    return true;
  }
  if (lower === 'application/octet-stream' || lower === 'application/pdf') {
    return true;
  }
  if (lower.startsWith('application/zip') || lower.startsWith('application/x-zip')) {
    return true;
  }
  // Office document MIME types and other vendor-specific binary formats.
  if (lower.startsWith('application/vnd.') || lower.startsWith('application/x-')) {
    // Be conservative: exclude MIME types that use the structured-syntax suffix
    // to declare a text serialization (e.g. application/vnd.api+json).
    if (lower.endsWith('+json') || lower.endsWith('+xml') || lower.endsWith('+text')) {
      return false;
    }
    return true;
  }
  return false;
}

interface GraphRequestOptions {
  headers?: Record<string, string>;
  method?: string;
  body?: string;
  rawResponse?: boolean;
  includeHeaders?: boolean;
  excludeResponse?: boolean;
  accessToken?: string;

  [key: string]: unknown;
}

/**
 * Renders a request for logging WITHOUT its payload or credentials.
 *
 * This previously logged `JSON.stringify(options)` wholesale, which wrote every
 * outbound body to the log — `send-mail` logged the email, `add-mail-attachment`
 * logged base64 file content — and, in multi-account mode, the caller's access
 * token alongside it. It also silently defeated the redaction `executeGraphTool`
 * performs one layer up, which strips the token before its own log line.
 *
 * Diagnostic value lives in the method, the endpoint and the rough size, none of
 * which require the content. Anything genuinely needing the body should be read
 * from a debugger, not from a log file that outlives the request.
 *
 * SEC-2026-001, .claude/security-docs/assessments/2026-08-14-audit.md
 */
export function describeRequestForLog(options: GraphRequestOptions = {}): string {
  const parts = [`method=${(options.method || 'GET').toUpperCase()}`];
  if (options.body !== undefined) {
    const size =
      typeof options.body === 'string'
        ? Buffer.byteLength(options.body, 'utf8')
        : Buffer.byteLength(String(options.body), 'utf8');
    parts.push(`bodyBytes=${size}`);
  }
  // Header NAMES are safe and useful (e.g. confirming ConsistencyLevel was
  // applied); their values are not, so they are never rendered.
  const headerNames = Object.keys(options.headers ?? {});
  if (headerNames.length) parts.push(`headers=[${headerNames.join(',')}]`);
  if (options.accessToken) parts.push('accessToken=[REDACTED]');
  if (options.rawResponse) parts.push('rawResponse=true');
  if (options.excludeResponse) parts.push('excludeResponse=true');
  return parts.join(' ');
}

interface ContentItem {
  type: 'text';
  text: string;

  [key: string]: unknown;
}

interface McpResponse {
  content: ContentItem[];
  _meta?: Record<string, unknown>;
  isError?: boolean;

  [key: string]: unknown;
}

class GraphClient {
  private authManager: AuthManager;
  private secrets: AppSecrets;
  private readonly outputFormat: 'json' | 'toon' = 'json';

  constructor(
    authManager: AuthManager,
    secrets: AppSecrets,
    outputFormat: 'json' | 'toon' = 'json'
  ) {
    this.authManager = authManager;
    this.secrets = secrets;
    this.outputFormat = outputFormat;
  }

  async makeRequest(endpoint: string, options: GraphRequestOptions = {}): Promise<unknown> {
    const contextTokens = getRequestTokens();
    const accessToken =
      options.accessToken ?? contextTokens?.accessToken ?? (await this.authManager.getToken());

    if (!accessToken) {
      throw new Error('No access token available');
    }

    try {
      const response = await this.performRequest(endpoint, accessToken, options);

      if (response.status === 403) {
        const errorText = await response.text();
        if (errorText.includes('scope') || errorText.includes('permission')) {
          throw new Error(
            `Microsoft Graph API scope error: ${response.status} ${response.statusText} - ${errorText}. This tool requires organization mode. Please restart with --org-mode flag.`
          );
        }
        throw new Error(
          `Microsoft Graph API error: ${response.status} ${response.statusText} - ${errorText}`
        );
      }

      if (!response.ok) {
        throw new Error(
          `Microsoft Graph API error: ${response.status} ${response.statusText} - ${await response.text()}`
        );
      }

      const contentTypeHeader = response.headers?.get?.('content-type') || '';
      const isBinaryResponse = isBinaryContentType(contentTypeHeader);

      let result: any;

      if (isBinaryResponse) {
        // Binary payloads (images, video, pdf, octet-stream, etc.) must not be
        // decoded with response.text() — that performs a lossy UTF-8 decode and
        // replaces every high byte with U+FFFD, destroying the file. Read the
        // raw bytes and return them as base64 so callers can reconstruct them.
        const buffer = Buffer.from(await response.arrayBuffer());
        result = {
          message: 'OK!',
          contentType: contentTypeHeader,
          encoding: 'base64',
          contentLength: buffer.byteLength,
          contentBytes: buffer.toString('base64'),
        };
      } else {
        const text = await response.text();

        if (text === '') {
          result = { message: 'OK!' };
        } else {
          try {
            result = JSON.parse(text);
          } catch {
            result = { message: 'OK!', rawResponse: text };
          }
        }
      }

      // If includeHeaders is requested, add response headers to the result
      if (options.includeHeaders) {
        const etag = response.headers.get('ETag') || response.headers.get('etag');

        // Simple approach: just add ETag to the result if it's an object
        if (result && typeof result === 'object' && !Array.isArray(result)) {
          return {
            ...result,
            _etag: etag || 'no-etag-found',
          };
        }
      }

      return result;
    } catch (error) {
      logger.error('Microsoft Graph API request failed:', error);
      throw error;
    }
  }

  private async performRequest(
    endpoint: string,
    accessToken: string,
    options: GraphRequestOptions
  ): Promise<Response> {
    const cloudEndpoints = getCloudEndpoints(this.secrets.cloudType);
    const url = `${cloudEndpoints.graphApi}/v1.0${endpoint}`;

    logger.info(`[GRAPH CLIENT] Final URL being sent to Microsoft: ${url}`);

    const headers: Record<string, string> = {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...options.headers,
    };

    return fetch(url, {
      method: options.method || 'GET',
      headers,
      body: options.body,
    });
  }

  private serializeData(data: unknown, outputFormat: 'json' | 'toon', pretty = false): string {
    if (outputFormat === 'toon') {
      try {
        return toonEncode(data);
      } catch (error) {
        logger.warn(`Failed to encode as TOON, falling back to JSON: ${error}`);
        return JSON.stringify(data, null, pretty ? 2 : undefined);
      }
    }
    return JSON.stringify(data, null, pretty ? 2 : undefined);
  }

  /**
   * Posts a Graph /$batch request with up to 20 sub-requests. Token resolution,
   * URL construction, and error handling come from makeRequest. Returns the raw
   * { responses: [...] } envelope from Graph so callers can correlate sub-responses
   * back to the original requests by id.
   */
  async batchRequest(
    subRequests: Array<{
      id: string;
      method: string;
      url: string;
      headers?: Record<string, string>;
      body?: unknown;
    }>
  ): Promise<{
    responses: Array<{
      id: string;
      status: number;
      headers?: Record<string, string>;
      body: unknown;
    }>;
  }> {
    logger.info(`[GRAPH CLIENT] Batch request with ${subRequests.length} sub-requests`);
    const result = await this.makeRequest('/$batch', {
      method: 'POST',
      body: JSON.stringify({ requests: subRequests }),
    });
    return result as {
      responses: Array<{
        id: string;
        status: number;
        headers?: Record<string, string>;
        body: unknown;
      }>;
    };
  }

  /**
   * Fetches raw bytes from a Graph endpoint (e.g. an attachment's /$value).
   * Unlike makeRequest, this never base64-encodes or UTF-8-decodes the body —
   * it returns the Buffer verbatim plus the content type, so callers can
   * re-stream the bytes (e.g. stage to OneDrive) without a lossy round-trip.
   */
  async fetchBinary(
    endpoint: string,
    options: GraphRequestOptions = {}
  ): Promise<{ buffer: Buffer; contentType: string }> {
    const contextTokens = getRequestTokens();
    const accessToken =
      options.accessToken ?? contextTokens?.accessToken ?? (await this.authManager.getToken());
    if (!accessToken) {
      throw new Error('No access token available');
    }
    const response = await this.performRequest(endpoint, accessToken, options);
    if (!response.ok) {
      throw new Error(
        `Microsoft Graph API error: ${response.status} ${response.statusText} - ${await response.text()}`
      );
    }
    const contentType = response.headers?.get?.('content-type') || 'application/octet-stream';
    const buffer = Buffer.from(await response.arrayBuffer());
    logger.info(`[GRAPH CLIENT] Fetched ${buffer.byteLength} binary bytes from ${endpoint}`);
    return { buffer, contentType };
  }

  /**
   * Uploads raw bytes to a Graph endpoint via PUT (e.g. a simple OneDrive
   * upload to .../content, which supports files up to 250 MB in one request).
   * Sends the Buffer directly with the given content type rather than a JSON
   * body, and returns the parsed driveItem JSON Graph responds with.
   */
  async putBinary(
    endpoint: string,
    buffer: Buffer,
    contentType: string,
    options: GraphRequestOptions = {}
  ): Promise<unknown> {
    const contextTokens = getRequestTokens();
    const accessToken =
      options.accessToken ?? contextTokens?.accessToken ?? (await this.authManager.getToken());
    if (!accessToken) {
      throw new Error('No access token available');
    }
    const cloudEndpoints = getCloudEndpoints(this.secrets.cloudType);
    const url = `${cloudEndpoints.graphApi}/v1.0${endpoint}`;
    logger.info(`[GRAPH CLIENT] Binary PUT to: ${url} (${buffer.byteLength} bytes)`);
    const response = await fetch(url, {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': contentType || 'application/octet-stream',
      },
      // Zero-copy Uint8Array view. The DOM lib's BodyInit type doesn't list
      // typed arrays under this TS config, but undici's fetch accepts them at
      // runtime — cast to satisfy the overload.
      body: new Uint8Array(
        buffer.buffer,
        buffer.byteOffset,
        buffer.byteLength
      ) as unknown as BodyInit,
    });
    if (!response.ok) {
      throw new Error(
        `Microsoft Graph API error: ${response.status} ${response.statusText} - ${await response.text()}`
      );
    }
    const text = await response.text();
    return text ? JSON.parse(text) : {};
  }

  /**
   * Uploads bytes through a Graph upload session, which is the only supported
   * route for mail attachments of 3 MB or more and the recommended one for
   * large drive items.
   *
   * Three rules from Microsoft's docs are load-bearing here and easy to get
   * wrong:
   *   - The `uploadUrl` is pre-authenticated and opaque. Sending an
   *     `Authorization` header on the PUT can fail with 401, so we deliberately
   *     send none. It is also NOT on the tenant host — mail attachment sessions
   *     live on outlook.office.com and drive sessions on a regional
   *     *.up.1drv.com host — which is why this upload runs server-side rather
   *     than from the client sandbox.
   *   - Every chunk except the last must be a multiple of 320 KiB, or large
   *     transfers fail when the final range is committed.
   *   - Ranges must be sent sequentially; out-of-order writes are rejected.
   *
   * @param createSessionEndpoint Graph path that mints the session.
   * @param body Request body for the session creation call.
   * @returns The final response Graph returns when the last chunk commits.
   */
  async uploadViaSession(
    createSessionEndpoint: string,
    body: Record<string, unknown>,
    buffer: Buffer,
    options: GraphRequestOptions = {}
  ): Promise<unknown> {
    const session = (await this.makeRequest(createSessionEndpoint, {
      ...options,
      method: 'POST',
      headers: { 'Content-Type': 'application/json', ...(options.headers ?? {}) },
      body: JSON.stringify(body),
    })) as { uploadUrl?: string };

    if (!session?.uploadUrl) {
      throw new Error(
        `Graph did not return an uploadUrl for ${createSessionEndpoint}: ${JSON.stringify(session)}`
      );
    }

    const total = buffer.byteLength;
    // 320 KiB is the required alignment; 10 MiB is Microsoft's recommended
    // chunk for stable connections and stays well under the 60 MiB ceiling.
    const CHUNK = 10 * 1024 * 1024;
    let offset = 0;
    let last: unknown = {};

    while (offset < total) {
      const end = Math.min(offset + CHUNK, total);
      const chunk = buffer.subarray(offset, end);
      const response = await fetch(session.uploadUrl, {
        method: 'PUT',
        headers: {
          // No Authorization header — the URL already carries its own auth.
          'Content-Length': String(chunk.byteLength),
          'Content-Range': `bytes ${offset}-${end - 1}/${total}`,
        },
        body: new Uint8Array(
          chunk.buffer,
          chunk.byteOffset,
          chunk.byteLength
        ) as unknown as BodyInit,
      });

      if (!response.ok) {
        // Best effort cleanup so an abandoned session doesn't linger.
        await fetch(session.uploadUrl, { method: 'DELETE' }).catch(() => {});
        throw new Error(
          `Upload session failed at bytes ${offset}-${end - 1}/${total}: ` +
            `${response.status} ${response.statusText} - ${await response.text()}`
        );
      }

      const text = await response.text();
      last = text ? JSON.parse(text) : {};
      offset = end;
    }

    logger.info(`[GRAPH CLIENT] Upload session complete: ${total} bytes`);
    return last;
  }

  async graphRequest(endpoint: string, options: GraphRequestOptions = {}): Promise<McpResponse> {
    try {
      logger.info(`Calling ${endpoint} ${describeRequestForLog(options)}`);

      // Use new OAuth-aware request method
      const result = await this.makeRequest(endpoint, options);

      return this.formatJsonResponse(result, options.rawResponse, options.excludeResponse);
    } catch (error) {
      logger.error(`Error in Graph API request: ${error}`);
      return {
        content: [{ type: 'text', text: JSON.stringify({ error: (error as Error).message }) }],
        isError: true,
      };
    }
  }

  formatJsonResponse(data: unknown, rawResponse = false, excludeResponse = false): McpResponse {
    // If excludeResponse is true, only return success indication
    if (excludeResponse) {
      return {
        content: [{ type: 'text', text: this.serializeData({ success: true }, this.outputFormat) }],
      };
    }

    // Handle the case where data includes headers metadata
    if (data && typeof data === 'object' && '_headers' in data) {
      const responseData = data as {
        data: unknown;
        _headers: Record<string, string>;
        _etag?: string;
      };

      const meta: Record<string, unknown> = {};
      if (responseData._etag) {
        meta.etag = responseData._etag;
      }
      if (responseData._headers) {
        meta.headers = responseData._headers;
      }

      if (rawResponse) {
        return {
          content: [
            { type: 'text', text: this.serializeData(responseData.data, this.outputFormat) },
          ],
          _meta: meta,
        };
      }

      if (responseData.data === null || responseData.data === undefined) {
        return {
          content: [
            { type: 'text', text: this.serializeData({ success: true }, this.outputFormat) },
          ],
          _meta: meta,
        };
      }

      // Remove OData properties
      const removeODataProps = (obj: Record<string, unknown>): void => {
        if (typeof obj === 'object' && obj !== null) {
          Object.keys(obj).forEach((key) => {
            if (
              key.startsWith('@odata.') &&
              key !== '@odata.nextLink' &&
              key !== '@odata.deltaLink'
            ) {
              delete obj[key];
            } else if (typeof obj[key] === 'object') {
              removeODataProps(obj[key] as Record<string, unknown>);
            }
          });
        }
      };

      removeODataProps(responseData.data as Record<string, unknown>);

      return {
        content: [
          { type: 'text', text: this.serializeData(responseData.data, this.outputFormat, true) },
        ],
        _meta: meta,
      };
    }

    // Original handling for backward compatibility
    if (rawResponse) {
      return {
        content: [{ type: 'text', text: this.serializeData(data, this.outputFormat) }],
      };
    }

    if (data === null || data === undefined) {
      return {
        content: [{ type: 'text', text: this.serializeData({ success: true }, this.outputFormat) }],
      };
    }

    // Remove OData properties
    const removeODataProps = (obj: Record<string, unknown>): void => {
      if (typeof obj === 'object' && obj !== null) {
        Object.keys(obj).forEach((key) => {
          if (
            key.startsWith('@odata.') &&
            key !== '@odata.nextLink' &&
            key !== '@odata.deltaLink'
          ) {
            delete obj[key];
          } else if (typeof obj[key] === 'object') {
            removeODataProps(obj[key] as Record<string, unknown>);
          }
        });
      }
    };

    removeODataProps(data as Record<string, unknown>);

    return {
      content: [{ type: 'text', text: this.serializeData(data, this.outputFormat, true) }],
    };
  }
}

export default GraphClient;
