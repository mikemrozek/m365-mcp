import logger from '../logger.js';
import type { VerifiedIdentity } from './verified-identity.js';

/**
 * Client for TSQ's Note Taker API v1 (Scott's service, published through Entra
 * Application Proxy in Passthrough mode). Built to Scott's 2026-10-01 spec;
 * documented in docs/2026-1005-1100 m365-mcp-notetaker-api-v1-client.md.
 *
 * Why the connector calls this as ITSELF, not as the user
 * -------------------------------------------------------
 * In production the connector never holds a refresh token and cannot ask
 * Microsoft for a second audience on the user's behalf. So it authenticates to
 * Note Taker with its own managed identity — a token minted for Note Taker's
 * audience and useless against Graph — and asserts the caller in ONE header,
 * `X-NoteTaker-User-Oid`. That object id is never taken from token claims or
 * from a tool argument: it comes from `verifyCallerIdentity`, which asks Graph
 * `/me` and fails closed.
 *
 * Note Taker is the authorization decision. It validates our bearer (v2
 * issuer, audience, app-only, our client id, role) and returns only meetings
 * the asserted person attended. "Not found" and "not yours" are deliberately
 * the same answer (`meeting_not_found`).
 *
 * Content discipline: nothing this module logs carries meeting content — only
 * HTTP status, Note Taker's error code, and its request id.
 */

/** Every refusal code Scott's API defines (2026-10-01), verbatim. */
export const NOTETAKER_API_CODES = [
  'invalid_token',
  'token_expired',
  'caller_not_allowed',
  'app_token_required',
  'missing_role',
  'user_identity_missing',
  'user_identity_invalid',
  'user_not_enabled',
  'meeting_not_found',
  'tenant_transcript_access_disabled',
  'api_disabled',
  'transcripts_disabled',
  'rate_limited_user',
  'rate_limited_connector',
  'daily_limit_user',
  'daily_limit_connector',
  'upstream_unavailable',
] as const;
export type NoteTakerApiCode = (typeof NOTETAKER_API_CODES)[number];

/** Failures that originate on our side of the wire, or that Note Taker did not name. */
export type NoteTakerLocalCode =
  /** MS365_MCP_NOTETAKER_API_URL / _AUDIENCE not set. */
  | 'not_configured'
  /** The managed identity could not mint a token for Note Taker's audience. */
  | 'token_unavailable'
  /** No answer within the timeout. */
  | 'timeout'
  /** Network failure, or a 5xx without one of Scott's codes (e.g. the App Proxy itself). */
  | 'unreachable'
  /** A 429 without one of Scott's codes. */
  | 'rate_limited'
  /** A response that is not in the contract: unknown code, or a 2xx we cannot read. */
  | 'unexpected_response';

export type NoteTakerCode = NoteTakerApiCode | NoteTakerLocalCode;

const API_CODES = new Set<string>(NOTETAKER_API_CODES);

export class NoteTakerError extends Error {
  constructor(
    message: string,
    readonly code: NoteTakerCode,
    readonly status?: number,
    readonly retryAfterSeconds?: number,
    /** Note Taker's request id, when it sent one. Safe to log; never content. */
    readonly requestId?: string,
    /** The raw code Note Taker sent when it was not one we know, sanitised. */
    readonly serverCode?: string
  ) {
    super(message);
    this.name = 'NoteTakerError';
  }
}

export interface NoteTakerConfig {
  /** External App Proxy URL, no trailing slash and no /api/v1 — the client adds that. */
  apiUrl: string;
  /** App ID URI of the Note Taker API registration — the token audience. */
  audience: string;
  /** Client id of the user-assigned managed identity to mint the token with. */
  identityClientId?: string;
  /** Per-request timeout, covering connect, headers and body. */
  timeoutMs: number;
}

export const DEFAULT_TIMEOUT_MS = 15_000;
const API_PREFIX = '/api/v1';

export function noteTakerConfigFromEnv(
  env: Record<string, string | undefined> = process.env
): NoteTakerConfig | undefined {
  const rawUrl = env.MS365_MCP_NOTETAKER_API_URL?.trim();
  const audience = env.MS365_MCP_NOTETAKER_AUDIENCE?.trim();
  if (!rawUrl || !audience) return undefined;
  const timeout = Number(env.MS365_MCP_NOTETAKER_TIMEOUT_MS);
  return {
    // Tolerate the URL being given with the prefix already on it, so a
    // container config written either way produces /api/v1/ exactly once.
    apiUrl: rawUrl.replace(/\/+$/, '').replace(/\/api\/v1$/i, ''),
    audience,
    identityClientId: env.MS365_MCP_NOTETAKER_IDENTITY_CLIENT_ID?.trim() || undefined,
    timeoutMs: Number.isFinite(timeout) && timeout > 0 ? timeout : DEFAULT_TIMEOUT_MS,
  };
}

/** The transcript tool has its own switch: Note Taker grants Transcripts.Read.All separately. */
export function transcriptsEnabledFromEnv(
  env: Record<string, string | undefined> = process.env
): boolean {
  const v = env.MS365_MCP_NOTETAKER_TRANSCRIPTS_ENABLED?.trim().toLowerCase();
  return v === '1' || v === 'true' || v === 'yes' || v === 'on';
}

/**
 * A meeting row. Field names follow Scott's prose (id, subject, start, end,
 * duration, organizer, participant count, conversation_id, has_transcript,
 * has_summary, attended); everything is passed through as received so a
 * renamed field degrades to "still there under its own name", not "dropped".
 */
export interface NoteTakerMeetingRow {
  id: string | number;
  [field: string]: unknown;
}

export interface NoteTakerMeetingList {
  meetings: NoteTakerMeetingRow[];
  /** Paging fields Note Taker returned alongside the rows (total, limit, offset, …). */
  paging: Record<string, unknown>;
  requestId?: string;
}

/** Row + summary (text, decisions, action items, key numbers) + participants — passed through. */
export interface NoteTakerMeetingDetail {
  meeting: Record<string, unknown>;
  requestId?: string;
}

export interface NoteTakerTranscriptPage {
  /** The page's transcript text. */
  text: string;
  /** Everything else Note Taker sent with the page (paging cursors, totals, meeting id). */
  meta: Record<string, unknown>;
  requestId?: string;
}

export interface ListMeetingsFilters {
  start_after?: string;
  start_before?: string;
  subject_contains?: string;
  organizer_email?: string;
  limit?: number;
  offset?: number;
}

export interface LookupRequest {
  conversation_id: string;
  start_after?: string;
  start_before?: string;
}

export type TokenSource = (audience: string) => Promise<string>;

interface ApiResponse {
  body: unknown;
  requestId?: string;
  status: number;
}

export class NoteTakerClient {
  private readonly tokenSource: TokenSource;
  private readonly fetchImpl: typeof fetch;

  constructor(
    private readonly config: NoteTakerConfig,
    options: { tokenSource?: TokenSource; fetchImpl?: typeof fetch } = {}
  ) {
    this.tokenSource = options.tokenSource ?? managedIdentityTokenSource(config.identityClientId);
    this.fetchImpl = options.fetchImpl ?? fetch;
  }

  get timeoutMs(): number {
    return this.config.timeoutMs;
  }

  /** GET /api/v1/status — our token only, no user header. For testing the door and the key. */
  async status(): Promise<{ body: unknown; requestId?: string }> {
    const r = await this.request('GET', '/status');
    return { body: r.body, requestId: r.requestId };
  }

  /** GET /api/v1/meetings — the caller's meetings, newest-first as Note Taker orders them. */
  async listMeetings(
    caller: VerifiedIdentity,
    filters: ListMeetingsFilters = {}
  ): Promise<NoteTakerMeetingList> {
    const query = new URLSearchParams();
    for (const [k, v] of Object.entries(filters)) {
      if (v !== undefined && v !== null && v !== '') query.set(k, String(v));
    }
    const qs = query.toString();
    const r = await this.request('GET', `/meetings${qs ? `?${qs}` : ''}`, caller);
    return { ...readRows(r), requestId: r.requestId };
  }

  /** POST /api/v1/meetings/lookup — that person's occurrences of a conversation, or []. */
  async lookupMeetings(
    caller: VerifiedIdentity,
    lookup: LookupRequest
  ): Promise<NoteTakerMeetingList> {
    const body: Record<string, string> = { conversation_id: lookup.conversation_id };
    if (lookup.start_after) body.start_after = lookup.start_after;
    if (lookup.start_before) body.start_before = lookup.start_before;
    const r = await this.request('POST', '/meetings/lookup', caller, body);
    return { ...readRows(r), requestId: r.requestId };
  }

  /** GET /api/v1/meetings/{id} — row plus summary and participants. */
  async getMeeting(caller: VerifiedIdentity, id: string): Promise<NoteTakerMeetingDetail> {
    const r = await this.request('GET', `/meetings/${encodeURIComponent(id)}`, caller);
    const meeting = r.body;
    if (!isRecord(meeting)) throw unexpected('a meeting detail that is not a JSON object', r);
    return { meeting, requestId: r.requestId };
  }

  /**
   * GET /api/v1/meetings/{id}/transcript — paged text. Behind Note Taker's
   * Transcripts.Read.All role, which is not assigned in v1; the tool calling
   * this is gated off by MS365_MCP_NOTETAKER_TRANSCRIPTS_ENABLED.
   */
  async getTranscript(
    caller: VerifiedIdentity,
    id: string,
    page?: string
  ): Promise<NoteTakerTranscriptPage> {
    const qs = page ? `?${new URLSearchParams({ page }).toString()}` : '';
    const r = await this.request(
      'GET',
      `/meetings/${encodeURIComponent(id)}/transcript${qs}`,
      caller
    );
    const page_ = r.body;
    if (!isRecord(page_)) throw unexpected('a transcript page that is not a JSON object', r);
    const textKey = ['text', 'content', 'transcript'].find((k) => typeof page_[k] === 'string');
    if (!textKey) throw unexpected('a transcript page without a text field', r);
    const { [textKey]: text, ...meta } = page_;
    return { text: text as string, meta, requestId: r.requestId };
  }

  private async request(
    method: 'GET' | 'POST',
    path: string,
    caller?: VerifiedIdentity,
    body?: unknown
  ): Promise<ApiResponse> {
    const token = await this.mintToken();

    const headers: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
    };
    // The ONLY identity Note Taker receives. A GUID from Graph /me, never a UPN.
    if (caller) headers['X-NoteTaker-User-Oid'] = caller.oid;
    if (body !== undefined) headers['Content-Type'] = 'application/json';

    const url = `${this.config.apiUrl}${API_PREFIX}${path}`;
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), this.config.timeoutMs);
    const route = `${method} ${API_PREFIX}${path.replace(/\/meetings\/[^/?]+/, '/meetings/{id}').replace(/\?.*$/, '')}`;

    try {
      let response: Response;
      try {
        response = await this.fetchImpl(url, {
          method,
          headers,
          body: body === undefined ? undefined : JSON.stringify(body),
          signal: controller.signal,
        });
      } catch (error) {
        if (controller.signal.aborted) throw this.timeoutError(route);
        throw new NoteTakerError(
          `Note Taker is unreachable: ${(error as Error).message}`,
          'unreachable'
        );
      }

      let text: string;
      try {
        text = await response.text();
      } catch (error) {
        if (controller.signal.aborted) throw this.timeoutError(route);
        throw new NoteTakerError(
          `Note Taker's response could not be read: ${(error as Error).message}`,
          'unreachable',
          response.status
        );
      }
      const parsed = parseJsonQuietly(text);
      const requestId = logSafeId(
        stringField(errorObject(parsed), 'request_id') ?? headerRequestId(response)
      );

      if (response.ok) {
        if (parsed === undefined) {
          throw unexpected('a non-JSON success response', {
            status: response.status,
            requestId,
            body: undefined,
          });
        }
        logger.info(
          `Note Taker ${route} -> ${response.status}${requestId ? ` request_id=${requestId}` : ''}`
        );
        return { body: parsed, requestId, status: response.status };
      }

      throw refusalFrom(response, parsed, requestId, route);
    } finally {
      clearTimeout(timer);
    }
  }

  private async mintToken(): Promise<string> {
    let timer: ReturnType<typeof setTimeout> | undefined;
    try {
      return await Promise.race([
        this.tokenSource(this.config.audience),
        new Promise<never>((_, reject) => {
          timer = setTimeout(
            () => reject(new Error(`no token within ${this.config.timeoutMs}ms`)),
            this.config.timeoutMs
          );
        }),
      ]);
    } catch (error) {
      throw new NoteTakerError(
        `Could not obtain a credential for Note Taker: ${(error as Error).message}`,
        'token_unavailable'
      );
    } finally {
      if (timer) clearTimeout(timer);
    }
  }

  private timeoutError(route: string): NoteTakerError {
    logger.warn(`Note Taker ${route} timed out after ${this.config.timeoutMs}ms`);
    return new NoteTakerError(
      `Note Taker did not answer within ${Math.round(this.config.timeoutMs / 1000)} seconds.`,
      'timeout'
    );
  }
}

function refusalFrom(
  response: Response,
  parsed: unknown,
  requestId: string | undefined,
  route: string
): NoteTakerError {
  const obj = errorObject(parsed);
  const rawCode =
    stringField(obj, 'code') ?? (typeof obj?.error === 'string' ? obj.error : undefined);
  const serverCode = rawCode
    ? rawCode
        .replace(/[^A-Za-z0-9_.:-]/g, '')
        .slice(0, 64)
        .toLowerCase() || undefined
    : undefined;
  const retryAfter = retryAfterSeconds(response, obj);
  const serverMessage = stringField(obj, 'message');
  const status = response.status;

  // Status, code and request id only — Note Taker's message may name a meeting.
  logger.warn(
    `Note Taker ${route} refused: HTTP ${status}${serverCode ? ` ${serverCode}` : ''}${requestId ? ` request_id=${requestId}` : ''}`
  );

  let code: NoteTakerCode;
  if (serverCode && API_CODES.has(serverCode)) code = serverCode as NoteTakerApiCode;
  else if (status === 429) code = 'rate_limited';
  else if (status >= 500) code = 'unreachable';
  else code = 'unexpected_response';

  const message =
    serverMessage?.slice(0, 500) ??
    `Note Taker returned HTTP ${status}${serverCode ? ` (${serverCode})` : ''}`;
  return new NoteTakerError(
    message,
    code,
    status,
    retryAfter,
    requestId,
    code === serverCode ? undefined : serverCode
  );
}

function unexpected(
  what: string,
  r: { status: number; requestId?: string; body: unknown }
): NoteTakerError {
  logger.warn(
    `Note Taker returned ${what} (HTTP ${r.status})${r.requestId ? ` request_id=${r.requestId}` : ''}`
  );
  return new NoteTakerError(
    `Note Taker returned ${what}; the connector and Note Taker disagree about the API.`,
    'unexpected_response',
    r.status,
    undefined,
    r.requestId
  );
}

/** Rows from a bare array or from the first array-valued wrapper key; other keys are paging. */
function readRows(r: ApiResponse): Omit<NoteTakerMeetingList, 'requestId'> {
  let rows: unknown;
  let paging: Record<string, unknown> = {};
  if (Array.isArray(r.body)) {
    rows = r.body;
  } else if (isRecord(r.body)) {
    const key = ['meetings', 'items', 'results', 'data', 'rows', 'value'].find((k) =>
      Array.isArray((r.body as Record<string, unknown>)[k])
    );
    if (key) {
      const { [key]: found, ...rest } = r.body;
      rows = found;
      delete rest.request_id;
      delete rest.requestId;
      paging = rest;
    }
  }
  if (!Array.isArray(rows)) throw unexpected('a meeting list without a list of meetings', r);
  const meetings = rows.filter(
    (row): row is NoteTakerMeetingRow =>
      isRecord(row) && ((typeof row.id === 'string' && row.id !== '') || typeof row.id === 'number')
  );
  if (meetings.length !== rows.length) {
    throw unexpected('meeting rows without an id', r);
  }
  return { meetings, paging };
}

function isRecord(v: unknown): v is Record<string, unknown> {
  return !!v && typeof v === 'object' && !Array.isArray(v);
}

function parseJsonQuietly(text: string): unknown {
  if (!text.trim()) return undefined;
  try {
    return JSON.parse(text);
  } catch {
    return undefined;
  }
}

/** The error fields may sit at the top level or under `error`. */
function errorObject(parsed: unknown): Record<string, unknown> | undefined {
  if (!isRecord(parsed)) return undefined;
  if (isRecord(parsed.error)) return { ...parsed, ...parsed.error };
  return parsed;
}

/** Case- and separator-insensitive field read: request_id, requestId, RequestId, request-id. */
function anyField(obj: Record<string, unknown> | undefined, name: string): unknown {
  if (!obj) return undefined;
  const want = name.replace(/[_-]/g, '').toLowerCase();
  for (const [k, v] of Object.entries(obj)) {
    if (k.replace(/[_-]/g, '').toLowerCase() === want) return v;
  }
  return undefined;
}

function stringField(obj: Record<string, unknown> | undefined, name: string): string | undefined {
  const v = anyField(obj, name);
  if (typeof v === 'number') return String(v);
  return typeof v === 'string' && v.trim() ? v.trim() : undefined;
}

function headerRequestId(response: Response): string | undefined {
  return (
    response.headers?.get?.('X-Request-Id') ?? response.headers?.get?.('Request-Id') ?? undefined
  );
}

/** Request ids are opaque but bounded to a log-safe token before they go anywhere. */
function logSafeId(raw: string | undefined): string | undefined {
  if (!raw) return undefined;
  return raw.replace(/[^A-Za-z0-9_.:-]/g, '').slice(0, 128) || undefined;
}

function retryAfterSeconds(response: Response, body?: Record<string, unknown>): number | undefined {
  const raw = response.headers?.get?.('Retry-After');
  if (raw) {
    const secs = Number(raw);
    if (Number.isFinite(secs) && secs > 0) return Math.ceil(secs);
    const at = Date.parse(raw);
    if (Number.isFinite(at)) {
      const delta = Math.ceil((at - Date.now()) / 1000);
      if (delta > 0) return delta;
    }
  }
  const fromBody = Number(anyField(body, 'retry_after_seconds'));
  return Number.isFinite(fromBody) && fromBody > 0 ? Math.ceil(fromBody) : undefined;
}

/**
 * Mints tokens with Azure's credential chain. In the container app that resolves
 * to the user-assigned managed identity (`tsqm365mcp-uami`); on a developer
 * machine it falls through to `az login`.
 */
function managedIdentityTokenSource(clientId?: string): TokenSource {
  let credential: { getToken(scopes: string): Promise<{ token: string } | null> } | undefined;
  return async (audience) => {
    if (!credential) {
      const { DefaultAzureCredential } = await import('@azure/identity');
      credential = new DefaultAzureCredential(
        clientId ? { managedIdentityClientId: clientId } : {}
      );
    }
    const scope = audience.endsWith('/.default') ? audience : `${audience}/.default`;
    const result = await credential.getToken(scope);
    if (!result?.token) throw new Error('credential chain returned no token');
    return result.token;
  };
}
