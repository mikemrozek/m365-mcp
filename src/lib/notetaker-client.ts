import logger from '../logger.js';
import type { VerifiedIdentity } from './verified-identity.js';

/**
 * Client for TSQ's Note Taker transcript API (Scott's service, published through
 * Entra Application Proxy in Passthrough mode).
 *
 * Why the connector calls this as ITSELF, not as the user
 * -------------------------------------------------------
 * In production the connector never holds a refresh token and cannot ask
 * Microsoft for a second audience on the user's behalf. So it authenticates to
 * Note Taker with its own managed identity — a token minted for Note Taker's
 * audience and useless against Graph — and asserts the caller's identity in
 * headers. That identity is not taken from token claims: it comes from
 * `verifyCallerIdentity`, which asks Graph `/me` and fails closed.
 *
 * The consequence, stated plainly because it is the crux of the design: Note
 * Taker is the authorization decision. Passthrough means Entra validates
 * nothing at the door, so Note Taker must check `aud`, `roles`, `appid` and
 * `tid` on our bearer, and must check the asserted caller was actually a
 * participant of the meeting — possession of a thread id proves nothing.
 *
 * Request / response shape below is OUR PROPOSAL (2026-09-10), written so Scott
 * can confirm or amend against something concrete. It is documented for him in
 * docs/2026-0910-* notetaker-api-contract.
 */

export type NoteTakerRefusal =
  /** Ours: MS365_MCP_NOTETAKER_API_URL / _AUDIENCE not set. */
  | 'not_configured'
  /** Ours: the managed identity could not mint a token for the audience. */
  | 'token_unavailable'
  /** Network failure or an unexpected 5xx. */
  | 'unreachable'
  /** 401 — Note Taker rejected OUR bearer. A deployment fault, never the user's. */
  | 'unauthorized'
  /** 403 not_participant — the asserted caller was not in the meeting. */
  | 'not_participant'
  /** 403 transcripts_disabled — the tenant's "no programmatic transcript access" switch. */
  | 'transcripts_disabled'
  /** 404 — Note Taker has no record of the meeting. */
  | 'meeting_not_found'
  /** 429 — Note Taker's per-caller or daily ceiling. */
  | 'rate_limited'
  /** 503 kill_switch — the endpoint was turned off without a deploy. */
  | 'disabled'
  | 'unexpected';

export class NoteTakerError extends Error {
  constructor(
    message: string,
    readonly refusal: NoteTakerRefusal,
    readonly status?: number,
    readonly retryAfterSeconds?: number
  ) {
    super(message);
    this.name = 'NoteTakerError';
  }
}

export interface NoteTakerConfig {
  /** External App Proxy URL of the API, no trailing slash. */
  apiUrl: string;
  /** App ID URI of the Note Taker API registration — the token audience. */
  audience: string;
  /** Client id of the user-assigned managed identity to mint the token with. */
  identityClientId?: string;
}

export function noteTakerConfigFromEnv(
  env: Record<string, string | undefined> = process.env
): NoteTakerConfig | undefined {
  const apiUrl = env.MS365_MCP_NOTETAKER_API_URL?.trim();
  const audience = env.MS365_MCP_NOTETAKER_AUDIENCE?.trim();
  if (!apiUrl || !audience) return undefined;
  return {
    apiUrl: apiUrl.replace(/\/+$/, ''),
    audience,
    identityClientId: env.MS365_MCP_NOTETAKER_IDENTITY_CLIENT_ID?.trim() || undefined,
  };
}

/** Proposed 200 response. `transcript.available=false` is a defined shape, not a null. */
export interface NoteTakerTranscript {
  threadId: string;
  subject?: string;
  organizer?: { oid?: string; upn?: string; displayName?: string };
  joinUrl?: string;
  start?: string;
  end?: string;
  transcript:
    | { available: true; format: 'text' | 'vtt'; language?: string; content: string }
    | { available: false; reason: string; detail?: string };
}

export type TokenSource = (audience: string) => Promise<string>;

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

  async getTranscript(threadId: string, caller: VerifiedIdentity): Promise<NoteTakerTranscript> {
    let token: string;
    try {
      token = await this.tokenSource(this.config.audience);
    } catch (error) {
      throw new NoteTakerError(
        `Could not obtain a credential for Note Taker: ${(error as Error).message}`,
        'token_unavailable'
      );
    }

    const url = `${this.config.apiUrl}/v1/meetings/${encodeURIComponent(threadId)}/transcript`;
    const headers: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
      'X-Caller-Oid': caller.oid,
      'X-Caller-Upn': caller.upn,
    };
    if (caller.tid) headers['X-Caller-Tid'] = caller.tid;

    let response: Response;
    try {
      response = await this.fetchImpl(url, { headers });
    } catch (error) {
      throw new NoteTakerError(
        `Note Taker is unreachable: ${(error as Error).message}`,
        'unreachable'
      );
    }

    if (response.ok) {
      const body = (await response.json()) as Partial<NoteTakerTranscript>;
      if (
        typeof body.threadId !== 'string' ||
        !body.transcript ||
        typeof body.transcript !== 'object'
      ) {
        throw new NoteTakerError(
          'Note Taker returned a response without the expected threadId/transcript fields',
          'unexpected',
          response.status
        );
      }
      return body as NoteTakerTranscript;
    }

    const body = await readJsonQuietly(response);
    const code = typeof body?.error === 'string' ? body.error : undefined;
    logger.warn(
      `Note Taker refused thread ${threadId}: HTTP ${response.status}${code ? ` ${code}` : ''}`
    );

    switch (response.status) {
      case 401:
        throw new NoteTakerError(
          "Note Taker rejected the connector's own credential (HTTP 401). The audience or app-role assignment is wrong on the server — not something the caller can fix.",
          'unauthorized',
          401
        );
      case 403:
        if (code === 'transcripts_disabled') {
          throw new NoteTakerError(
            'Programmatic transcript access is switched off for the organisation; Note Taker honours that setting.',
            'transcripts_disabled',
            403
          );
        }
        throw new NoteTakerError(
          `Note Taker did not find ${caller.upn} among the participants of this meeting.`,
          'not_participant',
          403
        );
      case 404:
        throw new NoteTakerError(
          'Note Taker has no record of this meeting.',
          'meeting_not_found',
          404
        );
      case 429: {
        const retryAfter = retryAfterSeconds(response, body);
        throw new NoteTakerError(
          `Note Taker's rate limit was reached${retryAfter ? `; retry after ${retryAfter}s` : ''}.`,
          'rate_limited',
          429,
          retryAfter
        );
      }
      case 503:
        throw new NoteTakerError(
          'The transcript endpoint has been switched off by its owner.',
          'disabled',
          503
        );
      default:
        throw new NoteTakerError(
          `Note Taker returned HTTP ${response.status}${code ? ` (${code})` : ''}`,
          response.status >= 500 ? 'unreachable' : 'unexpected',
          response.status
        );
    }
  }
}

async function readJsonQuietly(response: Response): Promise<Record<string, unknown> | undefined> {
  try {
    const parsed: unknown = await response.json();
    return parsed && typeof parsed === 'object' ? (parsed as Record<string, unknown>) : undefined;
  } catch {
    return undefined;
  }
}

function retryAfterSeconds(response: Response, body?: Record<string, unknown>): number | undefined {
  const header = Number(response.headers?.get?.('Retry-After'));
  if (Number.isFinite(header) && header > 0) return header;
  const fromBody = body?.retryAfterSeconds;
  return typeof fromBody === 'number' && fromBody > 0 ? fromBody : undefined;
}

/**
 * Mints tokens with Azure's credential chain. In the container app that resolves
 * to the user-assigned managed identity (`tsqm365mcp-uami`); on a developer
 * machine it falls through to `az login`, so the token path can be exercised
 * locally against our own audience before Note Taker's endpoint exists.
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
