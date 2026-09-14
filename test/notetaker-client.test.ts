import { describe, expect, it, vi } from 'vitest';
import {
  NoteTakerClient,
  NoteTakerError,
  noteTakerConfigFromEnv,
} from '../src/lib/notetaker-client.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn() },
}));

const CALLER = {
  oid: '2784a67f-8b11-4bec-9f7d-5914cee129b0',
  upn: 'mike.mrozek@townsquaremedia.com',
  tid: 'a473edd8-ba25-4f04-a0a8-e8ad25c19632',
};
const THREAD = '19:meeting_abc123@thread.v2';
const CONFIG = { apiUrl: 'https://notetakerapi.example.msappproxy.net', audience: 'api://nt' };

function res(status: number, body?: unknown, headers: Record<string, string> = {}) {
  return {
    ok: status >= 200 && status < 300,
    status,
    headers: { get: (k: string) => headers[k] ?? headers[k.toLowerCase()] ?? null },
    json: async () => {
      if (body === undefined) throw new Error('no body');
      return body;
    },
  } as unknown as Response;
}

function client(fetchImpl: typeof fetch, tokenSource = async () => 'nt-token') {
  return new NoteTakerClient(CONFIG, { fetchImpl, tokenSource });
}

async function refusal(fn: () => Promise<unknown>): Promise<NoteTakerError> {
  try {
    await fn();
  } catch (e) {
    if (e instanceof NoteTakerError) return e;
    throw e;
  }
  throw new Error('expected a NoteTakerError');
}

describe('noteTakerConfigFromEnv', () => {
  it('is undefined until both the API URL and the audience are set', () => {
    expect(noteTakerConfigFromEnv({})).toBeUndefined();
    expect(noteTakerConfigFromEnv({ MS365_MCP_NOTETAKER_API_URL: 'https://x' })).toBeUndefined();
    expect(noteTakerConfigFromEnv({ MS365_MCP_NOTETAKER_AUDIENCE: 'api://x' })).toBeUndefined();
  });

  it('normalises the URL and carries the managed identity client id', () => {
    const cfg = noteTakerConfigFromEnv({
      MS365_MCP_NOTETAKER_API_URL: 'https://x.msappproxy.net/ ',
      MS365_MCP_NOTETAKER_AUDIENCE: ' api://guid ',
      MS365_MCP_NOTETAKER_IDENTITY_CLIENT_ID: '541e740c-e8bd-4767-9aea-eed1eca4144b',
    });
    expect(cfg).toEqual({
      apiUrl: 'https://x.msappproxy.net',
      audience: 'api://guid',
      identityClientId: '541e740c-e8bd-4767-9aea-eed1eca4144b',
    });
  });
});

describe('NoteTakerClient.getTranscript', () => {
  it('calls the proposed path with our bearer and the verified caller in headers', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValue(
        res(200, {
          threadId: THREAD,
          transcript: { available: true, format: 'text', content: 'hi' },
        })
      );
    const out = await client(fetchImpl).getTranscript(THREAD, CALLER);

    expect(out.transcript).toEqual({ available: true, format: 'text', content: 'hi' });
    const [url, init] = fetchImpl.mock.calls[0];
    expect(url).toBe(`${CONFIG.apiUrl}/v1/meetings/${encodeURIComponent(THREAD)}/transcript`);
    expect(init.headers).toMatchObject({
      Authorization: 'Bearer nt-token',
      'X-Caller-Oid': CALLER.oid,
      'X-Caller-Upn': CALLER.upn,
      'X-Caller-Tid': CALLER.tid,
    });
  });

  it('mints the token for the configured audience', async () => {
    const tokenSource = vi.fn().mockResolvedValue('t');
    const fetchImpl = vi
      .fn()
      .mockResolvedValue(
        res(200, { threadId: THREAD, transcript: { available: false, reason: 'not_recorded' } })
      );
    await client(fetchImpl, tokenSource).getTranscript(THREAD, CALLER);
    expect(tokenSource).toHaveBeenCalledWith('api://nt');
  });

  it('passes the defined no-transcript shape through untouched', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValue(
        res(200, {
          threadId: THREAD,
          transcript: { available: false, reason: 'too_old', detail: 'x' },
        })
      );
    const out = await client(fetchImpl).getTranscript(THREAD, CALLER);
    expect(out.transcript).toEqual({ available: false, reason: 'too_old', detail: 'x' });
  });

  it('rejects a 200 that lacks the contract fields rather than guessing', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(res(200, { text: 'hello' }));
    const e = await refusal(() => client(fetchImpl).getTranscript(THREAD, CALLER));
    expect(e.refusal).toBe('unexpected');
  });

  it.each([
    [401, undefined, 'unauthorized'],
    [403, { error: 'transcripts_disabled' }, 'transcripts_disabled'],
    [403, { error: 'not_participant' }, 'not_participant'],
    [403, undefined, 'not_participant'],
    [404, { error: 'meeting_not_found' }, 'meeting_not_found'],
    [503, { error: 'kill_switch' }, 'disabled'],
    [500, undefined, 'unreachable'],
    [418, undefined, 'unexpected'],
  ] as const)('maps HTTP %s %j to %s', async (status, body, expected) => {
    const fetchImpl = vi.fn().mockResolvedValue(res(status, body));
    const e = await refusal(() => client(fetchImpl).getTranscript(THREAD, CALLER));
    expect(e.refusal).toBe(expected);
    expect(e.status).toBe(status);
  });

  it('surfaces Retry-After on a 429, from the header first and the body second', async () => {
    const withHeader = vi.fn().mockResolvedValue(res(429, undefined, { 'Retry-After': '30' }));
    expect(
      (await refusal(() => client(withHeader).getTranscript(THREAD, CALLER))).retryAfterSeconds
    ).toBe(30);

    const withBody = vi
      .fn()
      .mockResolvedValue(res(429, { error: 'rate_limited', retryAfterSeconds: 90 }));
    expect(
      (await refusal(() => client(withBody).getTranscript(THREAD, CALLER))).retryAfterSeconds
    ).toBe(90);
  });

  it('reports a network failure as unreachable, not as a refusal about the user', async () => {
    const fetchImpl = vi.fn().mockRejectedValue(new Error('ECONNRESET'));
    const e = await refusal(() => client(fetchImpl).getTranscript(THREAD, CALLER));
    expect(e.refusal).toBe('unreachable');
    expect(e.message).toContain('ECONNRESET');
  });

  it('fails closed when no credential can be minted, without calling Note Taker', async () => {
    const fetchImpl = vi.fn();
    const e = await refusal(() =>
      client(fetchImpl, async () => {
        throw new Error('IMDS unavailable');
      }).getTranscript(THREAD, CALLER)
    );
    expect(e.refusal).toBe('token_unavailable');
    expect(fetchImpl).not.toHaveBeenCalled();
  });
});
