import { describe, expect, it, vi } from 'vitest';
import {
  DEFAULT_TIMEOUT_MS,
  NOTETAKER_API_CODES,
  NoteTakerClient,
  NoteTakerError,
  noteTakerConfigFromEnv,
  transcriptsEnabledFromEnv,
} from '../src/lib/notetaker-client.js';
import logger from '../src/logger.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn() },
}));

const CALLER = {
  oid: '2784a67f-8b11-4bec-9f7d-5914cee129b0',
  upn: 'mike.mrozek@townsquaremedia.com',
  tid: 'a473edd8-ba25-4f04-a0a8-e8ad25c19632',
};
const BASE = 'https://notetakerapi.example.msappproxy.net';
const CONFIG = { apiUrl: BASE, audience: 'api://nt', timeoutMs: 1000 };
const ROW = {
  id: 'm-1',
  subject: 'IT Controls',
  start: '2026-10-01T14:00:00Z',
  end: '2026-10-01T14:30:00Z',
  conversation_id: '19:meeting_abc@thread.v2',
  has_summary: true,
  has_transcript: true,
  attended: true,
};

function res(status: number, body?: unknown, headers: Record<string, string> = {}) {
  const text = body === undefined ? '' : typeof body === 'string' ? body : JSON.stringify(body);
  return {
    ok: status >= 200 && status < 300,
    status,
    headers: { get: (k: string) => headers[k] ?? headers[k.toLowerCase()] ?? null },
    text: async () => text,
  } as unknown as Response;
}

function client(fetchImpl: typeof fetch, tokenSource = async () => 'nt-token', timeoutMs = 1000) {
  return new NoteTakerClient({ ...CONFIG, timeoutMs }, { fetchImpl, tokenSource });
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

  it('normalises the URL, carries the identity client id, defaults the timeout', () => {
    const cfg = noteTakerConfigFromEnv({
      MS365_MCP_NOTETAKER_API_URL: 'https://x.msappproxy.net/ ',
      MS365_MCP_NOTETAKER_AUDIENCE: ' api://guid ',
      MS365_MCP_NOTETAKER_IDENTITY_CLIENT_ID: '541e740c-e8bd-4767-9aea-eed1eca4144b',
    });
    expect(cfg).toEqual({
      apiUrl: 'https://x.msappproxy.net',
      audience: 'api://guid',
      identityClientId: '541e740c-e8bd-4767-9aea-eed1eca4144b',
      timeoutMs: DEFAULT_TIMEOUT_MS,
    });
  });

  it('strips an /api/v1 suffix so the prefix is applied exactly once', () => {
    const cfg = noteTakerConfigFromEnv({
      MS365_MCP_NOTETAKER_API_URL: 'https://x.msappproxy.net/api/v1/',
      MS365_MCP_NOTETAKER_AUDIENCE: 'api://guid',
    });
    expect(cfg?.apiUrl).toBe('https://x.msappproxy.net');
  });

  it('takes the timeout from MS365_MCP_NOTETAKER_TIMEOUT_MS, ignoring junk', () => {
    const base = { MS365_MCP_NOTETAKER_API_URL: 'https://x', MS365_MCP_NOTETAKER_AUDIENCE: 'a' };
    expect(
      noteTakerConfigFromEnv({ ...base, MS365_MCP_NOTETAKER_TIMEOUT_MS: '5000' })?.timeoutMs
    ).toBe(5000);
    expect(
      noteTakerConfigFromEnv({ ...base, MS365_MCP_NOTETAKER_TIMEOUT_MS: 'soon' })?.timeoutMs
    ).toBe(DEFAULT_TIMEOUT_MS);
  });
});

describe('transcriptsEnabledFromEnv', () => {
  it('is off unless explicitly switched on', () => {
    expect(transcriptsEnabledFromEnv({})).toBe(false);
    expect(transcriptsEnabledFromEnv({ MS365_MCP_NOTETAKER_TRANSCRIPTS_ENABLED: 'false' })).toBe(
      false
    );
    expect(transcriptsEnabledFromEnv({ MS365_MCP_NOTETAKER_TRANSCRIPTS_ENABLED: 'true' })).toBe(
      true
    );
    expect(transcriptsEnabledFromEnv({ MS365_MCP_NOTETAKER_TRANSCRIPTS_ENABLED: '1' })).toBe(true);
  });
});

describe('requests', () => {
  it('lists meetings with filters, our bearer and ONLY the user-oid identity header', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValue(res(200, { meetings: [ROW], total: 1, request_id: 'r-1' }));
    const out = await client(fetchImpl).listMeetings(CALLER, {
      start_after: '2026-10-01T00:00:00Z',
      subject_contains: 'IT',
      limit: 10,
      offset: 0,
    });

    expect(out.meetings).toEqual([ROW]);
    expect(out.paging).toEqual({ total: 1 });
    expect(out.requestId).toBe('r-1');
    const [url, init] = fetchImpl.mock.calls[0];
    expect(url).toBe(
      `${BASE}/api/v1/meetings?start_after=2026-10-01T00%3A00%3A00Z&subject_contains=IT&limit=10&offset=0`
    );
    expect(init.method).toBe('GET');
    expect(init.headers).toEqual({
      Authorization: 'Bearer nt-token',
      Accept: 'application/json',
      'X-NoteTaker-User-Oid': CALLER.oid,
    });
    expect(init.signal).toBeInstanceOf(AbortSignal);
  });

  it('accepts a bare array of rows', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(res(200, [ROW]));
    const out = await client(fetchImpl).listMeetings(CALLER);
    expect(out.meetings).toEqual([ROW]);
    expect(fetchImpl.mock.calls[0][0]).toBe(`${BASE}/api/v1/meetings`);
  });

  it('posts a lookup with conversation id and window as JSON', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(res(200, { meetings: [] }));
    const out = await client(fetchImpl).lookupMeetings(CALLER, {
      conversation_id: '19:meeting_abc@thread.v2',
      start_after: '2026-09-30T14:00:00Z',
      start_before: '2026-10-02T14:00:00Z',
    });
    expect(out.meetings).toEqual([]);
    const [url, init] = fetchImpl.mock.calls[0];
    expect(url).toBe(`${BASE}/api/v1/meetings/lookup`);
    expect(init.method).toBe('POST');
    expect(init.headers['Content-Type']).toBe('application/json');
    expect(JSON.parse(init.body)).toEqual({
      conversation_id: '19:meeting_abc@thread.v2',
      start_after: '2026-09-30T14:00:00Z',
      start_before: '2026-10-02T14:00:00Z',
    });
  });

  it('fetches meeting detail by Note Taker id, URL-encoded, request id from header', async () => {
    const detail = { ...ROW, summary: { text: 's', decisions: [], action_items: [] } };
    const fetchImpl = vi.fn().mockResolvedValue(res(200, detail, { 'X-Request-Id': 'hdr-9' }));
    const out = await client(fetchImpl).getMeeting(CALLER, 'a/b');
    expect(out.meeting).toEqual(detail);
    expect(out.requestId).toBe('hdr-9');
    expect(fetchImpl.mock.calls[0][0]).toBe(`${BASE}/api/v1/meetings/a%2Fb`);
  });

  it('fetches a transcript page and splits text from paging metadata', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValue(res(200, { text: 'hello', next_page: '2', request_id: 'r' }));
    const out = await client(fetchImpl).getTranscript(CALLER, 'm-1', '1');
    expect(out.text).toBe('hello');
    expect(out.meta).toMatchObject({ next_page: '2' });
    expect(fetchImpl.mock.calls[0][0]).toBe(`${BASE}/api/v1/meetings/m-1/transcript?page=1`);
  });

  it('calls status with our token and no user header', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(res(200, { ok: true }));
    await client(fetchImpl).status();
    const [url, init] = fetchImpl.mock.calls[0];
    expect(url).toBe(`${BASE}/api/v1/status`);
    expect(init.headers['X-NoteTaker-User-Oid']).toBeUndefined();
  });

  it('never sends the retired X-Caller-* headers', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(res(200, { meetings: [] }));
    await client(fetchImpl).listMeetings(CALLER);
    const headers = Object.keys(fetchImpl.mock.calls[0][1].headers).map((h) => h.toLowerCase());
    expect(headers.filter((h) => h.startsWith('x-caller'))).toEqual([]);
  });

  it('mints the token for the configured audience', async () => {
    const tokenSource = vi.fn().mockResolvedValue('t');
    const fetchImpl = vi.fn().mockResolvedValue(res(200, { meetings: [] }));
    await client(fetchImpl, tokenSource).listMeetings(CALLER);
    expect(tokenSource).toHaveBeenCalledWith('api://nt');
  });

  it.each([
    ['a list without rows', { total: 0 }],
    ['rows without ids', { meetings: [{ subject: 'x' }] }],
    ['non-JSON', '<html>proxy</html>'],
  ])('rejects a 200 with %s rather than guessing', async (_label, body) => {
    const fetchImpl = vi.fn().mockResolvedValue(res(200, body));
    const e = await refusal(() => client(fetchImpl).listMeetings(CALLER));
    expect(e.code).toBe('unexpected_response');
  });
});

describe('refusals', () => {
  const statusOf: Record<string, number> = {
    invalid_token: 401,
    token_expired: 401,
    caller_not_allowed: 403,
    app_token_required: 403,
    missing_role: 403,
    user_identity_missing: 400,
    user_identity_invalid: 400,
    user_not_enabled: 403,
    meeting_not_found: 404,
    tenant_transcript_access_disabled: 403,
    api_disabled: 503,
    transcripts_disabled: 503,
    rate_limited_user: 429,
    rate_limited_connector: 429,
    daily_limit_user: 429,
    daily_limit_connector: 429,
    upstream_unavailable: 503,
  };

  it('covers every code in the spec', () => {
    expect(Object.keys(statusOf).sort()).toEqual([...NOTETAKER_API_CODES].sort());
  });

  it.each(NOTETAKER_API_CODES.map((c) => [c]))(
    'maps %s from the JSON body, with request id',
    async (code) => {
      const fetchImpl = vi
        .fn()
        .mockResolvedValue(
          res(statusOf[code], { code, message: 'server text', request_id: 'req-42' })
        );
      const e = await refusal(() => client(fetchImpl).getMeeting(CALLER, 'm-1'));
      expect(e.code).toBe(code);
      expect(e.status).toBe(statusOf[code]);
      expect(e.requestId).toBe('req-42');
    }
  );

  it('tolerates camelCase and nested error bodies', async () => {
    const camel = vi
      .fn()
      .mockResolvedValue(
        res(429, { Code: 'rate_limited_user', requestId: 'c-1', retryAfterSeconds: 90 })
      );
    const a = await refusal(() => client(camel).listMeetings(CALLER));
    expect(a).toMatchObject({ code: 'rate_limited_user', requestId: 'c-1', retryAfterSeconds: 90 });

    const nested = vi
      .fn()
      .mockResolvedValue(res(404, { error: { code: 'meeting_not_found', request_id: 'n-1' } }));
    const b = await refusal(() => client(nested).getMeeting(CALLER, 'x'));
    expect(b).toMatchObject({ code: 'meeting_not_found', requestId: 'n-1' });
  });

  it('prefers the Retry-After header over the body', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValue(
        res(
          429,
          { code: 'daily_limit_connector', retry_after_seconds: 999 },
          { 'Retry-After': '30' }
        )
      );
    const e = await refusal(() => client(fetchImpl).listMeetings(CALLER));
    expect(e.retryAfterSeconds).toBe(30);
  });

  it.each([
    [502, '<html>Bad gateway</html>', 'unreachable'],
    [500, undefined, 'unreachable'],
    [429, undefined, 'rate_limited'],
    [404, '<html>not here</html>', 'unexpected_response'],
    [403, { code: 'brand_new_code' }, 'unexpected_response'],
  ] as const)('falls back by status for HTTP %s with no known code', async (status, body, code) => {
    const fetchImpl = vi.fn().mockResolvedValue(res(status, body));
    const e = await refusal(() => client(fetchImpl).getMeeting(CALLER, 'm'));
    expect(e.code).toBe(code);
    expect(e.status).toBe(status);
  });

  it('keeps an unknown server code, sanitised, for diagnosis', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(res(403, { code: 'Brand New<script>' }));
    const e = await refusal(() => client(fetchImpl).getMeeting(CALLER, 'm'));
    expect(e.serverCode).toBe('brandnewscript');
  });

  it('logs status, code and request id but not the server message', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValue(
        res(404, { code: 'meeting_not_found', message: 'Secret Board Meeting', request_id: 'q' })
      );
    await refusal(() => client(fetchImpl).getMeeting(CALLER, 'm'));
    const logged = vi.mocked(logger.warn).mock.calls.map((c) => String(c[0]));
    expect(logged.some((l) => l.includes('meeting_not_found') && l.includes('request_id=q'))).toBe(
      true
    );
    expect(logged.some((l) => l.includes('Secret Board Meeting'))).toBe(false);
  });
});

describe('failure modes before a response', () => {
  it('reports a network failure as unreachable', async () => {
    const fetchImpl = vi.fn().mockRejectedValue(new Error('ECONNRESET'));
    const e = await refusal(() => client(fetchImpl).listMeetings(CALLER));
    expect(e.code).toBe('unreachable');
    expect(e.message).toContain('ECONNRESET');
  });

  it('aborts a request that exceeds the timeout', async () => {
    const fetchImpl = vi.fn(
      (_url: string, init: { signal?: AbortSignal }) =>
        new Promise<Response>((_resolve, reject) => {
          init.signal?.addEventListener('abort', () => reject(new Error('aborted')));
        })
    ) as unknown as typeof fetch;
    const e = await refusal(() => client(fetchImpl, async () => 't', 20).listMeetings(CALLER));
    expect(e.code).toBe('timeout');
  });

  it('times out a hung token source without calling Note Taker', async () => {
    const fetchImpl = vi.fn();
    const e = await refusal(() =>
      client(fetchImpl, () => new Promise<string>(() => {}), 20).listMeetings(CALLER)
    );
    expect(e.code).toBe('token_unavailable');
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it('fails closed when no credential can be minted, without calling Note Taker', async () => {
    const fetchImpl = vi.fn();
    const e = await refusal(() =>
      client(fetchImpl, async () => {
        throw new Error('IMDS unavailable');
      }).listMeetings(CALLER)
    );
    expect(e.code).toBe('token_unavailable');
    expect(fetchImpl).not.toHaveBeenCalled();
  });
});
