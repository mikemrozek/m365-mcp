import { beforeEach, describe, expect, it, vi } from 'vitest';
import {
  DEFAULT_MAX_CHARS,
  MEETING_TOOL_NAMES,
  REFUSALS,
  explainRefusal,
  registerMeetingTools,
  resolveThreadId,
  toUtcIso,
} from '../src/meeting-tools.js';
import {
  NOTETAKER_API_CODES,
  NoteTakerError,
  type NoteTakerClient,
} from '../src/lib/notetaker-client.js';
import { getRequestTokens } from '../src/request-context.js';
import { IdentityVerificationError, verifyCallerIdentity } from '../src/lib/verified-identity.js';
import type GraphClient from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn() },
}));
vi.mock('../src/usage-log.js', () => ({
  withUsageLog: (_name: string, fn: () => Promise<unknown>) => fn(),
  logToolUsage: vi.fn(),
}));
vi.mock('../src/request-context.js', () => ({ getRequestTokens: vi.fn() }));
vi.mock('../src/lib/verified-identity.js', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../src/lib/verified-identity.js')>();
  return { ...actual, verifyCallerIdentity: vi.fn() };
});

const CALLER = { oid: 'oid-1', upn: 'mike.mrozek@townsquaremedia.com', tid: 'tid-1' };
const THREAD = '19:meeting_abc123@thread.v2';
const JOIN_URL = `https://teams.microsoft.com/l/meetup-join/${encodeURIComponent(THREAD).replace(/%3A/gi, '%3a')}/0?context=%7b%22Tid%22%3a%22t%22%7d`;
const RECAP_URL = `https://teams.microsoft.com/v2/#/meetingrecap?threadId=${encodeURIComponent(THREAD)}&tenantId=t&organizerId=o`;
const SHORT_URL = 'https://teams.microsoft.com/meet/29752586464443?p=abc';

const ROW = (id: string, start = '2026-10-01T14:00:00Z') => ({
  id,
  subject: 'IT Controls',
  start,
  conversation_id: THREAD,
  has_summary: true,
});
const DETAIL = {
  ...ROW('m-1'),
  summary: { text: 'We agreed X', decisions: ['X'], action_items: ['Mike: do Y'] },
  participants: [{ name: 'Mike' }],
};

type ToolResult = {
  content: { text: string }[];
  isError?: boolean;
  _meta?: { noteTakerRequestIds?: string[] };
};
type Handler = (p: Record<string, unknown>) => Promise<ToolResult>;

function harness(
  noteTaker: Partial<Record<keyof NoteTakerClient, unknown>> | null,
  opts: { makeRequest?: ReturnType<typeof vi.fn>; transcriptsEnabled?: boolean } = {}
) {
  const handlers: Record<string, Handler> = {};
  const schemas: Record<string, Record<string, unknown>> = {};
  const descriptions: Record<string, string> = {};
  const server = {
    tool: (n: string, d: string, s: Record<string, unknown>, _a: unknown, h: Handler) => {
      handlers[n] = h;
      schemas[n] = s;
      descriptions[n] = d;
    },
  } as unknown as Parameters<typeof registerMeetingTools>[0];
  const registered: string[] = [];
  const makeRequest = opts.makeRequest ?? vi.fn();
  registerMeetingTools(
    server,
    { makeRequest } as unknown as GraphClient,
    {
      isToolEnabled: () => true,
      push: (n) => registered.push(n),
      fail: (n, e) => {
        throw new Error(`registration failed for ${n}: ${e.message}`);
      },
    },
    {
      noteTaker: noteTaker as NoteTakerClient | null,
      transcriptsEnabled: opts.transcriptsEnabled ?? false,
    }
  );
  const call = async (tool: string, params: Record<string, unknown>) => {
    const r = await handlers[tool](params);
    return { payload: JSON.parse(r.content[0].text), isError: r.isError, meta: r._meta };
  };
  return { registered, call, makeRequest, schemas, descriptions };
}

function stubNoteTaker(overrides: Partial<Record<keyof NoteTakerClient, unknown>> = {}) {
  return {
    listMeetings: vi.fn(async () => ({ meetings: [ROW('m-1')], paging: {}, requestId: 'r-list' })),
    lookupMeetings: vi.fn(async () => ({ meetings: [ROW('m-1')], paging: {}, requestId: 'r-lk' })),
    getMeeting: vi.fn(async () => ({ meeting: DETAIL, requestId: 'r-det' })),
    getTranscript: vi.fn(async () => ({ text: 'hello', meta: {}, requestId: 'r-tx' })),
    ...overrides,
  };
}

beforeEach(() => {
  vi.mocked(getRequestTokens).mockReturnValue({ accessToken: 'graph-token' });
  vi.mocked(verifyCallerIdentity).mockReset().mockResolvedValue(CALLER);
});

describe('registration', () => {
  it('registers the three meeting tools', () => {
    expect(harness(null).registered).toEqual([...MEETING_TOOL_NAMES]);
  });

  it('respects the allowlist per tool', () => {
    const server = { tool: vi.fn() } as unknown as Parameters<typeof registerMeetingTools>[0];
    const registered: string[] = [];
    registerMeetingTools(
      server,
      {} as GraphClient,
      {
        isToolEnabled: (n) => n === 'get-meeting',
        push: (n) => registered.push(n),
        fail: () => {},
      },
      { noteTaker: null }
    );
    expect(registered).toEqual(['get-meeting']);
  });

  it.each(MEETING_TOOL_NAMES.map((n) => [n]))(
    '%s takes no user or object id: identity comes from the session only',
    (name) => {
      const { schemas } = harness(null);
      const keys = Object.keys(schemas[name]);
      expect(keys.filter((k) => /oid|user|upn|caller|identity|principal|email$/i.test(k))).toEqual(
        name === 'list-my-meetings' ? ['organizerEmail'] : []
      );
    }
  );

  it('front-loads the words people type', () => {
    const { descriptions } = harness(null);
    expect(descriptions['list-my-meetings']).toMatch(/^List my meetings/);
    expect(descriptions['get-meeting']).toMatch(
      /^Get meeting notes: the summary, decisions, action items/
    );
    expect(descriptions['get-meeting']).toMatch(/recap/);
    expect(descriptions['get-meeting-transcript']).toMatch(/^Get the transcript of a meeting/);
  });

  it.each(['list-my-meetings', 'get-meeting'])(
    '%s is dark until Note Taker is configured: says so, touches nothing',
    async (tool) => {
      const { call, makeRequest } = harness(null);
      const { payload, isError } = await call(tool, { meetingId: 'm-1' });
      expect(isError).toBe(true);
      expect(payload.error).toBe('not_configured');
      expect(makeRequest).not.toHaveBeenCalled();
      expect(verifyCallerIdentity).not.toHaveBeenCalled();
    }
  );
});

describe('list-my-meetings', () => {
  it('sends the verified identity and UTC-normalised filters to Note Taker', async () => {
    const nt = stubNoteTaker();
    const { payload, isError, meta } = await harness(nt).call('list-my-meetings', {
      startAfter: '2026-10-01',
      startBefore: '2026-10-02T09:00:00-04:00',
      subjectContains: ' IT ',
      organizerEmail: 'tony.chan@townsquaremedia.com',
      limit: 10,
    });
    expect(isError).toBeUndefined();
    expect(verifyCallerIdentity).toHaveBeenCalledWith('graph-token');
    expect(nt.listMeetings).toHaveBeenCalledWith(CALLER, {
      start_after: '2026-10-01T00:00:00Z',
      start_before: '2026-10-02T13:00:00Z',
      subject_contains: 'IT',
      organizer_email: 'tony.chan@townsquaremedia.com',
      limit: 10,
      offset: undefined,
    });
    expect(payload).toMatchObject({ count: 1, meetings: [ROW('m-1')], requestId: 'r-list' });
    expect(meta).toEqual({ noteTakerRequestIds: ['r-list'] });
  });

  it('suggests the next offset when the page is full', async () => {
    const nt = stubNoteTaker({
      listMeetings: vi.fn(async () => ({ meetings: [ROW('a'), ROW('b')], paging: {} })),
    });
    const { payload } = await harness(nt).call('list-my-meetings', { limit: 2, offset: 4 });
    expect(payload.note).toMatch(/offset 6/);
  });

  it('rejects a malformed date without calling Note Taker', async () => {
    const nt = stubNoteTaker();
    const { payload, isError } = await harness(nt).call('list-my-meetings', {
      startAfter: 'last tuesday',
    });
    expect(isError).toBe(true);
    expect(payload.error).toBe('invalid_input');
    expect(nt.listMeetings).not.toHaveBeenCalled();
  });

  it('refuses to call Note Taker when identity cannot be verified', async () => {
    vi.mocked(verifyCallerIdentity).mockRejectedValue(
      new IdentityVerificationError('nope', 'graph_rejected')
    );
    const nt = stubNoteTaker();
    const { payload, isError } = await harness(nt).call('list-my-meetings', {});
    expect(isError).toBe(true);
    expect(payload).toMatchObject({ error: 'identity_unverified', reason: 'graph_rejected' });
    expect(nt.listMeetings).not.toHaveBeenCalled();
  });

  it('calls a tenant mismatch a connector configuration problem', async () => {
    vi.mocked(verifyCallerIdentity).mockRejectedValue(
      new IdentityVerificationError('x', 'tenant_mismatch')
    );
    const { payload } = await harness(stubNoteTaker()).call('list-my-meetings', {});
    expect(payload).toMatchObject({ fault: 'connector_configuration', reportTo: 'Mike Mrozek' });
  });

  it('reports a missing caller credential', async () => {
    vi.mocked(getRequestTokens).mockReturnValue(undefined);
    const { payload, isError } = await harness(stubNoteTaker()).call('list-my-meetings', {});
    expect(isError).toBe(true);
    expect(payload.error).toBe('no_caller');
  });
});

describe('get-meeting', () => {
  it('by Note Taker id goes straight to the detail', async () => {
    const nt = stubNoteTaker();
    const { call, makeRequest } = harness(nt);
    const { payload, meta } = await call('get-meeting', { meetingId: ' m-1 ' });
    expect(nt.getMeeting).toHaveBeenCalledWith(CALLER, 'm-1');
    expect(nt.lookupMeetings).not.toHaveBeenCalled();
    expect(makeRequest).not.toHaveBeenCalled();
    expect(payload).toEqual({ meeting: DETAIL, requestId: 'r-det' });
    expect(meta).toEqual({ noteTakerRequestIds: ['r-det'] });
  });

  it('from a calendar event: resolves the conversation, windows ±1 day, fetches the one match', async () => {
    const makeRequest = vi.fn().mockResolvedValue({
      subject: 'IT Controls',
      start: { dateTime: '2026-10-01T14:00:00.0000000', timeZone: 'UTC' },
      onlineMeeting: { joinUrl: JOIN_URL },
    });
    const nt = stubNoteTaker();
    const { payload, meta } = await harness(nt, { makeRequest }).call('get-meeting', {
      eventId: 'AAMk=',
    });
    expect(nt.lookupMeetings).toHaveBeenCalledWith(CALLER, {
      conversation_id: THREAD,
      start_after: '2026-09-30T14:00:00Z',
      start_before: '2026-10-02T14:00:00Z',
    });
    expect(nt.getMeeting).toHaveBeenCalledWith(CALLER, 'm-1');
    expect(payload).toMatchObject({
      meeting: DETAIL,
      resolvedVia: 'event',
      conversationId: THREAD,
      requestIds: ['r-lk', 'r-det'],
    });
    expect(meta).toEqual({ noteTakerRequestIds: ['r-lk', 'r-det'] });
  });

  it('explicit window overrides the event-derived one', async () => {
    const makeRequest = vi.fn().mockResolvedValue({
      start: { dateTime: '2026-10-01T14:00:00.0000000', timeZone: 'UTC' },
      onlineMeeting: { joinUrl: JOIN_URL },
    });
    const nt = stubNoteTaker();
    await harness(nt, { makeRequest }).call('get-meeting', {
      eventId: 'e',
      startAfter: '2026-10-01T13:00:00Z',
    });
    expect(nt.lookupMeetings).toHaveBeenCalledWith(CALLER, {
      conversation_id: THREAD,
      start_after: '2026-10-01T13:00:00Z',
      start_before: '2026-10-02T14:00:00Z',
    });
  });

  it('several occurrences: returns them to choose from, no detail call', async () => {
    const nt = stubNoteTaker({
      lookupMeetings: vi.fn(async () => ({
        meetings: [ROW('m-1'), ROW('m-2', '2026-10-08T14:00:00Z')],
        paging: {},
        requestId: 'r-lk',
      })),
    });
    const { payload, isError } = await harness(nt).call('get-meeting', { meetingUrl: RECAP_URL });
    expect(isError).toBeUndefined();
    expect(nt.lookupMeetings).toHaveBeenCalledWith(CALLER, { conversation_id: THREAD });
    expect(nt.getMeeting).not.toHaveBeenCalled();
    expect(payload).toMatchObject({ found: 2, resolvedVia: 'url' });
    expect(payload.meetings.map((m: { id: string }) => m.id)).toEqual(['m-1', 'm-2']);
    expect(payload.message).toMatch(/meetingId/);
    expect(payload.message).toMatch(/startAfter\/startBefore/);
  });

  it('no occurrences: says so and mentions channel meetings and /meet/ links', async () => {
    const nt = stubNoteTaker({
      lookupMeetings: vi.fn(async () => ({ meetings: [], paging: {}, requestId: 'r' })),
    });
    const { payload, isError } = await harness(nt).call('get-meeting', { threadId: THREAD });
    expect(isError).toBeUndefined();
    expect(payload.found).toBe(0);
    expect(payload.message).toMatch(/Channel meetings/);
    expect(payload.message).toMatch(/\/meet\//);
    expect(nt.getMeeting).not.toHaveBeenCalled();
  });

  it('an unresolvable reference is an error without a Note Taker call', async () => {
    const nt = stubNoteTaker();
    const { payload, isError } = await harness(nt).call('get-meeting', { threadId: 'garbage' });
    expect(isError).toBe(true);
    expect(payload.error).toBe('unresolved_meeting');
    expect(nt.lookupMeetings).not.toHaveBeenCalled();
  });

  it('asks for an identifier when given none', async () => {
    const { payload, isError } = await harness(stubNoteTaker()).call('get-meeting', {});
    expect(isError).toBe(true);
    expect(payload.message).toMatch(/meetingId, eventId, meetingUrl or threadId/);
  });

  it('relays meeting_not_found in plain language with the request id', async () => {
    const nt = stubNoteTaker({
      getMeeting: vi.fn(async () => {
        throw new NoteTakerError('x', 'meeting_not_found', 404, undefined, 'req-7');
      }),
    });
    const { payload, isError, meta } = await harness(nt).call('get-meeting', { meetingId: 'm' });
    expect(isError).toBe(true);
    expect(payload).toMatchObject({
      error: 'meeting_not_found',
      status: 404,
      fault: 'not_found',
      requestId: 'req-7',
    });
    expect(payload.message).toMatch(/you were not in it/);
    expect(meta).toEqual({ noteTakerRequestIds: ['req-7'] });
  });
});

describe('get-meeting-transcript', () => {
  it('reports not_enabled when the transcripts flag is off, touching nothing', async () => {
    const nt = stubNoteTaker();
    const { payload, isError } = await harness(nt).call('get-meeting-transcript', {
      meetingId: 'm-1',
    });
    expect(isError).toBe(true);
    expect(payload.error).toBe('not_enabled');
    expect(nt.getTranscript).not.toHaveBeenCalled();
    expect(verifyCallerIdentity).not.toHaveBeenCalled();
  });

  it('when enabled: fetches by Note Taker id with the page token', async () => {
    const nt = stubNoteTaker({
      getTranscript: vi.fn(async () => ({
        text: 'hello',
        meta: { next_page: '3' },
        requestId: 't',
      })),
    });
    const { payload } = await harness(nt, { transcriptsEnabled: true }).call(
      'get-meeting-transcript',
      { meetingId: 'm-1', page: '2' }
    );
    expect(nt.getTranscript).toHaveBeenCalledWith(CALLER, 'm-1', '2');
    expect(payload).toMatchObject({
      meetingId: 'm-1',
      text: 'hello',
      next_page: '3',
      totalChars: 5,
      truncated: false,
      requestId: 't',
    });
  });

  it('caps long pages at maxChars and says so', async () => {
    const nt = stubNoteTaker({
      getTranscript: vi.fn(async () => ({ text: 'x'.repeat(1200), meta: {} })),
    });
    const { payload } = await harness(nt, { transcriptsEnabled: true }).call(
      'get-meeting-transcript',
      { meetingId: 'm-1', maxChars: 500 }
    );
    expect(payload.text).toHaveLength(500);
    expect(payload).toMatchObject({ totalChars: 1200, truncated: true });
    expect(payload.note).toMatch(/cut at 500 of 1200/);
  });

  it('defaults the cap to DEFAULT_MAX_CHARS', async () => {
    const nt = stubNoteTaker({
      getTranscript: vi.fn(async () => ({ text: 'y'.repeat(DEFAULT_MAX_CHARS + 1), meta: {} })),
    });
    const { payload } = await harness(nt, { transcriptsEnabled: true }).call(
      'get-meeting-transcript',
      { meetingId: 'm-1' }
    );
    expect(payload.text).toHaveLength(DEFAULT_MAX_CHARS);
    expect(payload.truncated).toBe(true);
  });

  it('takes only a Note Taker id, not a thread id or URL', () => {
    const { schemas } = harness(null);
    expect(Object.keys(schemas['get-meeting-transcript']).sort()).toEqual([
      'maxChars',
      'meetingId',
      'page',
    ]);
  });
});

describe('explainRefusal', () => {
  it('has a distinct message for every Note Taker code', () => {
    const messages = NOTETAKER_API_CODES.map((c) => REFUSALS[c].message);
    expect(new Set(messages).size).toBe(NOTETAKER_API_CODES.length);
  });

  it.each([
    'invalid_token',
    'token_expired',
    'caller_not_allowed',
    'app_token_required',
    'missing_role',
    'user_identity_missing',
    'user_identity_invalid',
  ] as const)('%s is a connector configuration problem to report to Mike', (code) => {
    const out = explainRefusal(new NoteTakerError('server text', code, 403, undefined, 'rid'));
    expect(out).toMatchObject({
      error: code,
      fault: 'connector_configuration',
      reportTo: 'Mike Mrozek',
      requestId: 'rid',
    });
    expect(out.message).toMatch(/not something you did/);
  });

  it.each([
    ['user_not_enabled', 'access'],
    ['meeting_not_found', 'not_found'],
    ['tenant_transcript_access_disabled', 'switched_off'],
    ['api_disabled', 'switched_off'],
    ['transcripts_disabled', 'switched_off'],
    ['upstream_unavailable', 'temporary'],
    ['timeout', 'temporary'],
  ] as const)('%s → %s, not blamed on configuration', (code, fault) => {
    const out = explainRefusal(new NoteTakerError('x', code));
    expect(out.fault).toBe(fault);
    expect(out.reportTo).toBeUndefined();
  });

  it.each([
    'rate_limited_user',
    'rate_limited_connector',
    'daily_limit_user',
    'daily_limit_connector',
  ] as const)('%s carries the retry delay into the message', (code) => {
    const out = explainRefusal(new NoteTakerError('x', code, 429, 45));
    expect(out).toMatchObject({ error: code, fault: 'rate_limit', retryAfterSeconds: 45 });
    expect(out.message).toMatch(/about 45 seconds/);
  });

  it("never relays Note Taker's own message text", () => {
    const out = explainRefusal(
      new NoteTakerError('Board meeting re: layoffs', 'meeting_not_found')
    );
    expect(JSON.stringify(out)).not.toMatch(/layoffs/);
  });
});

describe('resolveThreadId', () => {
  it('accepts a well-formed thread id and rejects anything else', async () => {
    const makeRequest = vi.fn();
    await expect(resolveThreadId({ makeRequest }, { threadId: ` ${THREAD} ` })).resolves.toEqual({
      threadId: THREAD,
      via: 'threadId',
    });
    await expect(
      resolveThreadId({ makeRequest }, { threadId: '19:abc@unq.gbl.spaces' })
    ).rejects.toThrow(/not a meeting thread id/);
    expect(makeRequest).not.toHaveBeenCalled();
  });

  it('resolves recap and full join URLs locally, with no Graph call', async () => {
    const makeRequest = vi.fn();
    expect((await resolveThreadId({ makeRequest }, { meetingUrl: RECAP_URL })).threadId).toBe(
      THREAD
    );
    expect((await resolveThreadId({ makeRequest }, { meetingUrl: JOIN_URL })).via).toBe('url');
    expect(makeRequest).not.toHaveBeenCalled();
  });

  it('resolves a short /meet/ link through /me/onlineMeetings', async () => {
    const makeRequest = vi.fn().mockResolvedValue({ value: [{ chatInfo: { threadId: THREAD } }] });
    const out = await resolveThreadId({ makeRequest }, { meetingUrl: SHORT_URL });
    expect(out).toEqual({ threadId: THREAD, via: 'graph-lookup' });
    expect(makeRequest.mock.calls[0][0]).toBe(
      `/me/onlineMeetings?$filter=joinWebUrl eq '${SHORT_URL}'&$select=id,subject,chatInfo`
    );
  });

  it('explains when a short link resolves to nothing', async () => {
    const makeRequest = vi.fn().mockResolvedValue({ value: [] });
    await expect(resolveThreadId({ makeRequest }, { meetingUrl: SHORT_URL })).rejects.toThrow(
      /Could not resolve that Teams link/
    );
  });

  it('follows a calendar event to its Teams link and carries its start in UTC', async () => {
    const makeRequest = vi.fn().mockResolvedValue({
      subject: 'Sync',
      start: { dateTime: '2026-10-01T14:00:00.0000000', timeZone: 'UTC' },
      onlineMeeting: { joinUrl: JOIN_URL },
    });
    const out = await resolveThreadId({ makeRequest }, { eventId: 'AAMk=' });
    expect(out).toEqual({ threadId: THREAD, via: 'event', eventStart: '2026-10-01T14:00:00Z' });
    expect(makeRequest.mock.calls[0][0]).toBe(
      '/me/events/AAMk%3D?$select=subject,start,isOnlineMeeting,onlineMeeting'
    );
  });

  it('says when a calendar event is not a Teams meeting', async () => {
    const makeRequest = vi.fn().mockResolvedValue({ subject: 'Lunch' });
    await expect(resolveThreadId({ makeRequest }, { eventId: 'x' })).rejects.toThrow(
      /"Lunch" has no Teams meeting link/
    );
  });
});

describe('toUtcIso', () => {
  it('normalises to UTC with a Z', () => {
    expect(toUtcIso('2026-10-01', 'f')).toBe('2026-10-01T00:00:00Z');
    expect(toUtcIso('2026-10-01T10:00:00+02:00', 'f')).toBe('2026-10-01T08:00:00Z');
    expect(() => toUtcIso('soon', 'startAfter')).toThrow(/startAfter "soon"/);
  });
});
