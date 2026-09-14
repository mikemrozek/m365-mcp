import { beforeEach, describe, expect, it, vi } from 'vitest';
import { registerMeetingTools, resolveThreadId, DEFAULT_MAX_CHARS } from '../src/meeting-tools.js';
import { NoteTakerError, type NoteTakerClient } from '../src/lib/notetaker-client.js';
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

type Handler = (
  p: Record<string, unknown>
) => Promise<{ content: { text: string }[]; isError?: boolean }>;

function harness(noteTaker: NoteTakerClient | null, makeRequest = vi.fn()) {
  let handler: Handler | undefined;
  const server = {
    tool: (_n: string, _d: string, _s: unknown, _a: unknown, h: Handler) => {
      handler = h;
    },
  } as unknown as Parameters<typeof registerMeetingTools>[0];
  const registered: string[] = [];
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
    { noteTaker }
  );
  const call = async (params: Record<string, unknown>) => {
    const r = await handler!(params);
    return { payload: JSON.parse(r.content[0].text), isError: r.isError };
  };
  return { registered, call, makeRequest };
}

function stubNoteTaker(impl: (threadId: string, caller: typeof CALLER) => Promise<unknown>) {
  return { getTranscript: vi.fn(impl) } as unknown as NoteTakerClient;
}

const transcriptOf = (content: string) => ({
  threadId: THREAD,
  subject: 'IT Controls discussion',
  organizer: { upn: 'tony.chan@townsquaremedia.com' },
  joinUrl: JOIN_URL,
  transcript: { available: true as const, format: 'text' as const, content },
});

beforeEach(() => {
  vi.mocked(getRequestTokens).mockReturnValue({ accessToken: 'graph-token' });
  vi.mocked(verifyCallerIdentity).mockResolvedValue(CALLER);
});

describe('registration', () => {
  it('registers get-meeting-transcript and reports it', () => {
    const { registered } = harness(null);
    expect(registered).toEqual(['get-meeting-transcript']);
  });

  it('is dark until Note Taker is configured: says so, touches nothing', async () => {
    const { call, makeRequest } = harness(null);
    const { payload, isError } = await call({ threadId: THREAD });
    expect(isError).toBe(true);
    expect(payload.error).toBe('not_configured');
    expect(makeRequest).not.toHaveBeenCalled();
    expect(verifyCallerIdentity).not.toHaveBeenCalled();
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

  it('follows a calendar event to its Teams link', async () => {
    const makeRequest = vi
      .fn()
      .mockResolvedValue({ subject: 'Sync', onlineMeeting: { joinUrl: JOIN_URL } });
    const out = await resolveThreadId({ makeRequest }, { eventId: 'AAMk=' });
    expect(out).toEqual({ threadId: THREAD, via: 'event' });
    expect(makeRequest.mock.calls[0][0]).toBe(
      '/me/events/AAMk%3D?$select=subject,isOnlineMeeting,onlineMeeting'
    );
  });

  it('says when a calendar event is not a Teams meeting', async () => {
    const makeRequest = vi.fn().mockResolvedValue({ subject: 'Lunch' });
    await expect(resolveThreadId({ makeRequest }, { eventId: 'x' })).rejects.toThrow(
      /"Lunch" has no Teams meeting link/
    );
  });

  it('asks for an identifier when given none', async () => {
    await expect(resolveThreadId({ makeRequest: vi.fn() }, {})).rejects.toThrow(
      /one of threadId, meetingUrl or eventId/
    );
  });
});

describe('get-meeting-transcript', () => {
  it('sends the verified identity — not the token claims — to Note Taker', async () => {
    const nt = stubNoteTaker(async () => transcriptOf('hello'));
    const { call } = harness(nt);
    const { payload, isError } = await call({ threadId: THREAD });

    expect(isError).toBeUndefined();
    expect(verifyCallerIdentity).toHaveBeenCalledWith('graph-token');
    expect(nt.getTranscript).toHaveBeenCalledWith(THREAD, CALLER);
    expect(payload).toMatchObject({
      threadId: THREAD,
      subject: 'IT Controls discussion',
      resolvedVia: 'threadId',
      totalChars: 5,
      truncated: false,
      transcript: { available: true, content: 'hello' },
    });
  });

  it('refuses to call Note Taker when identity cannot be verified', async () => {
    vi.mocked(verifyCallerIdentity).mockRejectedValue(
      new IdentityVerificationError('nope', 'graph_rejected')
    );
    const nt = stubNoteTaker(async () => transcriptOf('x'));
    const { call } = harness(nt);
    const { payload, isError } = await call({ threadId: THREAD });

    expect(isError).toBe(true);
    expect(payload).toMatchObject({ error: 'identity_unverified', reason: 'graph_rejected' });
    expect(nt.getTranscript).not.toHaveBeenCalled();
  });

  it('reports a missing caller credential', async () => {
    vi.mocked(getRequestTokens).mockReturnValue(undefined);
    const nt = stubNoteTaker(async () => transcriptOf('x'));
    const { payload, isError } = await harness(nt).call({ threadId: THREAD });
    expect(isError).toBe(true);
    expect(payload.error).toBe('no_caller');
  });

  it('returns an unresolved_meeting error without calling Note Taker', async () => {
    const nt = stubNoteTaker(async () => transcriptOf('x'));
    const { payload, isError } = await harness(nt).call({ threadId: 'garbage' });
    expect(isError).toBe(true);
    expect(payload.error).toBe('unresolved_meeting');
    expect(nt.getTranscript).not.toHaveBeenCalled();
  });

  it.each([
    ['not_participant', 403],
    ['transcripts_disabled', 403],
    ['meeting_not_found', 404],
    ['disabled', 503],
    ['unauthorized', 401],
  ] as const)('relays a %s refusal as a structured error', async (refusal, status) => {
    const nt = stubNoteTaker(async () => {
      throw new NoteTakerError('why', refusal, status);
    });
    const { payload, isError } = await harness(nt).call({ meetingUrl: RECAP_URL });
    expect(isError).toBe(true);
    expect(payload.error).toBe(refusal);
    expect(payload.message).toBe('why');
    expect(typeof payload.note).toBe('string');
  });

  it('carries retryAfterSeconds on a rate limit', async () => {
    const nt = stubNoteTaker(async () => {
      throw new NoteTakerError('slow down', 'rate_limited', 429, 45);
    });
    const { payload } = await harness(nt).call({ threadId: THREAD });
    expect(payload).toMatchObject({ error: 'rate_limited', retryAfterSeconds: 45 });
  });

  it('caps long transcripts at maxChars and says so', async () => {
    const nt = stubNoteTaker(async () => transcriptOf('x'.repeat(1200)));
    const { payload } = await harness(nt).call({ threadId: THREAD, maxChars: 500 });
    expect(payload.transcript.content).toHaveLength(500);
    expect(payload).toMatchObject({ totalChars: 1200, truncated: true });
    expect(payload.note).toMatch(/cut at 500 of 1200/);
  });

  it('defaults the cap to DEFAULT_MAX_CHARS', async () => {
    const nt = stubNoteTaker(async () => transcriptOf('y'.repeat(DEFAULT_MAX_CHARS + 1)));
    const { payload } = await harness(nt).call({ threadId: THREAD });
    expect(payload.transcript.content).toHaveLength(DEFAULT_MAX_CHARS);
    expect(payload.truncated).toBe(true);
  });

  it('passes the no-transcript shape through with its reason', async () => {
    const nt = stubNoteTaker(async () => ({
      threadId: THREAD,
      transcript: { available: false, reason: 'not_recorded' },
    }));
    const { payload, isError } = await harness(nt).call({ threadId: THREAD });
    expect(isError).toBeUndefined();
    expect(payload.transcript).toEqual({ available: false, reason: 'not_recorded' });
    expect(payload.truncated).toBeUndefined();
  });
});
