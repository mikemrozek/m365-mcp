import { describe, it, expect, vi, beforeEach } from 'vitest';
import {
  registerReplyTrackingTools,
  isQualifyingReply,
  repliesFilter,
  parseMeta,
  STATE_PROPERTY_ID,
  META_PROPERTY_ID,
  AWAITING_FILTER,
  EXPAND_TRACKING,
  LIST_CAP,
} from '../src/reply-tracking-tools.js';
import type GraphClient from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn() },
}));

vi.mock('../src/usage-log.js', () => ({
  withUsageLog: (_name: string, fn: () => Promise<unknown>) => fn(),
  logToolUsage: vi.fn(),
  default: { info: vi.fn() },
}));

vi.mock('../src/request-context.js', () => ({
  getRequestActor: () => ({ oid: 'oid-me', upn: 'me@tsq.com' }),
}));

interface Handlers {
  [name: string]: (params: Record<string, unknown>) => Promise<{
    content: { type: 'text'; text: string }[];
    isError?: boolean;
  }>;
}

function harness(client: Partial<GraphClient>, opts: { readOnly?: boolean } = {}) {
  const handlers: Handlers = {};
  const server = {
    tool: (name: string, _d: string, _s: unknown, _a: unknown, handler: Handlers[string]) => {
      handlers[name] = handler;
    },
  } as unknown as Parameters<typeof registerReplyTrackingTools>[0];
  const registered: string[] = [];
  const skipped: string[] = [];
  registerReplyTrackingTools(server, client as GraphClient, {
    isToolEnabled: () => true,
    readOnly: opts.readOnly ?? false,
    push: (n) => registered.push(n),
    fail: (n, e) => {
      throw new Error(`registration failed for ${n}: ${e.message}`);
    },
    skip: (n) => skipped.push(n),
  });
  return { handlers, registered, skipped };
}

const call = async (h: Handlers, name: string, params: Record<string, unknown>) => {
  const res = await h[name](params);
  return { payload: JSON.parse(res.content[0].text), isError: res.isError };
};

const who = (address: string) => ({ emailAddress: { address } });

function sent(id: string, overrides: Record<string, unknown> = {}) {
  return {
    id,
    subject: `Subject ${id}`,
    sentDateTime: '2026-10-01T10:00:00Z',
    conversationId: `conv-${id}`,
    parentFolderId: 'SENT',
    isDraft: false,
    from: who('me@tsq.com'),
    sender: who('me@tsq.com'),
    toRecipients: [who('bob@tsq.com'), who('carol@tsq.com')],
    ...overrides,
  };
}

function tracked(id: string, meta: Record<string, unknown> = {}, overrides = {}) {
  return sent(id, {
    singleValueExtendedProperties: [
      { id: STATE_PROPERTY_ID.toLowerCase(), value: 'awaiting' },
      { id: META_PROPERTY_ID, value: JSON.stringify({ v: 1, trackedAt: 'x', ...meta }) },
    ],
    ...overrides,
  });
}

/** Decodes a query parameter from an endpoint the module built. */
function param(url: string, key: string): string | null {
  return new URL(`https://g${url}`).searchParams.get(key);
}

describe('property ids and filter strings', () => {
  it('uses the String {guid} Name form for both properties', () => {
    expect(STATE_PROPERTY_ID).toMatch(
      /^String \{[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\} Name TsqAwaitingReply$/
    );
    expect(META_PROPERTY_ID).toMatch(/ Name TsqAwaitingReplyMeta$/);
    expect(META_PROPERTY_ID.split(' ')[1]).toBe(STATE_PROPERTY_ID.split(' ')[1]);
  });

  it('filters on the documented any(ep: id eq and value eq) form', () => {
    expect(AWAITING_FILTER).toBe(
      `singleValueExtendedProperties/any(ep: ep/id eq '${STATE_PROPERTY_ID}' and ep/value eq 'awaiting')`
    );
    expect(EXPAND_TRACKING).toBe(
      `singleValueExtendedProperties($filter=id eq '${STATE_PROPERTY_ID}' or id eq '${META_PROPERTY_ID}')`
    );
  });

  it('builds the reply filter with an unquoted datetime and escaped conversationId', () => {
    expect(repliesFilter("ab'c", '2026-10-01T10:00:00.000Z')).toBe(
      "conversationId eq 'ab''c' and receivedDateTime gt 2026-10-01T10:00:00.000Z"
    );
  });

  it('parseMeta tolerates garbage', () => {
    expect(parseMeta('not json')).toEqual({});
    expect(parseMeta(undefined)).toEqual({});
    expect(parseMeta('{"v":1,"dueBy":"d"}')).toEqual({ v: 1, dueBy: 'd' });
  });
});

describe('isQualifyingReply', () => {
  const self = new Set(['me@tsq.com']);
  it('excludes self, drafts and the tracked message itself', () => {
    expect(isQualifyingReply({ id: 'r', from: who('ME@tsq.com') }, { selfAddresses: self })).toBe(
      false
    );
    expect(
      isQualifyingReply(
        { id: 'r', isDraft: true, from: who('bob@tsq.com') },
        { selfAddresses: self }
      )
    ).toBe(false);
    expect(
      isQualifyingReply(
        { id: 't', from: who('bob@tsq.com') },
        { trackedId: 't', selfAddresses: self }
      )
    ).toBe(false);
    expect(isQualifyingReply({ id: 'r', from: who('bob@tsq.com') }, { selfAddresses: self })).toBe(
      true
    );
  });

  it('honours waitingOn case-insensitively', () => {
    const opts = { selfAddresses: self, waitingOn: ['Carol@tsq.com'] };
    expect(isQualifyingReply({ id: 'r', from: who('bob@tsq.com') }, opts)).toBe(false);
    expect(isQualifyingReply({ id: 'r', from: who('carol@TSQ.com') }, opts)).toBe(true);
  });
});

describe('registration', () => {
  it('registers three tools; read-only skips the two writes', () => {
    expect(harness({}).registered).toEqual([
      'track-reply',
      'list-awaiting-replies',
      'untrack-reply',
    ]);
    const ro = harness({}, { readOnly: true });
    expect(ro.registered).toEqual(['list-awaiting-replies']);
    expect(ro.skipped).toEqual(['track-reply', 'untrack-reply']);
  });
});

describe('track-reply', () => {
  let makeRequest: ReturnType<typeof vi.fn>;
  beforeEach(() => {
    makeRequest = vi.fn();
  });
  const h = () => harness({ makeRequest } as unknown as Partial<GraphClient>).handlers;

  it('by messageId: verifies Sent Items, PATCHes both properties, stores no content', async () => {
    makeRequest.mockImplementation(async (url: string, opts?: { method?: string }) => {
      if (url.startsWith('/me/mailFolders/sentitems')) return { id: 'SENT' };
      if (opts?.method === 'PATCH') return { message: 'OK!' };
      return sent('m1');
    });
    const { payload, isError } = await call(h(), 'track-reply', {
      messageId: 'm1',
      waitingOn: ['Bob@tsq.com', 'zed@tsq.com'],
      dueBy: '2026-10-10',
    });
    expect(isError).toBeUndefined();
    expect(payload.tracked).toBe(true);
    expect(payload.dueBy).toBe('2026-10-10T00:00:00.000Z');
    expect(payload.waitingOn).toEqual(['bob@tsq.com', 'zed@tsq.com']);
    expect(payload.warning).toContain('zed@tsq.com');

    const patch = makeRequest.mock.calls.find((c) => c[1]?.method === 'PATCH')!;
    expect(patch[0]).toBe('/me/messages/m1');
    const body = JSON.parse(patch[1].body);
    expect(body.singleValueExtendedProperties[0]).toEqual({
      id: STATE_PROPERTY_ID,
      value: 'awaiting',
    });
    const meta = JSON.parse(body.singleValueExtendedProperties[1].value);
    expect(body.singleValueExtendedProperties[1].id).toBe(META_PROPERTY_ID);
    expect(Object.keys(meta).sort()).toEqual(['dueBy', 'trackedAt', 'v', 'waitingOn']);
    expect(JSON.stringify(meta)).not.toContain('Subject');
  });

  it('refuses a message outside Sent Items, and a draft', async () => {
    makeRequest.mockImplementation(async (url: string) =>
      url.startsWith('/me/mailFolders/sentitems')
        ? { id: 'SENT' }
        : sent('m1', { parentFolderId: 'INBOX' })
    );
    let r = await call(h(), 'track-reply', { messageId: 'm1' });
    expect(r.isError).toBe(true);
    expect(r.payload.error).toContain('not in your Sent Items');

    makeRequest.mockImplementation(async (url: string) =>
      url.startsWith('/me/mailFolders/sentitems') ? { id: 'SENT' } : sent('m1', { isDraft: true })
    );
    r = await call(h(), 'track-reply', { messageId: 'm1' });
    expect(r.payload.error).toContain('draft');
    expect(makeRequest.mock.calls.some((c) => c[1]?.method === 'PATCH')).toBe(false);
  });

  it('by internetMessageId: filters Sent Items with an escaped literal', async () => {
    makeRequest.mockResolvedValueOnce({ value: [sent('m9')] }).mockResolvedValueOnce({});
    const { payload } = await call(h(), 'track-reply', { internetMessageId: "<a'b@host>" });
    expect(payload.messageId).toBe('m9');
    const url = makeRequest.mock.calls[0][0] as string;
    expect(url.startsWith('/me/mailFolders/sentitems/messages?')).toBe(true);
    expect(param(url, '$filter')).toBe("internetMessageId eq '<a''b@host>'");
  });

  it('by subject: uses $search without $orderby and tracks the newest when threads differ', async () => {
    makeRequest.mockResolvedValueOnce({
      value: [
        sent('old', { sentDateTime: '2026-09-01T00:00:00Z', conversationId: 'c1' }),
        sent('new', { sentDateTime: '2026-09-05T00:00:00Z', conversationId: 'c2' }),
      ],
    });
    const { payload, isError } = await call(h(), 'track-reply', {
      subject: 'Budget "Q4"',
      recipient: 'bob@tsq.com',
    });
    expect(isError).toBeUndefined();
    // tsq.27.1: the newest is tracked; the other thread is offered, not demanded.
    expect(payload.tracked).toBe(true);
    expect(payload.messageId).toBe('new');
    expect(payload.alsoMatched.map((c: { messageId: string }) => c.messageId)).toEqual(['old']);
    const url = makeRequest.mock.calls[0][0] as string;
    expect(param(url, '$search')).toBe(
      '"subject:\\"Budget Q4\\" AND recipients:\\"bob@tsq.com\\""'
    );
    expect(param(url, '$orderby')).toBeNull();
    // Search, then the PATCH that marks the newest.
    expect(makeRequest).toHaveBeenCalledTimes(2);
    expect(makeRequest.mock.calls[1][0]).toContain('/me/messages/new');
  });

  it('by subject: several hits in one thread track the newest', async () => {
    makeRequest
      .mockResolvedValueOnce({
        value: [
          sent('a', { sentDateTime: '2026-09-01T00:00:00Z', conversationId: 'c1' }),
          sent('b', { sentDateTime: '2026-09-05T00:00:00Z', conversationId: 'c1' }),
        ],
      })
      .mockResolvedValueOnce({});
    const { payload } = await call(h(), 'track-reply', { subject: 'Budget' });
    expect(payload.tracked).toBe(true);
    expect(payload.messageId).toBe('b');
  });

  it('errors when nothing identifies or matches the message, and on a bad dueBy', async () => {
    let r = await call(h(), 'track-reply', {});
    expect(r.isError).toBe(true);
    makeRequest.mockResolvedValueOnce({ value: [] });
    r = await call(h(), 'track-reply', { subject: 'nope' });
    expect(r.payload.error).toContain('No sent message matched');
    r = await call(h(), 'track-reply', { messageId: 'm', dueBy: 'someday' });
    expect(r.payload.error).toContain('dueBy');
  });
});

describe('list-awaiting-replies', () => {
  let makeRequest: ReturnType<typeof vi.fn>;
  let batchRequest: ReturnType<typeof vi.fn>;
  beforeEach(() => {
    makeRequest = vi.fn();
    batchRequest = vi.fn();
  });
  const h = () =>
    harness({ makeRequest, batchRequest } as unknown as Partial<GraphClient>).handlers;

  /** Answers each batch sub-request from a conversation→replies table. */
  function repliesTable(table: Record<string, unknown[]>) {
    batchRequest.mockImplementation(async (reqs: { id: string; url: string }[]) => ({
      responses: reqs.map((r) => ({
        id: r.id,
        status: 200,
        body: { value: table[r.id] ?? [] },
      })),
    }));
  }

  it('queries with the tracking filter and expand, no $orderby', async () => {
    makeRequest.mockResolvedValueOnce({ value: [] });
    const { payload } = await call(h(), 'list-awaiting-replies', {});
    const url = makeRequest.mock.calls[0][0] as string;
    expect(url.startsWith('/me/mailFolders/sentitems/messages?')).toBe(true);
    expect(param(url, '$filter')).toBe(AWAITING_FILTER);
    expect(param(url, '$expand')).toBe(EXPAND_TRACKING);
    expect(param(url, '$orderby')).toBeNull();
    expect(payload.items).toEqual([]);
    expect(batchRequest).not.toHaveBeenCalled();
  });

  it('classifies awaiting / replied, excludes self-replies, applies waitingOn and overdue', async () => {
    makeRequest.mockResolvedValueOnce({
      value: [
        tracked('a', { dueBy: '2020-01-01T00:00:00.000Z' }),
        tracked('b'),
        tracked('c', { waitingOn: ['carol@tsq.com'] }),
        // Untracked one that slipped past the server filter: skipped.
        sent('d', {
          singleValueExtendedProperties: [{ id: STATE_PROPERTY_ID, value: 'done' }],
        }),
      ],
    });
    repliesTable({
      a: [{ id: 'ra', from: who('me@tsq.com'), receivedDateTime: '2026-10-02T00:00:00Z' }],
      b: [{ id: 'rb', from: who('bob@tsq.com'), receivedDateTime: '2026-10-02T00:00:00Z' }],
      c: [{ id: 'rc', from: who('bob@tsq.com'), receivedDateTime: '2026-10-02T00:00:00Z' }],
    });

    const { payload } = await call(h(), 'list-awaiting-replies', { includeReplied: true });
    const byId = Object.fromEntries(
      payload.items.map((i: { messageId: string }) => [i.messageId, i])
    );
    expect(Object.keys(byId).sort()).toEqual(['a', 'b', 'c']);
    expect(byId.a.status).toBe('awaiting');
    expect(byId.a.overdue).toBe(true);
    expect(byId.b.status).toBe('replied');
    expect(byId.b.replies).toEqual([
      { from: 'bob@tsq.com', receivedDateTime: '2026-10-02T00:00:00Z', messageId: 'rb' },
    ]);
    expect(byId.c.status).toBe('awaiting');
    expect(byId.c.stillWaitingOn).toEqual(['carol@tsq.com']);
    expect(payload.summary).toEqual({ examined: 3, awaiting: 2, replied: 1 });

    const sub = batchRequest.mock.calls[0][0][0];
    expect(sub.method).toBe('GET');
    expect(param(sub.url, '$filter')).toBe(repliesFilter('conv-a', '2026-10-01T10:00:00.000Z'));
    expect(param(sub.url, '$orderby')).toBeNull();
    expect(param(sub.url, '$select')).not.toContain('bodyPreview');
  });

  it('omits replied messages by default', async () => {
    makeRequest.mockResolvedValueOnce({ value: [tracked('a'), tracked('b')] });
    repliesTable({
      b: [{ id: 'rb', from: who('bob@tsq.com'), receivedDateTime: '2026-10-02T00:00:00Z' }],
    });
    const { payload } = await call(h(), 'list-awaiting-replies', {});
    expect(payload.items.map((i: { messageId: string }) => i.messageId)).toEqual(['a']);
    expect(payload.summary.replied).toBe(1);
  });

  it('since: lists only newer replies and includes messages that got one', async () => {
    makeRequest.mockResolvedValueOnce({ value: [tracked('b'), tracked('c')] });
    repliesTable({
      b: [
        { id: 'old', from: who('bob@tsq.com'), receivedDateTime: '2026-10-02T00:00:00Z' },
        { id: 'new', from: who('bob@tsq.com'), receivedDateTime: '2026-10-04T00:00:00Z' },
      ],
      c: [{ id: 'oldc', from: who('bob@tsq.com'), receivedDateTime: '2026-10-02T00:00:00Z' }],
    });
    const { payload } = await call(h(), 'list-awaiting-replies', { since: '2026-10-03' });
    expect(payload.items).toHaveLength(1);
    expect(payload.items[0].messageId).toBe('b');
    expect(payload.items[0].replies.map((r: { messageId: string }) => r.messageId)).toEqual([
      'new',
    ]);
    expect(payload.summary.since).toBe('2026-10-03T00:00:00.000Z');
  });

  it(`caps at ${LIST_CAP} newest, batches in 20s, and says so`, async () => {
    const many = Array.from({ length: 30 }, (_, i) =>
      tracked(`m${i}`, {}, { sentDateTime: new Date(Date.UTC(2026, 8, 1 + i)).toISOString() })
    );
    makeRequest.mockResolvedValueOnce({ value: many });
    repliesTable({});
    const { payload } = await call(h(), 'list-awaiting-replies', {});
    expect(payload.capped).toBe(true);
    expect(payload.items).toHaveLength(LIST_CAP);
    expect(payload.items[0].messageId).toBe('m29');
    expect(payload.note).toContain('30');
    expect(batchRequest).toHaveBeenCalledTimes(2);
    expect(batchRequest.mock.calls[0][0]).toHaveLength(20);
    expect(batchRequest.mock.calls[1][0]).toHaveLength(5);
  });

  it('reports a failed sub-request as unknown rather than failing the call', async () => {
    makeRequest.mockResolvedValueOnce({ value: [tracked('a')] });
    batchRequest.mockResolvedValueOnce({
      responses: [{ id: 'a', status: 400, body: { error: { code: 'BadRequest' } } }],
    });
    const { payload, isError } = await call(h(), 'list-awaiting-replies', {});
    expect(isError).toBeUndefined();
    expect(payload.items[0].status).toBe('unknown');
    expect(payload.items[0].error).toContain('BadRequest');
  });
});

describe('untrack-reply', () => {
  it("overwrites the state with 'done' and leaves the message otherwise alone", async () => {
    const makeRequest = vi.fn().mockResolvedValue({});
    const { handlers } = harness({ makeRequest } as unknown as Partial<GraphClient>);
    const { payload } = await call(handlers, 'untrack-reply', { messageId: 'a/b' });
    expect(payload).toEqual({ untracked: true, messageId: 'a/b' });
    expect(makeRequest).toHaveBeenCalledWith('/me/messages/a%2Fb', {
      method: 'PATCH',
      body: JSON.stringify({
        singleValueExtendedProperties: [{ id: STATE_PROPERTY_ID, value: 'done' }],
      }),
    });
  });

  it('returns an error when Graph refuses', async () => {
    const makeRequest = vi.fn().mockRejectedValue(new Error('404 Not Found'));
    const { handlers } = harness({ makeRequest } as unknown as Partial<GraphClient>);
    const r = await call(handlers, 'untrack-reply', { messageId: 'x' });
    expect(r.isError).toBe(true);
  });
});

describe('withTransientRetry', () => {
  it('retries once on a Graph 503/504, and not on anything else', async () => {
    const { withTransientRetry } = await import('../src/reply-tracking-tools.js');
    const flaky = vi
      .fn()
      .mockRejectedValueOnce(new Error('Microsoft Graph API error: 504 Gateway Timeout'))
      .mockResolvedValueOnce('ok');
    await expect(withTransientRetry(flaky, 0)).resolves.toBe('ok');
    expect(flaky).toHaveBeenCalledTimes(2);

    const bad = vi.fn().mockRejectedValue(new Error('Microsoft Graph API error: 400 Bad Request'));
    await expect(withTransientRetry(bad, 0)).rejects.toThrow('400');
    expect(bad).toHaveBeenCalledTimes(1);
  });
});
