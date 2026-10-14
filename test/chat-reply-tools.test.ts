import { describe, it, expect, vi, beforeEach } from 'vitest';
import {
  registerChatReplyTools,
  judgeChat,
  matchesFrom,
  plainPreview,
  chatActivity,
  type ChatMessage,
} from '../src/chat-reply-tools.js';
import type GraphClient from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn() },
}));

vi.mock('../src/usage-log.js', () => ({
  withUsageLog: (_name: string, fn: () => Promise<unknown>) => fn(),
  logToolUsage: vi.fn(),
  default: { info: vi.fn() },
}));

const actor = vi.hoisted(() => ({
  value: { oid: 'oid-me', upn: 'me@tsq.com' } as object | undefined,
}));
vi.mock('../src/request-context.js', () => ({
  getRequestActor: () => actor.value,
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
  } as unknown as Parameters<typeof registerChatReplyTools>[0];
  const registered: string[] = [];
  registerChatReplyTools(server, client as GraphClient, {
    isToolEnabled: () => true,
    readOnly: opts.readOnly ?? false,
    push: (n) => registered.push(n),
    fail: (n, e) => {
      throw new Error(`registration failed for ${n}: ${e.message}`);
    },
    skip: vi.fn(),
  });
  return { handlers, registered };
}

const call = async (h: Handlers, params: Record<string, unknown> = {}) => {
  const res = await h['list-chat-replies'](params);
  return { payload: JSON.parse(res.content[0].text), isError: res.isError };
};

const ME = { user: { id: 'oid-me', displayName: 'Mike Mrozek' } };
const BOB = { user: { id: 'oid-bob', displayName: 'Bob Builder' } };
const SCOTT = { user: { id: 'oid-scott', displayName: 'Scott Schatz' } };

let seq = 0;
function msg(
  from: ChatMessage['from'],
  at: string,
  overrides: Partial<ChatMessage> = {}
): ChatMessage {
  seq++;
  return {
    id: `m${seq}`,
    messageType: 'message',
    createdDateTime: at,
    deletedDateTime: null,
    from,
    body: { contentType: 'html', content: `<p>message ${seq}</p>` },
    ...overrides,
  };
}

const SINCE = Date.parse('2026-10-07T00:00:00Z');

describe('judgeChat', () => {
  it('awaiting-you when someone else posted after my last message', () => {
    const v = judgeChat(
      [
        msg(BOB, '2026-10-10T10:05:00Z'),
        msg(ME, '2026-10-10T10:00:00Z'),
        msg(SCOTT, '2026-10-10T10:06:00Z'),
      ],
      { meId: 'oid-me', since: SINCE }
    );
    expect(v.status).toBe('awaiting-you');
    expect(v.myLastMessageAt).toBe('2026-10-10T10:00:00Z');
    expect(v.replies.map((r) => r.from?.user?.displayName)).toEqual([
      'Bob Builder',
      'Scott Schatz',
    ]);
    expect(v.others).toEqual(['Scott Schatz', 'Bob Builder']);
  });

  it('answered when I posted after their reply', () => {
    const v = judgeChat(
      [
        msg(ME, '2026-10-10T10:00:00Z'),
        msg(BOB, '2026-10-10T11:00:00Z'),
        msg(ME, '2026-10-10T12:00:00Z'),
      ],
      { meId: 'oid-me', since: SINCE }
    );
    expect(v.status).toBe('answered');
    expect(v.replies).toHaveLength(1);
  });

  it('none when I never posted — an unprompted message is not a reply', () => {
    const v = judgeChat([msg(BOB, '2026-10-10T10:00:00Z')], { meId: 'oid-me', since: SINCE });
    expect(v.status).toBe('none');
    expect(v.others).toEqual(['Bob Builder']);
  });

  it('matches the caller id case-insensitively', () => {
    const v = judgeChat(
      [msg({ user: { id: 'OID-ME' } }, '2026-10-10T10:00:00Z'), msg(BOB, '2026-10-10T11:00:00Z')],
      { meId: 'oid-me', since: SINCE }
    );
    expect(v.status).toBe('awaiting-you');
  });

  it('ignores replies posted before since, even after my message', () => {
    const v = judgeChat([msg(ME, '2026-10-01T10:00:00Z'), msg(BOB, '2026-10-02T10:00:00Z')], {
      meId: 'oid-me',
      since: SINCE,
    });
    expect(v.status).toBe('none');
  });

  it('counts a reply after since to a message of mine from before since', () => {
    const v = judgeChat([msg(ME, '2026-10-01T10:00:00Z'), msg(BOB, '2026-10-08T10:00:00Z')], {
      meId: 'oid-me',
      since: SINCE,
    });
    expect(v.status).toBe('awaiting-you');
  });

  it('ignores system events, deleted messages and bot posts', () => {
    const v = judgeChat(
      [
        msg(ME, '2026-10-10T10:00:00Z'),
        msg(null, '2026-10-10T10:01:00Z', { messageType: 'systemEventMessage' }),
        msg(BOB, '2026-10-10T10:02:00Z', { messageType: 'systemEventMessage' }),
        msg(BOB, '2026-10-10T10:03:00Z', { deletedDateTime: '2026-10-10T10:04:00Z' }),
        msg({ application: { id: 'app', displayName: 'Bot' } }, '2026-10-10T10:05:00Z'),
      ],
      { meId: 'oid-me', since: SINCE }
    );
    expect(v.status).toBe('none');
  });

  it('a deleted later message of mine does not answer the chat', () => {
    const v = judgeChat(
      [
        msg(ME, '2026-10-10T10:00:00Z'),
        msg(BOB, '2026-10-10T11:00:00Z'),
        msg(ME, '2026-10-10T12:00:00Z', { deletedDateTime: '2026-10-10T12:01:00Z' }),
      ],
      { meId: 'oid-me', since: SINCE }
    );
    expect(v.status).toBe('awaiting-you');
  });

  it('with from, only the named people count', () => {
    const msgs = [
      msg(ME, '2026-10-10T10:00:00Z'),
      msg(BOB, '2026-10-10T11:00:00Z'),
      msg(SCOTT, '2026-10-10T12:00:00Z'),
    ];
    const v = judgeChat(msgs, {
      meId: 'oid-me',
      since: SINCE,
      from: [{ input: 'scott', nameTokens: ['scott'] }],
    });
    expect(v.replies.map((r) => r.from?.user?.id)).toEqual(['oid-scott']);
    const none = judgeChat(msgs, {
      meId: 'oid-me',
      since: SINCE,
      from: [{ input: 'Carol', nameTokens: ['carol'] }],
    });
    expect(none.status).toBe('none');
  });
});

describe('matchesFrom', () => {
  it('matches by resolved user id, by every name word, and by email if Graph sends one', () => {
    expect(matchesFrom(SCOTT, [{ input: 'x@tsq.com', userId: 'OID-SCOTT', nameTokens: [] }])).toBe(
      true
    );
    expect(matchesFrom(SCOTT, [{ input: 'Schatz, Scott', nameTokens: ['schatz', 'scott'] }])).toBe(
      true
    );
    expect(matchesFrom(SCOTT, [{ input: 'Scott Smith', nameTokens: ['scott', 'smith'] }])).toBe(
      false
    );
    expect(
      matchesFrom({ user: { id: 'u', userPrincipalName: 'Ann@tsq.com' } }, [
        { input: 'ann@tsq.com', email: 'ann@tsq.com', nameTokens: [] },
      ])
    ).toBe(true);
    expect(matchesFrom({ application: { id: 'a' } }, [{ input: 'a', nameTokens: ['a'] }])).toBe(
      false
    );
  });
});

describe('plainPreview', () => {
  it('strips tags, decodes entities, collapses whitespace and truncates', () => {
    expect(
      plainPreview('<p>Hi&nbsp;<at id="0">Mike</at>,</p><p>sounds &amp; looks &#39;good&#39;</p>')
    ).toBe("Hi Mike, sounds & looks 'good'");
    const long = plainPreview(`<div>${'word '.repeat(60)}</div>`);
    expect(long.length).toBeLessThanOrEqual(120);
    expect(long.endsWith('…')).toBe(true);
    expect(plainPreview(undefined)).toBe('');
  });
});

describe('chatActivity', () => {
  it('takes the newer of last message and last update', () => {
    expect(
      chatActivity({
        lastUpdatedDateTime: '2026-10-01T00:00:00Z',
        lastMessagePreview: { createdDateTime: '2026-10-09T00:00:00Z' },
      })
    ).toBe(Date.parse('2026-10-09T00:00:00Z'));
    expect(chatActivity({})).toBe(0);
  });
});

// --- tool ---------------------------------------------------------------------

function chat(id: string, lastAt: string, overrides: Record<string, unknown> = {}) {
  return {
    id,
    chatType: 'oneOnOne',
    topic: null,
    lastUpdatedDateTime: lastAt,
    lastMessagePreview: { createdDateTime: lastAt },
    ...overrides,
  };
}

/** batchRequest mock answering from a chatId → messages (or status) table. */
function batchFrom(table: Record<string, ChatMessage[] | number>) {
  return vi.fn(async (reqs: { id: string; url: string }[]) => ({
    responses: reqs.map((r) => {
      const chatId = decodeURIComponent(r.url.split('/')[2]);
      const entry = table[chatId];
      if (typeof entry === 'number') {
        return { id: r.id, status: entry, body: { error: { code: 'Forbidden' } } };
      }
      return { id: r.id, status: 200, body: { value: entry ?? [] } };
    }),
  }));
}

describe('list-chat-replies tool', () => {
  beforeEach(() => {
    actor.value = { oid: 'oid-me', upn: 'me@tsq.com' };
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-10-14T12:00:00Z'));
  });

  it('registers as a read-only tool, also in read-only mode', () => {
    expect(harness({}, { readOnly: true }).registered).toEqual(['list-chat-replies']);
  });

  it('lists awaiting chats only by default, with summary counts and previews', async () => {
    const makeRequest = vi.fn(async () => ({
      value: [
        chat('19:a@thread.v2', '2026-10-12T10:00:00Z'),
        chat('19:b@thread.v2', '2026-10-11T10:00:00Z', { topic: 'Budget', chatType: 'group' }),
        chat('19:old@thread.v2', '2026-09-01T10:00:00Z'),
      ],
    }));
    const batchRequest = batchFrom({
      '19:a@thread.v2': [
        msg(ME, '2026-10-12T09:00:00Z'),
        msg(BOB, '2026-10-12T10:00:00Z', { body: { content: '<p>Yes &amp; done</p>' } }),
      ],
      '19:b@thread.v2': [
        msg(ME, '2026-10-11T09:00:00Z'),
        msg(SCOTT, '2026-10-11T09:30:00Z'),
        msg(ME, '2026-10-11T10:00:00Z'),
      ],
    });
    const { handlers } = harness({ makeRequest, batchRequest } as unknown as GraphClient);
    const { payload } = await call(handlers);

    expect(makeRequest).toHaveBeenCalledWith('/me/chats?$expand=lastMessagePreview&$top=50');
    expect(batchRequest).toHaveBeenCalledTimes(1);
    expect(batchRequest.mock.calls[0][0]).toHaveLength(2); // the September chat is not fetched
    expect(payload.summary).toMatchObject({
      chatsActive: 2,
      examined: 2,
      awaitingYou: 1,
      answered: 1,
    });
    expect(payload.summary.since).toBe('2026-10-07T12:00:00.000Z');
    expect(payload.chats).toHaveLength(1);
    expect(payload.chats[0]).toMatchObject({
      chatId: '19:a@thread.v2',
      status: 'awaiting-you',
      with: ['Bob Builder'],
      myLastMessageAt: '2026-10-12T09:00:00Z',
      replyCount: 1,
      replies: [{ from: 'Bob Builder', at: '2026-10-12T10:00:00Z', preview: 'Yes & done' }],
    });
    expect(payload.capped).toBeUndefined();
  });

  it('onlyUnanswered=false includes answered chats, awaiting ones first', async () => {
    const makeRequest = vi.fn(async () => ({
      value: [chat('c1', '2026-10-12T10:00:00Z'), chat('c2', '2026-10-13T10:00:00Z')],
    }));
    const batchRequest = batchFrom({
      c2: [
        msg(ME, '2026-10-13T08:00:00Z'),
        msg(BOB, '2026-10-13T09:00:00Z'),
        msg(ME, '2026-10-13T10:00:00Z'),
      ],
      c1: [msg(ME, '2026-10-12T08:00:00Z'), msg(BOB, '2026-10-12T09:00:00Z')],
    });
    const { handlers } = harness({ makeRequest, batchRequest } as unknown as GraphClient);
    const { payload } = await call(handlers, { onlyUnanswered: false });
    expect(
      payload.chats.map((c: { chatId: string; status: string }) => [c.chatId, c.status])
    ).toEqual([
      ['c1', 'awaiting-you'],
      ['c2', 'answered'],
    ]);
  });

  it('since narrows both the chat list and the replies counted', async () => {
    const makeRequest = vi.fn(async () => ({
      value: [chat('c1', '2026-10-13T10:00:00Z'), chat('c2', '2026-10-12T10:00:00Z')],
    }));
    const batchRequest = batchFrom({
      c1: [
        msg(ME, '2026-10-12T08:00:00Z'),
        msg(BOB, '2026-10-12T09:00:00Z'),
        msg(BOB, '2026-10-13T10:00:00Z'),
      ],
    });
    const { handlers } = harness({ makeRequest, batchRequest } as unknown as GraphClient);
    const { payload } = await call(handlers, { since: '2026-10-13T00:00:00Z' });
    expect(batchRequest.mock.calls[0][0]).toHaveLength(1);
    expect(payload.chats[0].replies).toHaveLength(1);
    expect(payload.chats[0].replies[0].at).toBe('2026-10-13T10:00:00Z');
  });

  it('rejects an unparseable since', async () => {
    const { handlers } = harness({} as GraphClient);
    const { isError, payload } = await call(handlers, { since: 'last tuesday' });
    expect(isError).toBe(true);
    expect(payload.error).toMatch(/ISO 8601/);
  });

  it('from resolves an email to a user id and matches by it', async () => {
    const makeRequest = vi.fn(async (path: string) => {
      if (path.startsWith('/users/')) return { id: 'oid-scott', displayName: 'Scott Schatz' };
      return { value: [chat('c1', '2026-10-12T10:00:00Z'), chat('c2', '2026-10-12T11:00:00Z')] };
    });
    const batchRequest = batchFrom({
      c1: [
        msg(ME, '2026-10-12T08:00:00Z'),
        msg({ user: { id: 'oid-scott', displayName: 'S. Schatz (CIO)' } }, '2026-10-12T09:00:00Z'),
      ],
      c2: [msg(ME, '2026-10-12T08:00:00Z'), msg(BOB, '2026-10-12T09:00:00Z')],
    });
    const { handlers } = harness({ makeRequest, batchRequest } as unknown as GraphClient);
    const { payload } = await call(handlers, { from: ['scott.schatz@tsq.com'] });
    expect(makeRequest).toHaveBeenCalledWith(
      '/users/scott.schatz%40tsq.com?$select=id,displayName'
    );
    expect(payload.chats.map((c: { chatId: string }) => c.chatId)).toEqual(['c1']);
    expect(payload.summary.from).toEqual(['scott.schatz@tsq.com']);
  });

  it('warns when an email cannot be resolved', async () => {
    const makeRequest = vi.fn(async (path: string) => {
      if (path.startsWith('/users/')) throw new Error('404 Request_ResourceNotFound');
      return { value: [] };
    });
    const { handlers } = harness({
      makeRequest,
      batchRequest: batchFrom({}),
    } as unknown as GraphClient);
    const { payload } = await call(handlers, { from: ['ghost@tsq.com'] });
    expect(payload.warning).toMatch(/ghost@tsq.com/);
  });

  it('batches 20 chats per envelope and caps at top, reporting capped', async () => {
    const value = Array.from({ length: 45 }, (_, i) =>
      chat(`c${i}`, new Date(Date.parse('2026-10-13T00:00:00Z') - i * 60_000).toISOString())
    );
    const makeRequest = vi.fn(async () => ({ value }));
    const batchRequest = batchFrom({});
    const { handlers } = harness({ makeRequest, batchRequest } as unknown as GraphClient);

    const { payload } = await call(handlers, { top: 42 });
    expect(batchRequest.mock.calls.map((c) => c[0].length)).toEqual([20, 20, 2]);
    expect(payload.summary).toMatchObject({ chatsActive: 45, examined: 42 });
    expect(payload.capped).toBe(true);
    // The most recently active chats are the ones examined.
    const urls = batchRequest.mock.calls.flatMap((c) => c[0].map((r: { url: string }) => r.url));
    expect(urls[0]).toBe('/chats/c0/messages?$top=50');
    expect(urls).not.toContain('/chats/c44/messages?$top=50');

    batchRequest.mockClear();
    await call(handlers);
    expect(batchRequest.mock.calls.map((c) => c[0].length)).toEqual([20, 5]);
  });

  it('follows nextLink while pages still hold active chats', async () => {
    const makeRequest = vi.fn(async (path: string) => {
      if (path.includes('skiptoken=p2')) {
        return {
          value: [chat('c2', '2026-10-10T00:00:00Z')],
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/chats?$skiptoken=p3',
        };
      }
      if (path.includes('skiptoken=p3'))
        return {
          value: [chat('old', '2026-09-01T00:00:00Z')],
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/chats?$skiptoken=p4',
        };
      return {
        value: [chat('c1', '2026-10-12T00:00:00Z')],
        '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/chats?$skiptoken=p2',
      };
    });
    const { handlers } = harness({
      makeRequest,
      batchRequest: batchFrom({}),
    } as unknown as GraphClient);
    const { payload } = await call(handlers);
    expect(makeRequest).toHaveBeenCalledWith('/me/chats?$skiptoken=p2');
    expect(makeRequest).toHaveBeenCalledWith('/me/chats?$skiptoken=p3');
    // A page with nothing newer than since ends the walk.
    expect(makeRequest).not.toHaveBeenCalledWith('/me/chats?$skiptoken=p4');
    expect(payload.summary.chatsActive).toBe(2);
    expect(payload.capped).toBeUndefined();
  });

  it('reports a chat it cannot read as unknown instead of failing', async () => {
    const makeRequest = vi.fn(async () => ({
      value: [chat('c1', '2026-10-12T00:00:00Z'), chat('c2', '2026-10-12T01:00:00Z')],
    }));
    const batchRequest = batchFrom({
      c1: 403,
      c2: [msg(ME, '2026-10-12T00:00:00Z'), msg(BOB, '2026-10-12T01:00:00Z')],
    });
    const { handlers } = harness({ makeRequest, batchRequest } as unknown as GraphClient);
    const { payload, isError } = await call(handlers);
    expect(isError).toBeUndefined();
    expect(payload.chats.map((c: { status: string }) => c.status)).toEqual([
      'awaiting-you',
      'unknown',
    ]);
    expect(payload.chats[1].error).toMatch(/403 Forbidden/);
    expect(payload.summary.unknown).toBe(1);
  });

  it('a failed batch envelope marks its chats unknown', async () => {
    const makeRequest = vi.fn(async () => ({ value: [chat('c1', '2026-10-12T00:00:00Z')] }));
    const batchRequest = vi.fn(async () => {
      throw new Error('503');
    });
    const { handlers } = harness({ makeRequest, batchRequest } as unknown as GraphClient);
    const { payload } = await call(handlers);
    expect(payload.chats[0].status).toBe('unknown');
  });

  it('chatId checks one chat without listing chats', async () => {
    const makeRequest = vi.fn(async () => ({
      id: '19:x@thread.v2',
      topic: 'Launch',
      chatType: 'group',
    }));
    const batchRequest = batchFrom({
      '19:x@thread.v2': [msg(ME, '2026-10-12T00:00:00Z'), msg(BOB, '2026-10-12T01:00:00Z')],
    });
    const { handlers } = harness({ makeRequest, batchRequest } as unknown as GraphClient);
    const { payload } = await call(handlers, { chatId: '19:x@thread.v2' });
    expect(makeRequest).toHaveBeenCalledTimes(1);
    expect(makeRequest.mock.calls[0][0]).toMatch(/^\/chats\/19%3Ax%40thread\.v2\?/);
    expect(payload.chats[0]).toMatchObject({ topic: 'Launch', status: 'awaiting-you' });
    expect(payload.chats[0].with).toBeUndefined();
  });

  it('falls back to /me for the caller id when there is no request actor', async () => {
    actor.value = undefined;
    const makeRequest = vi.fn(async (path: string) =>
      path.startsWith('/me?') ? { id: 'oid-me' } : { value: [chat('c1', '2026-10-12T00:00:00Z')] }
    );
    const batchRequest = batchFrom({
      c1: [msg(ME, '2026-10-12T00:00:00Z'), msg(BOB, '2026-10-12T01:00:00Z')],
    });
    const { handlers } = harness({ makeRequest, batchRequest } as unknown as GraphClient);
    const { payload } = await call(handlers);
    expect(makeRequest).toHaveBeenCalledWith('/me?$select=id');
    expect(payload.summary.awaitingYou).toBe(1);
  });

  it('shows at most 5 replies per chat, the newest, but counts them all', async () => {
    const makeRequest = vi.fn(async () => ({ value: [chat('c1', '2026-10-12T10:00:00Z')] }));
    const replies = Array.from({ length: 8 }, (_, i) => msg(BOB, `2026-10-12T0${i + 1}:30:00Z`));
    const batchRequest = batchFrom({ c1: [msg(ME, '2026-10-12T00:00:00Z'), ...replies] });
    const { handlers } = harness({ makeRequest, batchRequest } as unknown as GraphClient);
    const { payload } = await call(handlers);
    expect(payload.chats[0].replyCount).toBe(8);
    expect(payload.chats[0].replies).toHaveLength(5);
    expect(payload.chats[0].replies[4].at).toBe('2026-10-12T08:30:00Z');
  });
});
