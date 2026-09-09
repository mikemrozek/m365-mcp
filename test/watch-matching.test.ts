import { describe, it, expect, beforeEach, vi } from 'vitest';
import { matchWatches } from '../src/notification-tools.js';
import {
  __resetForTests,
  registerSubscription,
  registerWatch,
  type NotificationEntry,
  type SubscriptionRecord,
  type WatchRecord,
} from '../src/notifications.js';
import type GraphClient from '../src/graph-client.js';

const OWNER_OID = 'oid-alice';
const OWNER_UPN = 'alice@example.com';

function sub(overrides: Partial<SubscriptionRecord> = {}): SubscriptionRecord {
  return {
    subscriptionId: 'sub-1',
    ownerOid: OWNER_OID,
    ownerUpn: OWNER_UPN,
    resource: "/me/mailFolders('inbox')/messages",
    friendly: 'inbox',
    changeType: 'created',
    clientState: 's',
    expiresAt: Date.now() + 3600_000,
    ...overrides,
  };
}

function watch(overrides: Partial<WatchRecord> = {}): WatchRecord {
  return {
    watchId: 'w-1',
    ownerOid: OWNER_OID,
    subscriptionId: 'sub-1',
    kind: 'mail',
    conversationId: 'conv-1',
    note: 'why I wait',
    createdAt: Date.now(),
    matchedCount: 0,
    ...overrides,
  };
}

function entry(overrides: Partial<NotificationEntry> = {}): NotificationEntry {
  return {
    subscriptionId: 'sub-1',
    friendly: 'inbox',
    resource: 'r',
    changeType: 'created',
    resourceId: 'msg-1',
    receivedAt: new Date().toISOString(),
    ...overrides,
  };
}

function graphMock(payload: unknown): GraphClient {
  return { makeRequest: vi.fn().mockResolvedValue(payload) } as unknown as GraphClient;
}

function mailMsg(from: string, conversationId = 'conv-1') {
  return {
    conversationId,
    subject: 'watched thread',
    receivedDateTime: '2026-09-09T03:14:36Z',
    bodyPreview: 'hello',
    from: { emailAddress: { address: from } },
  };
}

describe('matchWatches', () => {
  beforeEach(() => __resetForTests());

  it('wakes on a reply from someone else in the watched conversation', async () => {
    registerSubscription(sub());
    registerWatch(watch());
    const wakes = await matchWatches(graphMock(mailMsg('tiffany@example.com')), OWNER_OID, [
      entry(),
    ]);
    expect(wakes).toHaveLength(1);
    expect(wakes[0]).toMatchObject({ watchId: 'w-1', note: 'why I wait', from: 'tiffany@example.com' });
  });

  it('does not wake on a different conversation', async () => {
    registerSubscription(sub());
    registerWatch(watch());
    const wakes = await matchWatches(
      graphMock(mailMsg('tiffany@example.com', 'conv-other')),
      OWNER_OID,
      [entry()]
    );
    expect(wakes).toHaveLength(0);
  });

  it('does not wake on your own mail arriving in your own inbox', async () => {
    // The 09-09 E2E finding: a reply-all or self-send lands your own message
    // in your inbox; without a fromFilter that must not count as an answer.
    registerSubscription(sub());
    registerWatch(watch());
    const wakes = await matchWatches(graphMock(mailMsg(OWNER_UPN)), OWNER_OID, [entry()]);
    expect(wakes).toHaveLength(0);
  });

  it('wakes on your own mail when the filter names you explicitly', async () => {
    registerSubscription(sub());
    registerWatch(watch({ fromFilter: OWNER_UPN }));
    const wakes = await matchWatches(graphMock(mailMsg(OWNER_UPN)), OWNER_OID, [entry()]);
    expect(wakes).toHaveLength(1);
  });

  it('respects a fromFilter naming someone else', async () => {
    registerSubscription(sub());
    registerWatch(watch({ fromFilter: 'tiffany@example.com' }));
    const wrong = await matchWatches(graphMock(mailMsg('bob@example.com')), OWNER_OID, [entry()]);
    expect(wrong).toHaveLength(0);
    const right = await matchWatches(graphMock(mailMsg('Tiffany@Example.com')), OWNER_OID, [
      entry({ resourceId: 'msg-2' }),
    ]);
    expect(right).toHaveLength(1);
  });

  it('chat: skips your own message but honors an explicit self filter', async () => {
    registerSubscription(sub({ subscriptionId: 'sub-chat', friendly: 'chat:c1' }));
    registerWatch(watch({ watchId: 'w-chat', subscriptionId: 'sub-chat', kind: 'chat', chatId: 'c1' }));
    const selfMsg = {
      createdDateTime: '2026-09-09T03:00:00Z',
      body: { content: '<p>me</p>' },
      from: { user: { id: OWNER_OID, displayName: 'Alice' } },
    };
    const silent = await matchWatches(graphMock(selfMsg), OWNER_OID, [
      entry({ subscriptionId: 'sub-chat' }),
    ]);
    expect(silent).toHaveLength(0);

    registerWatch(
      watch({
        watchId: 'w-chat-self',
        subscriptionId: 'sub-chat',
        kind: 'chat',
        chatId: 'c1',
        fromFilter: OWNER_OID,
      })
    );
    const woken = await matchWatches(graphMock(selfMsg), OWNER_OID, [
      entry({ subscriptionId: 'sub-chat', resourceId: 'msg-3' }),
    ]);
    expect(woken.map((w) => w.watchId)).toEqual(['w-chat-self']);
  });

  it('ignores entries whose subscription has no watches', async () => {
    registerSubscription(sub());
    const graph = graphMock(mailMsg('tiffany@example.com'));
    const wakes = await matchWatches(graph, OWNER_OID, [entry()]);
    expect(wakes).toHaveLength(0);
    // And it must not have fetched anything to decide that.
    expect((graph as unknown as { makeRequest: ReturnType<typeof vi.fn> }).makeRequest).not.toHaveBeenCalled();
  });
});
