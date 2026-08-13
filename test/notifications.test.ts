import { describe, it, expect, beforeEach } from 'vitest';
import {
  __resetForTests,
  drain,
  dueForRenewal,
  enqueueNotification,
  listSubscriptions,
  markLapsed,
  markRenewed,
  newClientState,
  registerSubscription,
  stats,
  unregisterSubscription,
  waitForNotifications,
  type SubscriptionRecord,
} from '../src/notifications.js';

const OWNER = 'oid-alice';
const OTHER = 'oid-bob';

function makeRecord(overrides: Partial<SubscriptionRecord> = {}): SubscriptionRecord {
  return {
    subscriptionId: 'sub-1',
    ownerOid: OWNER,
    ownerUpn: 'alice@example.com',
    resource: "/me/mailFolders('inbox')/messages",
    friendly: 'inbox',
    changeType: 'created',
    clientState: 'secret-state',
    expiresAt: Date.now() + 60 * 60 * 1000,
    ...overrides,
  };
}

describe('notifications registry', () => {
  beforeEach(() => __resetForTests());

  it('lists only the caller’s own subscriptions', () => {
    registerSubscription(makeRecord());
    registerSubscription(makeRecord({ subscriptionId: 'sub-2', ownerOid: OTHER }));

    expect(listSubscriptions(OWNER).map((r) => r.subscriptionId)).toEqual(['sub-1']);
    expect(listSubscriptions(OTHER).map((r) => r.subscriptionId)).toEqual(['sub-2']);
  });

  it('refuses to unregister another user’s subscription', () => {
    registerSubscription(makeRecord());
    expect(unregisterSubscription('sub-1', OTHER)).toBe(false);
    expect(listSubscriptions(OWNER)).toHaveLength(1);
    expect(unregisterSubscription('sub-1', OWNER)).toBe(true);
    expect(listSubscriptions(OWNER)).toHaveLength(0);
  });

  it('purges queued notifications when a subscription is cancelled', () => {
    registerSubscription(makeRecord());
    registerSubscription(
      makeRecord({ subscriptionId: 'sub-2', friendly: 'chat:x', clientState: 'other-state' })
    );
    enqueueNotification('sub-1', 'secret-state', { resource: 'r', changeType: 'created' });
    enqueueNotification('sub-2', 'other-state', { resource: 'r', changeType: 'created' });

    expect(unregisterSubscription('sub-1', OWNER)).toBe(true);

    // Entries queued before the cancel must not survive to a later drain, where
    // they would look like a cancelled subscription still delivering.
    const result = drain(OWNER);
    expect(result.notifications).toHaveLength(1);
    expect(result.notifications[0].subscriptionId).toBe('sub-2');
  });

  it('generates distinct clientState secrets', () => {
    expect(newClientState()).not.toEqual(newClientState());
  });
});

describe('notification intake', () => {
  beforeEach(() => __resetForTests());

  it('accepts a notification with the matching clientState', () => {
    registerSubscription(makeRecord());
    const ok = enqueueNotification('sub-1', 'secret-state', {
      resource: "Users/x/Messages('AAA')",
      changeType: 'created',
      resourceId: 'AAA',
    });
    expect(ok).toBe(true);

    const result = drain(OWNER);
    expect(result.notifications).toHaveLength(1);
    expect(result.notifications[0]).toMatchObject({
      subscriptionId: 'sub-1',
      friendly: 'inbox',
      resourceId: 'AAA',
      changeType: 'created',
    });
    // Draining clears.
    expect(drain(OWNER).notifications).toHaveLength(0);
  });

  it('rejects a clientState mismatch', () => {
    registerSubscription(makeRecord());
    expect(
      enqueueNotification('sub-1', 'wrong', { resource: 'r', changeType: 'created' })
    ).toBe(false);
    expect(drain(OWNER).notifications).toHaveLength(0);
  });

  it('rejects an unknown subscription and counts it', () => {
    expect(
      enqueueNotification('ghost', 'whatever', { resource: 'r', changeType: 'created' })
    ).toBe(false);
    expect(stats().unknownDeliveries).toBe(1);
  });

  it('caps the queue, dropping oldest and reporting the drop', () => {
    registerSubscription(makeRecord());
    for (let i = 0; i < 205; i++) {
      enqueueNotification('sub-1', 'secret-state', {
        resource: 'r',
        changeType: 'created',
        resourceId: String(i),
      });
    }
    const result = drain(OWNER);
    expect(result.notifications).toHaveLength(200);
    expect(result.dropped).toBeGreaterThan(0);
    // Oldest were discarded, so the newest survives.
    expect(result.notifications.at(-1)?.resourceId).toBe('204');
    // The drop counter resets after being reported.
    expect(drain(OWNER).dropped).toBe(0);
  });

  it('keeps users’ queues separate', () => {
    registerSubscription(makeRecord());
    registerSubscription(
      makeRecord({ subscriptionId: 'sub-2', ownerOid: OTHER, clientState: 'other-state' })
    );
    enqueueNotification('sub-2', 'other-state', { resource: 'r', changeType: 'created' });

    expect(drain(OWNER).notifications).toHaveLength(0);
    expect(drain(OTHER).notifications).toHaveLength(1);
  });
});

describe('long-poll', () => {
  beforeEach(() => __resetForTests());

  it('returns immediately when something is already queued', async () => {
    registerSubscription(makeRecord());
    enqueueNotification('sub-1', 'secret-state', { resource: 'r', changeType: 'created' });
    const entries = await waitForNotifications(OWNER, 5000);
    expect(entries).toHaveLength(1);
  });

  it('resolves early when a notification arrives mid-wait', async () => {
    registerSubscription(makeRecord());
    const started = Date.now();
    const pending = waitForNotifications(OWNER, 5000);
    setTimeout(
      () => enqueueNotification('sub-1', 'secret-state', { resource: 'r', changeType: 'created' }),
      20
    );
    const entries = await pending;
    expect(entries).toHaveLength(1);
    expect(Date.now() - started).toBeLessThan(4000);
  });

  it('times out empty', async () => {
    registerSubscription(makeRecord());
    expect(await waitForNotifications(OWNER, 30)).toEqual([]);
  });
});

describe('per-subscription filtering', () => {
  beforeEach(() => __resetForTests());

  function twoSubscriptions() {
    registerSubscription(makeRecord({ subscriptionId: 'sub-mail', friendly: 'inbox' }));
    registerSubscription(
      makeRecord({ subscriptionId: 'sub-chat', friendly: 'chat:abc', clientState: 'chat-state' })
    );
    enqueueNotification('sub-mail', 'secret-state', { resource: 'm', changeType: 'created' });
    enqueueNotification('sub-chat', 'chat-state', { resource: 'c', changeType: 'created' });
  }

  it('drains only the named subscription and leaves the rest queued', () => {
    twoSubscriptions();

    const chatOnly = drain(OWNER, 'sub-chat');
    expect(chatOnly.notifications).toHaveLength(1);
    expect(chatOnly.notifications[0].subscriptionId).toBe('sub-chat');

    // The mail notification survived the filtered drain.
    const rest = drain(OWNER);
    expect(rest.notifications).toHaveLength(1);
    expect(rest.notifications[0].subscriptionId).toBe('sub-mail');
  });

  it('a scoped wait ignores — and preserves — other subscriptions’ notifications', async () => {
    registerSubscription(makeRecord({ subscriptionId: 'sub-mail', friendly: 'inbox' }));
    registerSubscription(
      makeRecord({ subscriptionId: 'sub-chat', friendly: 'chat:abc', clientState: 'chat-state' })
    );

    const pending = waitForNotifications(OWNER, 120, 'sub-chat');
    // Mail arrives first; it must NOT wake a chat-scoped wait.
    enqueueNotification('sub-mail', 'secret-state', { resource: 'm', changeType: 'created' });
    const entries = await pending;

    expect(entries).toEqual([]);
    // And the mail notification is still there, not silently consumed.
    expect(drain(OWNER).notifications).toHaveLength(1);
  });

  it('a scoped wait wakes for its own subscription', async () => {
    twoSubscriptions();
    const entries = await waitForNotifications(OWNER, 5000, 'sub-mail');
    expect(entries).toHaveLength(1);
    expect(entries[0].subscriptionId).toBe('sub-mail');
  });

  it('a filtered drain preserves the truncation counter', () => {
    registerSubscription(makeRecord());
    for (let i = 0; i < 205; i++) {
      enqueueNotification('sub-1', 'secret-state', { resource: 'r', changeType: 'created' });
    }
    // Filtered read must not swallow the drop signal for the whole queue.
    expect(drain(OWNER, 'other-sub').dropped).toBeGreaterThan(0);
    expect(drain(OWNER).dropped).toBeGreaterThan(0);
  });
});

describe('renewal bookkeeping', () => {
  beforeEach(() => __resetForTests());

  it('reports only subscriptions inside the renewal window', () => {
    registerSubscription(makeRecord({ expiresAt: Date.now() + 60 * 60 * 1000 })); // 1h — due
    registerSubscription(
      makeRecord({ subscriptionId: 'sub-far', expiresAt: Date.now() + 48 * 60 * 60 * 1000 })
    );
    const due = dueForRenewal(OWNER, 12 * 60 * 60 * 1000);
    expect(due.map((r) => r.subscriptionId)).toEqual(['sub-1']);
  });

  it('excludes lapsed subscriptions from renewal and surfaces them on drain', () => {
    registerSubscription(makeRecord({ expiresAt: Date.now() + 1000 }));
    markLapsed('sub-1', 'renewal failed: token expired');

    expect(dueForRenewal(OWNER, 12 * 60 * 60 * 1000)).toHaveLength(0);
    expect(drain(OWNER).lapsedSubscriptions).toEqual([
      { subscriptionId: 'sub-1', friendly: 'inbox', reason: 'renewal failed: token expired' },
    ]);
  });

  it('clears lapsed state on successful renewal', () => {
    registerSubscription(makeRecord({ expiresAt: Date.now() + 1000 }));
    markLapsed('sub-1', 'transient');
    const newExpiry = Date.now() + 24 * 60 * 60 * 1000;
    markRenewed('sub-1', newExpiry);

    expect(drain(OWNER).lapsedSubscriptions).toHaveLength(0);
    expect(listSubscriptions(OWNER)[0].expiresAt).toBe(newExpiry);
  });
});
