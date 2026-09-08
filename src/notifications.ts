import crypto from 'node:crypto';
import logger from './logger.js';
import usageLogger from './usage-log.js';

/**
 * In-memory state for Microsoft Graph change notifications: a registry of the
 * subscriptions we created, plus a per-user queue of notifications received
 * from Graph but not yet drained by the agent.
 *
 * Why in-memory (v1): the container runs at a fixed single replica
 * (minReplicas = maxReplicas = 1), so there is no cross-replica visibility
 * problem — a notification delivered to the receiver is readable by the same
 * process that serves the user's next tool call. The cost is that a container
 * restart loses the registry and the queues while Graph keeps delivering to the
 * receiver; those deliveries are then rejected as unknown and the user must
 * re-subscribe. A durable store (Redis/Table) is the v2 answer if usage
 * justifies it.
 *
 * Renewal note: this module deliberately holds NO credentials. In HTTP mode the
 * user's access token arrives per request and is never stored, so subscriptions
 * are renewed opportunistically on the owner's next authenticated call rather
 * than by a background timer. See `dueForRenewal`.
 */

/** A subscription we created on behalf of a user, as we track it locally. */
export interface SubscriptionRecord {
  subscriptionId: string;
  /** Stable Entra object id of the owner — the key everything is bucketed by. */
  ownerOid: string;
  ownerUpn?: string;
  /** Graph resource path, e.g. `/me/mailFolders('inbox')/messages`. */
  resource: string;
  /** Short label the user asked for, e.g. `inbox` or `chat:19:...`. */
  friendly: string;
  changeType: string;
  /** Shared secret echoed by Graph in every notification; proves origin. */
  clientState: string;
  /** Epoch ms. */
  expiresAt: number;
  /** Set when renewal failed or Graph told us the subscription is gone. */
  lapsed?: boolean;
  lapsedReason?: string;
}

/**
 * A correspondence watch: "wake me when THIS conversation gets an answer",
 * layered on top of a subscription (tsq.20, Correspondence Watch).
 *
 * A subscription watches a container (an inbox, a chat); a watch narrows that
 * to one correspondence — a mail conversation, or a chat partner — plus an
 * optional sender filter. Matching happens at drain time in the tool layer,
 * because deciding whether a changed item belongs to the watched conversation
 * requires fetching it, and a delegated token only exists during the owner's
 * own call. This module stores the intent; it still holds no credentials.
 */
export interface WatchRecord {
  watchId: string;
  ownerOid: string;
  /** The subscription whose notifications this watch filters. */
  subscriptionId: string;
  kind: 'mail' | 'chat';
  /** Mail: the conversationId a reply must belong to. */
  conversationId?: string;
  /** Chat: the chat whose messages are watched (redundant with the subscription, kept for output). */
  chatId?: string;
  /** Optional sender filter: SMTP address (mail) or user id / display name (chat). */
  fromFilter?: string;
  /** Free text from the caller, echoed verbatim on the wake — the "why was I waiting". */
  note?: string;
  /** Human context captured at creation: mail subject or chat topic. */
  context?: string;
  createdAt: number;
  matchedCount: number;
}

/** One received notification, reduced to a pointer. Never carries content. */
export interface NotificationEntry {
  subscriptionId: string;
  /** Which subscription label it came from, for agent-friendly output. */
  friendly: string;
  resource: string;
  changeType: string;
  /** Id of the changed item, when Graph supplies it in resourceData. */
  resourceId?: string;
  receivedAt: string;
}

/** Per-user queue plus the counters we surface on drain. */
interface UserQueue {
  entries: NotificationEntry[];
  /** Count of entries discarded because the queue hit its cap. */
  dropped: number;
}

interface Waiter {
  resolve: (entries: NotificationEntry[]) => void;
  timer: NodeJS.Timeout;
  /** When set, this waiter only wakes for notifications from that subscription. */
  subscriptionId?: string;
}

/**
 * Cap per user. Notifications are pointers (a few hundred bytes), so 200 is
 * generous for a session while still bounding memory if an agent subscribes to
 * a busy resource and never drains.
 */
const MAX_QUEUE_PER_USER = 200;

const registry = new Map<string, SubscriptionRecord>();
const queues = new Map<string, UserQueue>();
const waiters = new Map<string, Waiter[]>();
const watches = new Map<string, WatchRecord>();

/**
 * Cap per user. A watch is a few hundred bytes of intent; twenty concurrent
 * awaited answers is already an unusual working style, and the cap bounds a
 * runaway agent registering watches in a loop.
 */
const MAX_WATCHES_PER_USER = 20;

/** Notifications received for a subscription we don't know about (post-restart). */
let unknownDeliveries = 0;

function queueFor(ownerOid: string): UserQueue {
  let q = queues.get(ownerOid);
  if (!q) {
    q = { entries: [], dropped: 0 };
    queues.set(ownerOid, q);
  }
  return q;
}

/** Generates the per-subscription clientState secret sent to Graph. */
export function newClientState(): string {
  return crypto.randomBytes(24).toString('base64url');
}

export function registerSubscription(record: SubscriptionRecord): void {
  registry.set(record.subscriptionId, record);
  logger.info(
    `Registered subscription ${record.subscriptionId} (${record.friendly}) for ${
      record.ownerUpn ?? record.ownerOid
    }, expires ${new Date(record.expiresAt).toISOString()}`
  );
}

export function getSubscription(subscriptionId: string): SubscriptionRecord | undefined {
  return registry.get(subscriptionId);
}

/** Subscriptions owned by a user, newest registration first. */
export function listSubscriptions(ownerOid: string): SubscriptionRecord[] {
  return [...registry.values()].filter((r) => r.ownerOid === ownerOid);
}

/** Removes a subscription, but only if the caller owns it. */
export function unregisterSubscription(subscriptionId: string, ownerOid: string): boolean {
  const record = registry.get(subscriptionId);
  if (!record || record.ownerOid !== ownerOid) return false;
  registry.delete(subscriptionId);

  // A watch is meaningless without its subscription — cancel any that rode on it.
  for (const [watchId, w] of watches) {
    if (w.subscriptionId === subscriptionId && w.ownerOid === ownerOid) watches.delete(watchId);
  }

  // Purge anything this subscription already queued. Without this, entries that
  // arrived before the cancel survive in memory and surface on some later drain
  // — which reads exactly like a cancelled subscription still delivering.
  // Observed 2026-08-12: two notifications queued on 08-10 reappeared three days
  // later, after their subscription had been cancelled and confirmed gone.
  const q = queues.get(ownerOid);
  if (q) {
    q.entries = q.entries.filter((e) => e.subscriptionId !== subscriptionId);
  }
  return true;
}

export function markLapsed(subscriptionId: string, reason: string): void {
  const record = registry.get(subscriptionId);
  if (!record) return;
  record.lapsed = true;
  record.lapsedReason = reason;
  logger.warn(`Subscription ${subscriptionId} (${record.friendly}) lapsed: ${reason}`);
}

/** Records a successful renewal. Clears any prior lapsed state. */
export function markRenewed(subscriptionId: string, expiresAt: number): void {
  const record = registry.get(subscriptionId);
  if (!record) return;
  record.expiresAt = expiresAt;
  record.lapsed = false;
  record.lapsedReason = undefined;
}

/**
 * Subscriptions belonging to `ownerOid` that expire within `withinMs`.
 *
 * Called on the owner's own authenticated requests, which is the only moment a
 * usable delegated token exists. The verified Graph maxima make this viable:
 * mail subscriptions last up to ~7 days and Teams chat up to 3 days, so any
 * user active often enough to care about notifications is active often enough
 * to keep them alive.
 */
export function dueForRenewal(ownerOid: string, withinMs: number): SubscriptionRecord[] {
  const threshold = Date.now() + withinMs;
  return [...registry.values()].filter(
    (r) => r.ownerOid === ownerOid && !r.lapsed && r.expiresAt <= threshold
  );
}

/**
 * Accepts a notification from the receiver after clientState validation.
 * Returns false when the subscription is unknown, which happens after a restart
 * and is the caller's cue to log and drop.
 */
export function enqueueNotification(
  subscriptionId: string,
  clientState: string | undefined,
  entry: Omit<NotificationEntry, 'subscriptionId' | 'friendly' | 'receivedAt'>
): boolean {
  const record = registry.get(subscriptionId);
  if (!record) {
    unknownDeliveries++;
    return false;
  }
  // Constant-time compare: clientState is the only thing authenticating an
  // otherwise-public endpoint.
  const expected = Buffer.from(record.clientState);
  const actual = Buffer.from(clientState ?? '');
  if (expected.length !== actual.length || !crypto.timingSafeEqual(expected, actual)) {
    logger.warn(`Rejected notification for ${subscriptionId}: clientState mismatch`);
    return false;
  }

  const full: NotificationEntry = {
    subscriptionId,
    friendly: record.friendly,
    resource: entry.resource,
    changeType: entry.changeType,
    resourceId: entry.resourceId,
    receivedAt: new Date().toISOString(),
  };

  const q = queueFor(record.ownerOid);
  q.entries.push(full);
  if (q.entries.length > MAX_QUEUE_PER_USER) {
    q.entries.splice(0, q.entries.length - MAX_QUEUE_PER_USER);
    q.dropped++;
  }

  usageLogger.info('notification-received', {
    type: 'm365-notification',
    tool: 'graph-notification',
    outcome: 'success',
    subscriptionId,
    friendly: record.friendly,
    changeType: entry.changeType,
    oid: record.ownerOid,
    upn: record.ownerUpn,
  });

  releaseWaiters(record.ownerOid);
  return true;
}

/**
 * Hands queued notifications to any long-poll waiters. A waiter scoped to one
 * subscription only takes entries from it and leaves the rest queued, so a user
 * waiting on a chat doesn't silently consume their own inbox notifications.
 */
function releaseWaiters(ownerOid: string): void {
  const list = waiters.get(ownerOid);
  if (!list?.length) return;
  const remaining: Waiter[] = [];
  for (const w of list) {
    const entries = drainQueue(ownerOid, w.subscriptionId);
    if (entries.length) {
      clearTimeout(w.timer);
      w.resolve(entries);
    } else {
      remaining.push(w);
    }
  }
  if (remaining.length) waiters.set(ownerOid, remaining);
  else waiters.delete(ownerOid);
}

/** Removes and returns queued entries, optionally only those from one subscription. */
function drainQueue(ownerOid: string, subscriptionId?: string): NotificationEntry[] {
  const q = queues.get(ownerOid);
  if (!q || q.entries.length === 0) return [];
  if (!subscriptionId) {
    const entries = q.entries;
    q.entries = [];
    return entries;
  }
  const taken: NotificationEntry[] = [];
  const kept: NotificationEntry[] = [];
  for (const entry of q.entries) {
    (entry.subscriptionId === subscriptionId ? taken : kept).push(entry);
  }
  q.entries = kept;
  return taken;
}

export interface DrainResult {
  notifications: NotificationEntry[];
  /** Entries discarded since the last drain because the queue was full. */
  dropped: number;
  lapsedSubscriptions: Array<{ subscriptionId: string; friendly: string; reason?: string }>;
}

/**
 * Instant drain: returns and clears what is queued for the caller, or only the
 * given subscription's entries when `subscriptionId` is supplied.
 */
export function drain(ownerOid: string, subscriptionId?: string): DrainResult {
  const q = queues.get(ownerOid);
  const dropped = q?.dropped ?? 0;
  // Only clear the drop counter on a full drain — a filtered read shouldn't
  // discard a truncation signal that applies to the whole queue.
  if (q && !subscriptionId) q.dropped = 0;
  return {
    notifications: drainQueue(ownerOid, subscriptionId),
    dropped,
    lapsedSubscriptions: listSubscriptions(ownerOid)
      .filter((r) => r.lapsed)
      .map((r) => ({
        subscriptionId: r.subscriptionId,
        friendly: r.friendly,
        reason: r.lapsedReason,
      })),
  };
}

/**
 * Long-poll: resolves as soon as a notification arrives, or with an empty array
 * at `timeoutMs`. Cheap to call in a loop — an empty cycle costs a few tokens
 * instead of a full list query.
 */
export function waitForNotifications(
  ownerOid: string,
  timeoutMs: number,
  subscriptionId?: string
): Promise<NotificationEntry[]> {
  // Anything already queued returns immediately — the timeout bounds how long
  // we are willing to wait, it does not restrict results to the wait window.
  const queued = drainQueue(ownerOid, subscriptionId);
  if (queued.length) return Promise.resolve(queued);

  return new Promise((resolve) => {
    const timer = setTimeout(() => {
      const list = waiters.get(ownerOid);
      if (list) {
        const remaining = list.filter((w) => w.timer !== timer);
        if (remaining.length) waiters.set(ownerOid, remaining);
        else waiters.delete(ownerOid);
      }
      resolve([]);
    }, timeoutMs);

    const list = waiters.get(ownerOid) ?? [];
    list.push({ resolve, timer, subscriptionId });
    waiters.set(ownerOid, list);
  });
}

// --- correspondence watches (tsq.20) ----------------------------------------

export function newWatchId(): string {
  return `w-${crypto.randomBytes(4).toString('hex')}`;
}

/** Registers a watch. Returns an error string instead of throwing on the cap. */
export function registerWatch(record: WatchRecord): { error?: string } {
  const mine = [...watches.values()].filter((w) => w.ownerOid === record.ownerOid);
  if (mine.length >= MAX_WATCHES_PER_USER) {
    return {
      error:
        `You already have ${mine.length} active watches (the maximum). Cancel one with ` +
        'cancel-watch, or let its subscription expire.',
    };
  }
  watches.set(record.watchId, record);
  logger.info(
    `Registered watch ${record.watchId} (${record.kind}) on subscription ` +
      `${record.subscriptionId} for ${record.ownerOid}`
  );
  return {};
}

export function getWatch(watchId: string, ownerOid: string): WatchRecord | undefined {
  const w = watches.get(watchId);
  return w && w.ownerOid === ownerOid ? w : undefined;
}

/** Watches owned by a user. */
export function listWatches(ownerOid: string): WatchRecord[] {
  return [...watches.values()].filter((w) => w.ownerOid === ownerOid);
}

/** Watches riding on one subscription, for drain-time matching. */
export function watchesForSubscription(subscriptionId: string, ownerOid: string): WatchRecord[] {
  return [...watches.values()].filter(
    (w) => w.subscriptionId === subscriptionId && w.ownerOid === ownerOid
  );
}

/** Removes a watch, but only if the caller owns it. The subscription stays. */
export function cancelWatch(watchId: string, ownerOid: string): boolean {
  const w = watches.get(watchId);
  if (!w || w.ownerOid !== ownerOid) return false;
  watches.delete(watchId);
  return true;
}

export function recordWatchMatch(watchId: string): void {
  const w = watches.get(watchId);
  if (w) w.matchedCount++;
}

/**
 * Puts entries back at the FRONT of the owner's queue without waking waiters.
 *
 * Used by watch-scoped waits: entries drained during the wait that did not
 * match the watch belong to generic consumers, so they are held aside for the
 * duration of the call and restored here. Not waking waiters is what prevents
 * the obvious spin — a requeue that woke the very waiter that is requeueing
 * would drain the same entries forever.
 */
export function requeue(ownerOid: string, entries: NotificationEntry[]): void {
  if (!entries.length) return;
  const q = queueFor(ownerOid);
  q.entries.unshift(...entries);
  if (q.entries.length > MAX_QUEUE_PER_USER) {
    const overflow = q.entries.length - MAX_QUEUE_PER_USER;
    q.entries.splice(MAX_QUEUE_PER_USER, overflow);
    q.dropped += overflow;
  }
}

/** Diagnostics for logging and tests. */
export function stats(): {
  subscriptions: number;
  queuedUsers: number;
  queuedEntries: number;
  unknownDeliveries: number;
} {
  let queuedEntries = 0;
  for (const q of queues.values()) queuedEntries += q.entries.length;
  return {
    subscriptions: registry.size,
    queuedUsers: queues.size,
    queuedEntries,
    unknownDeliveries,
  };
}

/** Test-only reset. */
export function __resetForTests(): void {
  registry.clear();
  queues.clear();
  for (const list of waiters.values()) for (const w of list) clearTimeout(w.timer);
  waiters.clear();
  watches.clear();
  unknownDeliveries = 0;
}
