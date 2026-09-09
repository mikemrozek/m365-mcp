import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type GraphClient from './graph-client.js';
import logger from './logger.js';
import { getRequestActor, getRequestTokens } from './request-context.js';
import { withUsageLog } from './usage-log.js';
import {
  cancelWatch,
  drain,
  dueForRenewal,
  getSubscription,
  getWatch,
  listSubscriptions,
  listWatches,
  markLapsed,
  markRenewed,
  newClientState,
  newWatchId,
  recordWatchMatch,
  registerSubscription,
  registerWatch,
  requeue,
  unregisterSubscription,
  waitForNotifications,
  watchesForSubscription,
  type NotificationEntry,
  type SubscriptionRecord,
  type WatchRecord,
} from './notifications.js';

/**
 * Microsoft Graph change-notification tools.
 *
 * The point of these is token cost: instead of re-listing a mailbox or chat to
 * discover "did anything happen", the agent subscribes once and then drains a
 * queue of pointers. An empty check costs a few tokens; a full list query costs
 * hundreds to thousands.
 *
 * What this does NOT do — stated in every tool description because it is the
 * single most likely misunderstanding: it cannot wake a closed conversation.
 * Nothing can push into claude.ai. Notifications accumulate server-side and are
 * read when the user is present.
 */

/** Subscription expiry we request, per resource family. */
const EXPIRY_MINUTES = {
  // Graph maximum for Outlook message resources is 10,080 minutes (under 7
  // days); back off slightly so clock skew can't push the request over.
  mail: 9900,
  // Graph maximum for Teams chatMessage is 4,320 minutes (3 days).
  teams: 4200,
} as const;

/** Renew anything expiring within this window on the owner's next call. */
const RENEW_WITHIN_MS = 12 * 60 * 60 * 1000;

interface ResolvedResource {
  /** Graph resource path used in the subscription. */
  resource: string;
  family: keyof typeof EXPIRY_MINUTES;
}

/**
 * Maps the friendly shorthand the agent passes into a Graph resource path.
 *
 * Deliberately excludes the tenant-wide `getAllMessages` forms: those are
 * application-permission only (delegated is unsupported) and are the metered
 * Teams surface. Per-chat and per-channel subscriptions avoid both problems.
 */
function resolveResource(input: string): ResolvedResource | { error: string } {
  const value = input.trim();

  if (value === 'inbox') {
    return { resource: "/me/mailFolders('inbox')/messages", family: 'mail' };
  }
  const mailFolder = value.match(/^mail-folder:(.+)$/);
  if (mailFolder) {
    return { resource: `/me/mailFolders('${mailFolder[1]}')/messages`, family: 'mail' };
  }
  const chat = value.match(/^chat:(.+)$/);
  if (chat) {
    return { resource: `/chats/${chat[1]}/messages`, family: 'teams' };
  }
  const channel = value.match(/^channel:([^/]+)\/(.+)$/);
  if (channel) {
    return {
      resource: `/teams/${channel[1]}/channels/${channel[2]}/messages`,
      family: 'teams',
    };
  }
  return {
    error:
      `Unrecognized resource '${input}'. Supported: 'inbox', 'mail-folder:{folderId}', ` +
      `'chat:{chatId}', 'channel:{teamId}/{channelId}'. Tenant-wide subscriptions are not ` +
      `supported (they require application permissions, which this connector does not use).`,
  };
}

/** Public base for the callback URLs Graph will POST to. */
function callbackBase(): string | null {
  const configured = process.env.MS365_MCP_PUBLIC_URL || process.env.MS365_MCP_BASE_URL;
  if (configured) return configured.replace(/\/$/, '');
  const origin = getRequestTokens()?.origin;
  return origin ? origin.replace(/\/$/, '') : null;
}

function jsonResult(payload: unknown, isError = false) {
  return {
    content: [{ type: 'text' as const, text: JSON.stringify(payload) }],
    ...(isError ? { isError: true } : {}),
  };
}

/**
 * Renews any of the caller's subscriptions that are close to expiry, using the
 * token on the current request.
 *
 * This replaces the background timer the design originally called for: in HTTP
 * mode the user's access token exists only for the life of a request and is
 * never stored, so there is no credential a timer could use. Because the expiry
 * windows are days long, "renew whenever the owner does anything" is sufficient
 * — and it keeps the server free of stored user credentials.
 */
async function renewDueSubscriptions(graphClient: GraphClient, ownerOid: string): Promise<void> {
  const due = dueForRenewal(ownerOid, RENEW_WITHIN_MS);
  for (const record of due) {
    const minutes = record.resource.startsWith('/me/mailFolders')
      ? EXPIRY_MINUTES.mail
      : EXPIRY_MINUTES.teams;
    const expirationDateTime = new Date(Date.now() + minutes * 60_000).toISOString();
    try {
      await graphClient.makeRequest(`/subscriptions/${record.subscriptionId}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ expirationDateTime }),
      });
      markRenewed(record.subscriptionId, Date.parse(expirationDateTime));
      logger.info(`Renewed subscription ${record.subscriptionId} (${record.friendly})`);
    } catch (error) {
      markLapsed(record.subscriptionId, `renewal failed: ${(error as Error).message}`);
    }
  }
}

function describe(record: SubscriptionRecord) {
  return {
    subscriptionId: record.subscriptionId,
    resource: record.friendly,
    graphResource: record.resource,
    changeType: record.changeType,
    expiresAt: new Date(record.expiresAt).toISOString(),
    autoRenew: true,
    ...(record.lapsed ? { lapsed: true, lapsedReason: record.lapsedReason } : {}),
  };
}

const NO_IDENTITY =
  'Cannot determine the calling user. Change-notification tools require HTTP/OAuth mode ' +
  'where the caller is identified by their access token.';

/**
 * Creates a Graph subscription and registers it locally. Shared by
 * subscribe-to-changes and watch-for-reply so the two cannot drift.
 */
async function createSubscription(
  graphClient: GraphClient,
  actor: { oid: string; upn?: string },
  resolved: ResolvedResource,
  friendly: string,
  changeType: string
): Promise<{ record: SubscriptionRecord } | { error: string }> {
  const base = callbackBase();
  if (!base) {
    return {
      error:
        'No public callback URL available. Microsoft Graph must be able to reach this ' +
        'server over HTTPS; set MS365_MCP_PUBLIC_URL to the public base URL.',
    };
  }
  if (!base.startsWith('https://')) {
    return { error: `Callback URL must be HTTPS; resolved '${base}'.` };
  }

  const clientState = newClientState();
  const minutes = EXPIRY_MINUTES[resolved.family];
  const expirationDateTime = new Date(Date.now() + minutes * 60_000).toISOString();

  const created = (await graphClient.makeRequest('/subscriptions', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      changeType,
      notificationUrl: `${base}/graph-notifications`,
      // Required by Graph for Teams resources whenever expiry is more
      // than an hour out; harmless and useful elsewhere, since it is
      // how we learn a subscription needs reauthorization or was removed.
      lifecycleNotificationUrl: `${base}/graph-lifecycle`,
      resource: resolved.resource,
      expirationDateTime,
      clientState,
      includeResourceData: false,
    }),
  })) as { id?: string; expirationDateTime?: string };

  if (!created?.id) {
    return { error: 'Graph did not return a subscription id.' };
  }

  const record: SubscriptionRecord = {
    subscriptionId: created.id,
    ownerOid: actor.oid,
    ownerUpn: actor.upn,
    resource: resolved.resource,
    friendly,
    changeType,
    clientState,
    expiresAt: Date.parse(created.expirationDateTime ?? expirationDateTime),
  };
  registerSubscription(record);
  return { record };
}

/**
 * Reuses the caller's live subscription on `friendly` when one exists, else
 * creates one. Watches ride on subscriptions; two watches on the inbox should
 * share one subscription, not race to create duplicates.
 */
async function ensureSubscription(
  graphClient: GraphClient,
  actor: { oid: string; upn?: string },
  friendly: string
): Promise<{ record: SubscriptionRecord } | { error: string }> {
  const existing = listSubscriptions(actor.oid).find(
    (r) => r.friendly === friendly && !r.lapsed && r.expiresAt > Date.now()
  );
  if (existing) return { record: existing };
  const resolved = resolveResource(friendly);
  if ('error' in resolved) return resolved;
  try {
    return await createSubscription(graphClient, actor, resolved, friendly, 'created');
  } catch (error) {
    return { error: (error as Error).message };
  }
}

/**
 * Entries already checked against a watch and found non-matching. Keyed by the
 * entry object itself (entries are requeued by reference), so a watch-scoped
 * wait loop doesn't re-fetch the same inbox noise every cycle. WeakMap so a
 * drained-and-delivered entry costs nothing after it leaves the queue.
 */
const noMatchMemo = new WeakMap<NotificationEntry, Set<string>>();

function memoNoMatch(entry: NotificationEntry, watchId: string): void {
  const set = noMatchMemo.get(entry) ?? new Set<string>();
  set.add(watchId);
  noMatchMemo.set(entry, set);
}

/** What a matched watch hands back: enough to act on without another lookup. */
interface Wake {
  watchId: string;
  kind: WatchRecord['kind'];
  note?: string;
  context?: string;
  /** Id of the matched message; fetch it for the full content. */
  messageId: string;
  from?: string;
  preview?: string;
  receivedAt: string;
}

/** Strips tags and collapses whitespace for a short, safe preview. */
function toPreview(text: string | undefined): string | undefined {
  if (!text) return undefined;
  const flat = text
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return flat ? flat.slice(0, 140) : undefined;
}

/**
 * Decides which drained entries answer a registered watch.
 *
 * Runs at drain time because matching requires fetching the changed item —
 * mail notifications carry only an id, and the delegated token needed to
 * fetch exists only during the owner's own call. A 404 on fetch means the
 * item moved or vanished before we looked: not a match, not an error.
 * Fetches are capped and cached per call.
 */
// Exported for tests; not part of the tool surface.
export async function matchWatches(
  graphClient: GraphClient,
  ownerOid: string,
  entries: NotificationEntry[]
): Promise<Wake[]> {
  const wakes: Wake[] = [];
  const fetchCache = new Map<string, unknown>();
  let fetches = 0;
  const FETCH_CAP = 20;

  for (const entry of entries) {
    if (!entry.resourceId) continue;
    const candidates = watchesForSubscription(entry.subscriptionId, ownerOid);
    if (!candidates.length) continue;
    if (fetches >= FETCH_CAP) break;

    for (const w of candidates) {
      if (noMatchMemo.get(entry)?.has(w.watchId)) continue;
      try {
        if (w.kind === 'mail') {
          const key = `mail:${entry.resourceId}`;
          if (!fetchCache.has(key)) {
            fetches++;
            fetchCache.set(
              key,
              await graphClient.makeRequest(
                `/me/messages/${entry.resourceId}` +
                  `?$select=conversationId,subject,from,receivedDateTime,bodyPreview`
              )
            );
          }
          const msg = fetchCache.get(key) as {
            conversationId?: string;
            subject?: string;
            receivedDateTime?: string;
            bodyPreview?: string;
            from?: { emailAddress?: { address?: string; name?: string } };
          };
          if (msg?.conversationId !== w.conversationId) {
            memoNoMatch(entry, w.watchId);
            continue;
          }
          const sender = msg.from?.emailAddress?.address ?? '';
          if (w.fromFilter && sender.toLowerCase() !== w.fromFilter.toLowerCase()) {
            memoNoMatch(entry, w.watchId);
            continue;
          }
          if (!w.fromFilter) {
            // Your own message arriving in your own inbox — a reply-all on a
            // thread you're also a recipient of, a DL you're on, a self-send —
            // is not an answer, and without this check it woke the watch
            // (found in the 09-09 E2E). Compared by sign-in address, so an
            // alias can still slip through; watching for yourself on purpose
            // is done explicitly with from: <your address>.
            const ownerUpn = getSubscription(entry.subscriptionId)?.ownerUpn ?? '';
            if (ownerUpn && sender.toLowerCase() === ownerUpn.toLowerCase()) {
              memoNoMatch(entry, w.watchId);
              continue;
            }
          }
          recordWatchMatch(w.watchId);
          wakes.push({
            watchId: w.watchId,
            kind: 'mail',
            note: w.note,
            context: w.context ?? msg.subject,
            messageId: entry.resourceId,
            from: sender || undefined,
            preview: toPreview(msg.bodyPreview),
            receivedAt: msg.receivedDateTime ?? entry.receivedAt,
          });
        } else {
          const key = `chat:${w.chatId}:${entry.resourceId}`;
          if (!fetchCache.has(key)) {
            fetches++;
            fetchCache.set(
              key,
              await graphClient.makeRequest(`/chats/${w.chatId}/messages/${entry.resourceId}`)
            );
          }
          const msg = fetchCache.get(key) as {
            createdDateTime?: string;
            body?: { content?: string };
            from?: { user?: { id?: string; displayName?: string } };
          };
          const senderId = msg?.from?.user?.id ?? '';
          const senderName = msg?.from?.user?.displayName ?? '';
          // Your own messages in the chat fire notifications too; a reply-watch
          // must not wake on the very message it is waiting for an answer to.
          // Exception: a fromFilter naming your own id is an explicit request
          // to wake on yourself (mirrors the mail path; used by E2E tests).
          if (!senderId || (senderId === ownerOid && w.fromFilter !== senderId)) {
            memoNoMatch(entry, w.watchId);
            continue;
          }
          if (
            w.fromFilter &&
            senderId !== w.fromFilter &&
            senderName.toLowerCase() !== w.fromFilter.toLowerCase()
          ) {
            memoNoMatch(entry, w.watchId);
            continue;
          }
          recordWatchMatch(w.watchId);
          wakes.push({
            watchId: w.watchId,
            kind: 'chat',
            note: w.note,
            context: w.context,
            messageId: entry.resourceId,
            from: senderName || senderId,
            preview: toPreview(msg?.body?.content),
            receivedAt: msg?.createdDateTime ?? entry.receivedAt,
          });
        }
      } catch (error) {
        // Expected for moved/deleted items; anything else is still not worth
        // failing the drain over — the entry stays visible as a plain
        // notification either way. Memoized so a vanished item is not
        // re-fetched on every wait cycle.
        memoNoMatch(entry, w.watchId);
        logger.info(`watch match fetch skipped for ${entry.resourceId}: ${(error as Error).message}`);
      }
    }
  }
  return wakes;
}

function describeWatch(w: WatchRecord) {
  return {
    watchId: w.watchId,
    kind: w.kind,
    ...(w.conversationId ? { conversationId: w.conversationId } : {}),
    ...(w.chatId ? { chatId: w.chatId } : {}),
    ...(w.fromFilter ? { from: w.fromFilter } : {}),
    ...(w.note ? { note: w.note } : {}),
    ...(w.context ? { context: w.context } : {}),
    subscriptionId: w.subscriptionId,
    createdAt: new Date(w.createdAt).toISOString(),
    matchedCount: w.matchedCount,
  };
}

export interface RegisterHooks {
  isToolEnabled: (name: string) => boolean;
  readOnly: boolean;
  push: (name: string) => void;
  fail: (name: string, error: Error) => void;
  skip: (name: string) => void;
}

export function registerNotificationTools(
  server: McpServer,
  graphClient: GraphClient,
  hooks: RegisterHooks
): void {
  const { isToolEnabled, readOnly, push, fail, skip } = hooks;

  const register = (name: string, isWrite: boolean, registerFn: () => void) => {
    if (isWrite && readOnly) {
      skip(name);
      return;
    }
    if (!isToolEnabled(name)) return;
    try {
      registerFn();
      push(name);
    } catch (error) {
      fail(name, error as Error);
    }
  };

  // --- subscribe-to-changes -------------------------------------------------
  register('subscribe-to-changes', true, () => {
    server.tool(
      'subscribe-to-changes',
      'Subscribes to Microsoft Graph change notifications for a mailbox folder, Teams chat, ' +
        'or Teams channel, so you can react to new items without repeatedly re-listing them.\n\n' +
        'Flow: (1) subscribe-to-changes; (2) check-notifications for an instant drain, or ' +
        'wait-for-notifications to block cheaply until something arrives; (3) fetch content ' +
        'with get-mail-message / get-messages-batch / get-chat-message using the returned ids.\n\n' +
        'IMPORTANT LIMITATIONS: notifications accumulate on the server and are only readable ' +
        'while this conversation is open — this cannot wake a closed conversation, because ' +
        'nothing can push into the client. Notifications carry ids only, never message content, ' +
        'and they are not a guarantee the item still exists — a 404 when you fetch one later ' +
        'is expected, not an error. Subscriptions are held in memory, so after a connector ' +
        'restart or reconnect you must re-subscribe. Renewal happens automatically whenever ' +
        'you make any call.',
      {
        resource: z
          .string()
          .min(1)
          .describe(
            "What to watch: 'inbox', 'mail-folder:{folderId}', 'chat:{chatId}', or " +
              "'channel:{teamId}/{channelId}'. Tenant-wide watching is not supported."
          ),
        changeType: z
          .string()
          .optional()
          .describe("Comma-separated: created, updated, deleted. Defaults to 'created'."),
      },
      {
        title: 'subscribe-to-changes',
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: true,
      },
      async ({ resource, changeType }) =>
        withUsageLog('subscribe-to-changes', async () => {
          const actor = getRequestActor();
          if (!actor?.oid) return jsonResult({ error: NO_IDENTITY }, true);

          const resolved = resolveResource(resource);
          if ('error' in resolved) return jsonResult({ error: resolved.error }, true);

          try {
            const result = await createSubscription(
              graphClient,
              { oid: actor.oid, upn: actor.upn },
              resolved,
              resource.trim(),
              changeType?.trim() || 'created'
            );
            if ('error' in result) return jsonResult({ error: result.error }, true);

            return jsonResult({
              ...describe(result.record),
              note:
                'Notifications accumulate server-side. Drain them with check-notifications, or ' +
                'block cheaply with wait-for-notifications. This does not work across closed ' +
                'conversations, and you must re-subscribe after a connector reconnect.',
            });
          } catch (error) {
            return jsonResult({ error: (error as Error).message }, true);
          }
        })
    );
  });

  // --- list-my-subscriptions ------------------------------------------------
  register('list-my-subscriptions', false, () => {
    server.tool(
      'list-my-subscriptions',
      'Lists the change-notification subscriptions you currently hold on this connector, ' +
        'including when each expires and whether any have lapsed. Subscriptions are held in ' +
        'memory, so this returns nothing after a connector restart even if Microsoft Graph ' +
        'still considers them active.',
      {},
      {
        title: 'list-my-subscriptions',
        readOnlyHint: true,
        openWorldHint: false,
      },
      async () =>
        withUsageLog('list-my-subscriptions', async () => {
          const actor = getRequestActor();
          if (!actor?.oid) return jsonResult({ error: NO_IDENTITY }, true);
          const records = listSubscriptions(actor.oid);
          return jsonResult({ count: records.length, subscriptions: records.map(describe) });
        })
    );
  });

  // --- unsubscribe-from-changes ---------------------------------------------
  register('unsubscribe-from-changes', true, () => {
    server.tool(
      'unsubscribe-from-changes',
      'Cancels a change-notification subscription, both at Microsoft Graph and on this ' +
        'connector. Use list-my-subscriptions to find the subscriptionId.',
      {
        subscriptionId: z.string().min(1).describe('The subscription to cancel.'),
      },
      {
        title: 'unsubscribe-from-changes',
        readOnlyHint: false,
        destructiveHint: true,
        openWorldHint: true,
      },
      async ({ subscriptionId }) =>
        withUsageLog('unsubscribe-from-changes', async () => {
          const actor = getRequestActor();
          if (!actor?.oid) return jsonResult({ error: NO_IDENTITY }, true);

          const id = subscriptionId.trim();
          // Ownership is checked locally first so one user cannot delete
          // another's subscription by guessing an id.
          if (!listSubscriptions(actor.oid).some((r) => r.subscriptionId === id)) {
            return jsonResult(
              { error: `No subscription '${id}' is registered to you on this connector.` },
              true
            );
          }

          try {
            await graphClient.makeRequest(`/subscriptions/${id}`, { method: 'DELETE' });
          } catch (error) {
            // Graph may already have expired it; drop it locally regardless.
            logger.warn(`Graph delete failed for subscription ${id}: ${(error as Error).message}`);
          }
          unregisterSubscription(id, actor.oid);
          return jsonResult({ subscriptionId: id, unsubscribed: true });
        })
    );
  });

  // --- check-notifications --------------------------------------------------
  register('check-notifications', false, () => {
    server.tool(
      'check-notifications',
      'Returns and clears any change notifications queued for you since the last check. ' +
        'Cheap — returns ids and resource pointers only, never message content; follow up with ' +
        'get-mail-message, get-messages-batch, or get-chat-message to read anything. Returns ' +
        'immediately even when empty; use wait-for-notifications instead if you want to block ' +
        'until something arrives.\n\n' +
        'A notification is a record that something changed, NOT a guarantee the item still ' +
        'exists. By the time you fetch it the item may have been moved by a rule or deleted, ' +
        'and a 404 / ErrorItemNotFound on follow-up is an expected outcome, not a failure — ' +
        'skip it and carry on. Pass subscriptionId to drain just one subscription and leave ' +
        'the rest queued.',
      {
        subscriptionId: z
          .string()
          .optional()
          .describe(
            'Only drain notifications from this subscription. Omit to drain everything queued.'
          ),
      },
      {
        title: 'check-notifications',
        readOnlyHint: true,
        openWorldHint: false,
      },
      async ({ subscriptionId }) =>
        withUsageLog('check-notifications', async () => {
          const actor = getRequestActor();
          if (!actor?.oid) return jsonResult({ error: NO_IDENTITY }, true);
          await renewDueSubscriptions(graphClient, actor.oid);
          const result = drain(actor.oid, subscriptionId?.trim() || undefined);
          // Annotate, never suppress: entries answering a registered watch also
          // appear as wakes, carrying the watch's note.
          const wakes = result.notifications.length
            ? await matchWatches(graphClient, actor.oid, result.notifications)
            : [];
          return jsonResult({
            count: result.notifications.length,
            notifications: result.notifications,
            ...(wakes.length ? { wakes } : {}),
            ...(result.dropped ? { droppedOldest: result.dropped } : {}),
            ...(result.lapsedSubscriptions.length
              ? {
                  lapsedSubscriptions: result.lapsedSubscriptions,
                  hint: 'Lapsed subscriptions no longer deliver. Re-subscribe with subscribe-to-changes.',
                }
              : {}),
          });
        })
    );
  });

  // --- wait-for-notifications -----------------------------------------------
  register('wait-for-notifications', false, () => {
    server.tool(
      'wait-for-notifications',
      'Waits until a change notification arrives for you, or until the timeout elapses. Call ' +
        'this in a loop to watch for activity cheaply: each empty cycle costs a few tokens ' +
        'instead of a full list query. Like check-notifications, it returns pointers only — ' +
        'fetch content separately, and treat a 404 on follow-up as expected, since the item ' +
        'may have been moved or deleted since the notification fired. This only works while ' +
        'the conversation is open; it cannot wake a closed one.\n\n' +
        'The timeout is a ceiling on how long this waits, NOT a window on what it returns: ' +
        'anything already queued comes back immediately, including events from before the ' +
        'call. Pass subscriptionId to wait on one subscription only — otherwise a wait ' +
        'intended for a chat will also return, and consume, your queued mail notifications.\n\n' +
        'Pass watchId (from watch-for-reply) for a QUIET wait on one correspondence: the call ' +
        'returns a wake only when the watched conversation is actually answered, and other ' +
        'notifications stay queued for their own consumers. Loop it: each empty cycle costs ' +
        'a few tokens, and a wake arrives carrying the note you set, so you can resume ' +
        'mid-task without re-deriving why you were waiting.',
      {
        timeoutSeconds: z
          .number()
          .int()
          .min(1)
          .max(50)
          .optional()
          .describe('How long to wait before returning empty. Default 45, max 50.'),
        subscriptionId: z
          .string()
          .optional()
          .describe(
            'Only wake for notifications from this subscription; others stay queued. ' +
              'Omit to wait on all of them.'
          ),
        watchId: z
          .string()
          .optional()
          .describe(
            'Quiet mode: only return when this watch (from watch-for-reply) matches a reply. ' +
              'Overrides subscriptionId.'
          ),
      },
      {
        title: 'wait-for-notifications',
        readOnlyHint: true,
        openWorldHint: false,
      },
      async ({ timeoutSeconds, subscriptionId, watchId }) =>
        withUsageLog('wait-for-notifications', async () => {
          const actor = getRequestActor();
          if (!actor?.oid) return jsonResult({ error: NO_IDENTITY }, true);
          await renewDueSubscriptions(graphClient, actor.oid);
          // Capped below the client's tool-call timeout so a quiet period
          // returns cleanly instead of erroring.
          const seconds = Math.min(timeoutSeconds ?? 45, 50);
          // Actual elapsed, so waitedSeconds means what it says: a small value
          // with timedOut:false is a prompt wake; a value near the ceiling is
          // Graph delivery latency, not this server holding the notification.
          const startedAt = Date.now();
          const elapsed = () => Math.round((Date.now() - startedAt) / 1000);

          const watch = watchId?.trim() ? getWatch(watchId.trim(), actor.oid) : undefined;
          if (watchId?.trim() && !watch) {
            return jsonResult(
              {
                error:
                  `No watch '${watchId.trim()}' is registered to you. Watches do not survive ` +
                  'a connector restart — list-watches shows what exists, and watch-for-reply ' +
                  're-creates one.',
              },
              true
            );
          }

          if (!watch) {
            const notifications = await waitForNotifications(
              actor.oid,
              seconds * 1000,
              subscriptionId?.trim() || undefined
            );
            // Annotate, never suppress: a generic drain still returns every
            // entry, and any that answer a registered watch also appear as
            // wakes so the agent need not correlate by hand.
            const wakes = notifications.length
              ? await matchWatches(graphClient, actor.oid, notifications)
              : [];
            return jsonResult({
              count: notifications.length,
              notifications,
              ...(wakes.length ? { wakes } : {}),
              timedOut: notifications.length === 0,
              waitedSeconds: elapsed(),
            });
          }

          // Watch-scoped wait: quiet until the watched correspondence is
          // actually answered. Entries drained along the way that do not match
          // belong to generic consumers — they are held for the duration of the
          // call and put back afterwards, which is also what prevents a
          // requeue-wake spin.
          const deadline = Date.now() + seconds * 1000;
          const held: NotificationEntry[] = [];
          try {
            while (Date.now() < deadline) {
              const entries = await waitForNotifications(
                actor.oid,
                deadline - Date.now(),
                watch.subscriptionId
              );
              if (!entries.length) break; // timed out inside the store
              const wakes = await matchWatches(graphClient, actor.oid, entries);
              const wokenIds = new Set(wakes.map((w) => w.messageId));
              held.push(...entries.filter((e) => !e.resourceId || !wokenIds.has(e.resourceId)));
              if (wakes.length) {
                return jsonResult({
                  wakes,
                  watchId: watch.watchId,
                  timedOut: false,
                  waitedSeconds: elapsed(),
                  note:
                    'The watched correspondence was answered. Fetch the message for full ' +
                    'content; the watch stays active for further replies until cancelled.',
                });
              }
            }
            return jsonResult({
              wakes: [],
              watch: describeWatch(watch),
              timedOut: true,
              waitedSeconds: elapsed(),
              note:
                'No reply yet on the watched correspondence. Call again with the same watchId ' +
                'to keep waiting — each quiet cycle costs a few tokens.',
            });
          } finally {
            requeue(actor.oid, held);
          }
        })
    );
  });

  // --- watch-for-reply ------------------------------------------------------
  register('watch-for-reply', true, () => {
    server.tool(
      'watch-for-reply',
      'Watch one correspondence for an answer — "I just sent this; tell me when it is ' +
        'replied to." Give it the email you sent (messageId) or the Teams chat you wrote in ' +
        '(chatId), optionally who must answer (from), and a short note saying why you are ' +
        'waiting. Then loop wait-for-notifications with the returned watchId: the wait stays ' +
        'quiet through unrelated activity and returns only when that conversation is actually ' +
        'answered, echoing your note so you can resume mid-task.\n\n' +
        'Flow: send-mail / send-chat-message → watch-for-reply → wait-for-notifications ' +
        '(watchId, looped) → on wake, get-mail-message or get-chat-message with the returned ' +
        'id.\n\n' +
        'Works only while a session is open to do the waiting — nothing can wake a closed ' +
        'conversation. Watches live in memory: a connector restart clears them, and ' +
        'list-watches shows what survives.',
      {
        messageId: z
          .string()
          .optional()
          .describe(
            'Watch a mail conversation: the id of a message in the thread — typically the ' +
              'one you just sent or are replying to. Its conversation is what gets watched.'
          ),
        chatId: z
          .string()
          .optional()
          .describe('Watch a Teams chat instead: the chat whose next reply matters.'),
        from: z
          .string()
          .optional()
          .describe(
            'Only wake for this sender. Mail: their SMTP address. Chat: their display name ' +
              'or user id. Omit to wake for any reply that is not your own.'
          ),
        note: z
          .string()
          .max(300)
          .optional()
          .describe(
            'Why you are waiting, in your own words — echoed verbatim on the wake, e.g. ' +
              '"Tiffany, sign-off on the connector rollout".'
          ),
      },
      {
        title: 'watch-for-reply',
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: true,
      },
      async (params) =>
        withUsageLog('watch-for-reply', async () => {
          const actor = getRequestActor();
          if (!actor?.oid) return jsonResult({ error: NO_IDENTITY }, true);

          const messageId = params.messageId?.trim();
          const chatId = params.chatId?.trim();
          if (!messageId === !chatId) {
            return jsonResult(
              { error: 'Pass exactly one of messageId (mail) or chatId (Teams chat).' },
              true
            );
          }

          let watch: WatchRecord;
          if (messageId) {
            let msg: { conversationId?: string; subject?: string };
            try {
              msg = (await graphClient.makeRequest(
                `/me/messages/${messageId}?$select=conversationId,subject`
              )) as { conversationId?: string; subject?: string };
            } catch (error) {
              return jsonResult(
                { error: `Could not read message '${messageId}': ${(error as Error).message}` },
                true
              );
            }
            if (!msg?.conversationId) {
              return jsonResult(
                { error: 'The message has no conversationId; cannot watch its thread.' },
                true
              );
            }
            const sub = await ensureSubscription(
              graphClient,
              { oid: actor.oid, upn: actor.upn },
              'inbox'
            );
            if ('error' in sub) return jsonResult({ error: sub.error }, true);
            watch = {
              watchId: newWatchId(),
              ownerOid: actor.oid,
              subscriptionId: sub.record.subscriptionId,
              kind: 'mail',
              conversationId: msg.conversationId,
              fromFilter: params.from?.trim() || undefined,
              note: params.note?.trim() || undefined,
              context: msg.subject,
              createdAt: Date.now(),
              matchedCount: 0,
            };
          } else {
            // Confirm the chat exists and capture its topic for the wake.
            let topic: string | undefined;
            try {
              const chat = (await graphClient.makeRequest(
                `/chats/${chatId}?$select=id,topic`
              )) as { topic?: string };
              topic = chat?.topic ?? undefined;
            } catch (error) {
              return jsonResult(
                { error: `Could not read chat '${chatId}': ${(error as Error).message}` },
                true
              );
            }
            const sub = await ensureSubscription(
              graphClient,
              { oid: actor.oid, upn: actor.upn },
              `chat:${chatId}`
            );
            if ('error' in sub) return jsonResult({ error: sub.error }, true);
            watch = {
              watchId: newWatchId(),
              ownerOid: actor.oid,
              subscriptionId: sub.record.subscriptionId,
              kind: 'chat',
              chatId,
              fromFilter: params.from?.trim() || undefined,
              note: params.note?.trim() || undefined,
              context: topic,
              createdAt: Date.now(),
              matchedCount: 0,
            };
          }

          const registered = registerWatch(watch);
          if (registered.error) return jsonResult({ error: registered.error }, true);

          return jsonResult({
            ...describeWatch(watch),
            subscriptionExpiresAt: new Date(
              listSubscriptions(actor.oid).find(
                (r) => r.subscriptionId === watch.subscriptionId
              )?.expiresAt ?? Date.now()
            ).toISOString(),
            next:
              `Loop wait-for-notifications with watchId '${watch.watchId}'. Each quiet cycle ` +
              'is cheap; the wake carries your note and the reply’s id.',
          });
        })
    );
  });

  // --- list-watches ---------------------------------------------------------
  register('list-watches', false, () => {
    server.tool(
      'list-watches',
      'Lists your active correspondence watches (from watch-for-reply): what each is ' +
        'watching, its note, and how many replies it has matched. Watches live in memory and ' +
        'do not survive a connector restart.',
      {},
      {
        title: 'list-watches',
        readOnlyHint: true,
        openWorldHint: false,
      },
      async () =>
        withUsageLog('list-watches', async () => {
          const actor = getRequestActor();
          if (!actor?.oid) return jsonResult({ error: NO_IDENTITY }, true);
          const mine = listWatches(actor.oid).map(describeWatch);
          return jsonResult({ count: mine.length, watches: mine });
        })
    );
  });

  // --- cancel-watch ---------------------------------------------------------
  register('cancel-watch', true, () => {
    server.tool(
      'cancel-watch',
      'Cancels one correspondence watch. The underlying subscription stays (other watches ' +
        'or generic notification consumers may share it) — use unsubscribe-from-changes to ' +
        'remove that too.',
      {
        watchId: z.string().min(1).describe('The watch to cancel, from watch-for-reply.'),
      },
      {
        title: 'cancel-watch',
        readOnlyHint: false,
        destructiveHint: true,
        openWorldHint: false,
      },
      async ({ watchId }) =>
        withUsageLog('cancel-watch', async () => {
          const actor = getRequestActor();
          if (!actor?.oid) return jsonResult({ error: NO_IDENTITY }, true);
          const id = watchId.trim();
          if (!cancelWatch(id, actor.oid)) {
            return jsonResult(
              { error: `No watch '${id}' is registered to you on this connector.` },
              true
            );
          }
          return jsonResult({ watchId: id, cancelled: true });
        })
    );
  });
}
