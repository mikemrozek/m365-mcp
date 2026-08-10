import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type GraphClient from './graph-client.js';
import logger from './logger.js';
import { getRequestActor, getRequestTokens } from './request-context.js';
import { withUsageLog } from './usage-log.js';
import {
  drain,
  dueForRenewal,
  listSubscriptions,
  markLapsed,
  markRenewed,
  newClientState,
  registerSubscription,
  unregisterSubscription,
  waitForNotifications,
  type SubscriptionRecord,
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
        'nothing can push into the client. Notifications carry ids only, never message content. ' +
        'Subscriptions are held in memory, so after a connector restart or reconnect you must ' +
        're-subscribe. Renewal happens automatically whenever you make any call.',
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

          const base = callbackBase();
          if (!base) {
            return jsonResult(
              {
                error:
                  'No public callback URL available. Microsoft Graph must be able to reach this ' +
                  'server over HTTPS; set MS365_MCP_PUBLIC_URL to the public base URL.',
              },
              true
            );
          }
          if (!base.startsWith('https://')) {
            return jsonResult(
              { error: `Callback URL must be HTTPS; resolved '${base}'.` },
              true
            );
          }

          const clientState = newClientState();
          const minutes = EXPIRY_MINUTES[resolved.family];
          const expirationDateTime = new Date(Date.now() + minutes * 60_000).toISOString();

          try {
            const created = (await graphClient.makeRequest('/subscriptions', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                changeType: changeType?.trim() || 'created',
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
              return jsonResult(
                { error: 'Graph did not return a subscription id.', response: created },
                true
              );
            }

            const record: SubscriptionRecord = {
              subscriptionId: created.id,
              ownerOid: actor.oid,
              ownerUpn: actor.upn,
              resource: resolved.resource,
              friendly: resource.trim(),
              changeType: changeType?.trim() || 'created',
              clientState,
              expiresAt: Date.parse(created.expirationDateTime ?? expirationDateTime),
            };
            registerSubscription(record);

            return jsonResult({
              ...describe(record),
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
        'until something arrives.',
      {},
      {
        title: 'check-notifications',
        readOnlyHint: true,
        openWorldHint: false,
      },
      async () =>
        withUsageLog('check-notifications', async () => {
          const actor = getRequestActor();
          if (!actor?.oid) return jsonResult({ error: NO_IDENTITY }, true);
          await renewDueSubscriptions(graphClient, actor.oid);
          const result = drain(actor.oid);
          return jsonResult({
            count: result.notifications.length,
            notifications: result.notifications,
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
        'instead of a full list query. Returns as soon as anything arrives. Like ' +
        'check-notifications, it returns pointers only — fetch content separately. This only ' +
        'works while the conversation is open; it cannot wake a closed one.',
      {
        timeoutSeconds: z
          .number()
          .int()
          .min(1)
          .max(50)
          .optional()
          .describe('How long to wait before returning empty. Default 45, max 50.'),
      },
      {
        title: 'wait-for-notifications',
        readOnlyHint: true,
        openWorldHint: false,
      },
      async ({ timeoutSeconds }) =>
        withUsageLog('wait-for-notifications', async () => {
          const actor = getRequestActor();
          if (!actor?.oid) return jsonResult({ error: NO_IDENTITY }, true);
          await renewDueSubscriptions(graphClient, actor.oid);
          // Capped below the client's tool-call timeout so a quiet period
          // returns cleanly instead of erroring.
          const seconds = Math.min(timeoutSeconds ?? 45, 50);
          const notifications = await waitForNotifications(actor.oid, seconds * 1000);
          return jsonResult({
            count: notifications.length,
            notifications,
            timedOut: notifications.length === 0,
            waitedSeconds: seconds,
          });
        })
    );
  });
}
