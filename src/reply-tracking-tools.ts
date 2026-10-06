import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type GraphClient from './graph-client.js';
import logger from './logger.js';
import { withUsageLog } from './usage-log.js';
import { getRequestActor } from './request-context.js';

/**
 * Reply tracking (tsq.27): "I sent this — tell me later whether anyone answered."
 *
 * watch-for-reply (notification-tools.ts) answers the same question only while a
 * session is open to wait; its watches live in this server's memory and die with
 * a restart. This module answers it across conversations by storing NOTHING here.
 * The tracking state lives in the user's own mailbox, as single-value extended
 * properties on the SENT message (the Sent Items copy). Any later session, on any
 * replica, after any restart, reads it back with the user's own token.
 *
 * Two properties, not one, so the list query can use the documented
 * `ep/value eq '<literal>'` filter form:
 *   - STATE: 'awaiting' while tracked, 'done' once untracked. Graph cannot delete a
 *     single-value extended property short of deleting the item, so untracking
 *     overwrites the state rather than removing it.
 *   - META: small JSON — trackedAt, optional dueBy, optional waitingOn addresses.
 *     No subject, body or note is stored: the message already holds its content.
 */

/** Our MAPI named-property namespace. Fixed forever: changing it orphans every tracked message. */
export const TRACKING_NAMESPACE = '{6f0d3c1e-8a4b-4f7e-9c2d-27a5b3e1d904}';
export const STATE_PROPERTY_ID = `String ${TRACKING_NAMESPACE} Name TsqAwaitingReply`;
export const META_PROPERTY_ID = `String ${TRACKING_NAMESPACE} Name TsqAwaitingReplyMeta`;
export const STATE_AWAITING = 'awaiting';
export const STATE_DONE = 'done';

/** Newest tracked messages examined per list call; each costs one conversation lookup. */
export const LIST_CAP = 25;
/** Tracked messages fetched before the client-side sort that picks the newest LIST_CAP. */
const LIST_FETCH = 100;
/** Inbound messages fetched per conversation. */
const REPLIES_PER_THREAD = 10;
/** Graph's ceiling on sub-requests in one $batch. */
const BATCH_MAX = 20;
/** Candidates fetched when finding a sent message by subject/recipient. */
const SEARCH_TOP = 5;

export interface TrackingMeta {
  v: 1;
  trackedAt: string;
  dueBy?: string;
  waitingOn?: string[];
}

interface Address {
  emailAddress?: { address?: string; name?: string };
}

interface MailMessage {
  id?: string;
  subject?: string;
  sentDateTime?: string;
  receivedDateTime?: string;
  conversationId?: string;
  internetMessageId?: string;
  parentFolderId?: string;
  isDraft?: boolean;
  from?: Address;
  sender?: Address;
  toRecipients?: Address[];
  ccRecipients?: Address[];
  bccRecipients?: Address[];
  singleValueExtendedProperties?: { id?: string; value?: string }[];
}

export interface RegisterHooks {
  isToolEnabled: (name: string) => boolean;
  readOnly: boolean;
  push: (name: string) => void;
  fail: (name: string, error: Error) => void;
  skip: (name: string) => void;
}

function jsonResult(payload: unknown, isError = false) {
  return {
    content: [{ type: 'text' as const, text: JSON.stringify(payload) }],
    ...(isError ? { isError: true } : {}),
  };
}

/** OData string literal: single quotes doubled. */
export function odataString(value: string): string {
  return `'${value.replace(/'/g, "''")}'`;
}

/** Query string from raw values; every value URL-encoded so OData punctuation survives. */
function query(params: Record<string, string | undefined>): string {
  const parts = Object.entries(params)
    .filter(([, v]) => v !== undefined)
    .map(([k, v]) => `${k}=${encodeURIComponent(v as string)}`);
  return parts.length ? `?${parts.join('&')}` : '';
}

const SENT_SELECT =
  'id,subject,sentDateTime,conversationId,internetMessageId,parentFolderId,isDraft,' +
  'from,sender,toRecipients,ccRecipients,bccRecipients';

/** $expand clause that brings back only our two properties. */
export const EXPAND_TRACKING =
  `singleValueExtendedProperties($filter=id eq ${odataString(STATE_PROPERTY_ID)}` +
  ` or id eq ${odataString(META_PROPERTY_ID)})`;

/** $filter selecting messages whose tracking state is 'awaiting' (the documented eq form). */
export const AWAITING_FILTER =
  `singleValueExtendedProperties/any(ep: ep/id eq ${odataString(STATE_PROPERTY_ID)}` +
  ` and ep/value eq ${odataString(STATE_AWAITING)})`;

/** $filter for inbound messages in a conversation after a point in time. */
export function repliesFilter(conversationId: string, after: string): string {
  return `conversationId eq ${odataString(conversationId)} and receivedDateTime gt ${after}`;
}

/** Epoch ms; 0 for missing or unparseable, so such items sort oldest. */
function ts(value: string | undefined): number {
  const t = Date.parse(value ?? '');
  return Number.isNaN(t) ? 0 : t;
}

function addr(a?: Address): string {
  return (a?.emailAddress?.address ?? '').trim();
}

function recipients(msg: MailMessage): string[] {
  return [...(msg.toRecipients ?? []), ...(msg.ccRecipients ?? []), ...(msg.bccRecipients ?? [])]
    .map(addr)
    .filter(Boolean);
}

function prop(msg: MailMessage, id: string): string | undefined {
  return msg.singleValueExtendedProperties?.find((p) => p.id?.toLowerCase() === id.toLowerCase())
    ?.value;
}

export function parseMeta(raw: string | undefined): Partial<TrackingMeta> {
  if (!raw) return {};
  try {
    const parsed = JSON.parse(raw) as Partial<TrackingMeta>;
    return parsed && typeof parsed === 'object' ? parsed : {};
  } catch {
    return {};
  }
}

function normalizeIso(value: string | undefined, field: string): string | { error: string } {
  if (value === undefined) return '';
  const t = Date.parse(value);
  if (Number.isNaN(t)) return { error: `${field} is not a date: '${value}'. Use ISO 8601.` };
  return new Date(t).toISOString();
}

/**
 * Whether an inbound message counts as a reply. Pure; exported for tests.
 * Excluded: the tracked message itself, drafts, anything from the user's own
 * addresses, and — when waitingOn is set — anyone not on that list.
 */
export function isQualifyingReply(
  reply: MailMessage,
  opts: { trackedId?: string; selfAddresses: Set<string>; waitingOn?: string[] }
): boolean {
  if (!reply.id || reply.id === opts.trackedId) return false;
  if (reply.isDraft) return false;
  const from = addr(reply.from).toLowerCase();
  if (!from || opts.selfAddresses.has(from)) return false;
  if (opts.waitingOn?.length) {
    return opts.waitingOn.some((w) => w.toLowerCase() === from);
  }
  return true;
}

function describeSent(msg: MailMessage) {
  return {
    messageId: msg.id,
    subject: msg.subject,
    sentDateTime: msg.sentDateTime,
    to: recipients(msg),
  };
}

const SENT_ITEMS_HINT =
  'Pass the id of the message as it sits in your Sent Items (list-mail-folder-messages with ' +
  "mailFolderId 'sentitems'), its internetMessageId, or a subject / recipient to find it.";

export function registerReplyTrackingTools(
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

  /** Finds the sent message to track, or explains why it cannot. */
  async function resolveSent(params: {
    messageId?: string;
    internetMessageId?: string;
    subject?: string;
    recipient?: string;
  }): Promise<
    { message: MailMessage; alsoMatched?: ReturnType<typeof describeSent>[] } | { error: string }
  > {
    const messageId = params.messageId?.trim();
    const internetMessageId = params.internetMessageId?.trim();
    const subject = params.subject?.trim();
    const recipient = params.recipient?.trim();

    if (messageId) {
      const [msg, folder] = await Promise.all([
        graphClient.makeRequest(
          `/me/messages/${encodeURIComponent(messageId)}${query({ $select: SENT_SELECT })}`
        ) as Promise<MailMessage>,
        graphClient.makeRequest(`/me/mailFolders/sentitems${query({ $select: 'id' })}`) as Promise<{
          id?: string;
        }>,
      ]);
      if (msg?.isDraft) {
        return { error: 'That message is a draft — send it first, then track the sent copy.' };
      }
      if (folder?.id && msg?.parentFolderId && msg.parentFolderId !== folder.id) {
        return {
          error:
            'That message is not in your Sent Items, so it is not one you sent (or it was ' +
            `moved). ${SENT_ITEMS_HINT}`,
        };
      }
      return { message: msg };
    }

    let found: MailMessage[];
    if (internetMessageId) {
      const res = (await graphClient.makeRequest(
        `/me/mailFolders/sentitems/messages${query({
          $filter: `internetMessageId eq ${odataString(internetMessageId)}`,
          $select: SENT_SELECT,
          $top: String(SEARCH_TOP),
        })}`
      )) as { value?: MailMessage[] };
      found = res?.value ?? [];
    } else if (subject || recipient) {
      // KQL in $search; Graph rejects $search with $orderby, so sort client-side.
      const kql = [
        subject ? `subject:"${subject.replace(/"/g, '')}"` : '',
        recipient ? `recipients:"${recipient.replace(/"/g, '')}"` : '',
      ]
        .filter(Boolean)
        .join(' AND ');
      const res = (await graphClient.makeRequest(
        `/me/mailFolders/sentitems/messages${query({
          $search: `"${kql.replace(/"/g, '\\"')}"`,
          $select: SENT_SELECT,
          $top: String(SEARCH_TOP),
        })}`
      )) as { value?: MailMessage[] };
      found = res?.value ?? [];
    } else {
      return { error: `Nothing identifies the message. ${SENT_ITEMS_HINT}` };
    }

    found = found
      .filter((m) => m.id && !m.isDraft)
      .sort((a, b) => ts(b.sentDateTime) - ts(a.sentDateTime));
    if (!found.length) {
      return { error: `No sent message matched. ${SENT_ITEMS_HINT}` };
    }
    // The newest match is tracked. Until tsq.27.1 several threads meant "ambiguous,
    // pick one" — but people ask for "my last email to X", which names the newest,
    // and the first live test spent a round trip on five candidates to get there.
    // The other threads come back as alsoMatched so a wrong guess is one call away.
    const threads = new Set(found.map((m) => m.conversationId ?? m.id));
    if (found.length === 1 || threads.size === 1) return { message: found[0] };
    const others = found
      .slice(1)
      .filter((m) => (m.conversationId ?? m.id) !== (found[0].conversationId ?? found[0].id));
    return { message: found[0], alsoMatched: others.slice(0, 4).map(describeSent) };
  }

  // --- track-reply -------------------------------------------------------------
  register('track-reply', true, () => {
    server.tool(
      'track-reply',
      // Front-loaded with the words people type: follow up, waiting for a reply, track email.
      'Follow up on an email you sent — track it as "waiting for a reply" so a later ' +
        'conversation can ask "did anyone respond?" with list-awaiting-replies. Works across ' +
        'conversations and survives connector restarts: the tracking is stored on your own ' +
        'sent message in your mailbox (Sent Items), not on this server. Unlike ' +
        'watch-for-reply, nothing has to stay open while you wait.\n\n' +
        'Identify the sent message by messageId (from Sent Items), by internetMessageId, or ' +
        'by subject and/or recipient — the most recent matching sent message is tracked, and ' +
        'other matching threads are listed as alsoMatched. Optionally restrict whose reply counts (waitingOn) and set a follow-up ' +
        'date (dueBy) after which it is reported overdue. Tracking the same message again ' +
        'replaces its settings.',
      {
        messageId: z
          .string()
          .optional()
          .describe('The sent message id, as it appears in Sent Items.'),
        internetMessageId: z
          .string()
          .optional()
          .describe("The message's internetMessageId header, e.g. <abc@host>."),
        subject: z
          .string()
          .optional()
          .describe('Find the most recent sent message with this subject (words match).'),
        recipient: z
          .string()
          .optional()
          .describe('Find the most recent sent message to this address (combinable with subject).'),
        waitingOn: z
          .array(z.string().email())
          .max(20)
          .optional()
          .describe(
            'Only replies from these addresses count. Omit to count a reply from anyone but you.'
          ),
        dueBy: z
          .string()
          .optional()
          .describe('Follow-up date (ISO 8601). Past it with no reply, the message shows overdue.'),
      },
      {
        title: 'track-reply',
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: true,
      },
      async (params) =>
        withUsageLog('track-reply', async () => {
          const due = normalizeIso(params.dueBy, 'dueBy');
          if (typeof due === 'object') return jsonResult(due, true);

          let resolved: Awaited<ReturnType<typeof resolveSent>>;
          try {
            resolved = await resolveSent(params);
          } catch (error) {
            return jsonResult(
              { error: `Could not find the sent message: ${(error as Error).message}` },
              true
            );
          }
          if ('error' in resolved) return jsonResult({ error: resolved.error }, true);

          const msg = resolved.message;
          if (!msg?.id) return jsonResult({ error: 'The message came back without an id.' }, true);

          const waitingOn = params.waitingOn?.map((a) => a.trim().toLowerCase()).filter(Boolean);
          const meta: TrackingMeta = {
            v: 1,
            trackedAt: new Date().toISOString(),
            ...(due ? { dueBy: due } : {}),
            ...(waitingOn?.length ? { waitingOn } : {}),
          };

          try {
            await graphClient.makeRequest(`/me/messages/${encodeURIComponent(msg.id)}`, {
              method: 'PATCH',
              body: JSON.stringify({
                singleValueExtendedProperties: [
                  { id: STATE_PROPERTY_ID, value: STATE_AWAITING },
                  { id: META_PROPERTY_ID, value: JSON.stringify(meta) },
                ],
              }),
            });
          } catch (error) {
            return jsonResult(
              { error: `Could not mark the message as tracked: ${(error as Error).message}` },
              true
            );
          }

          const to = recipients(msg).map((a) => a.toLowerCase());
          const notRecipients = (waitingOn ?? []).filter((w) => !to.includes(w));
          return jsonResult({
            tracked: true,
            ...describeSent(msg),
            trackedAt: meta.trackedAt,
            ...(meta.dueBy ? { dueBy: meta.dueBy } : {}),
            waitingOn: meta.waitingOn ?? 'anyone but you',
            ...(notRecipients.length
              ? { warning: `Not recipients of this message: ${notRecipients.join(', ')}.` }
              : {}),
            storedOn:
              'your own sent message (mailbox extended property); nothing is kept on the server',
            next: 'Ask list-awaiting-replies in any later conversation; untrack-reply stops tracking.',
            ...(resolved.alsoMatched?.length
              ? {
                  alsoMatched: resolved.alsoMatched,
                  note:
                    'Tracked the most recent match. If you meant one of alsoMatched, untrack ' +
                    'this one and track that messageId instead.',
                }
              : {}),
          });
        })
    );
  });

  // --- list-awaiting-replies -----------------------------------------------------
  register('list-awaiting-replies', false, () => {
    server.tool(
      'list-awaiting-replies',
      'Did anyone respond? Lists the emails you are tracking as waiting for a reply ' +
        '(see track-reply) and checks each thread for answers — who replied and when, which ' +
        'are still awaiting a reply, and which are overdue for a follow up. Works across ' +
        'conversations and survives restarts, because the tracking is stored on your own ' +
        'sent messages in your mailbox.\n\n' +
        'By default returns only those still awaiting a reply; includeReplied adds the ' +
        'answered ones. since ("replies since Monday") lists only replies received after ' +
        'that time and includes every message that got one. Replies come back as sender, ' +
        'time and messageId — open one with get-mail-message. Your own messages in the ' +
        `thread never count. Examines the ${LIST_CAP} most recently sent tracked messages.`,
      {
        includeReplied: z
          .boolean()
          .optional()
          .describe('Also return tracked messages that have already been answered.'),
        since: z
          .string()
          .optional()
          .describe('Only report replies received after this time (ISO 8601).'),
      },
      {
        title: 'list-awaiting-replies',
        readOnlyHint: true,
        openWorldHint: true,
      },
      async (params) =>
        withUsageLog('list-awaiting-replies', async () => {
          const since = normalizeIso(params.since, 'since');
          if (typeof since === 'object') return jsonResult(since, true);

          let tracked: MailMessage[];
          let moreExist = false;
          try {
            // No $orderby: combined with an extended-property filter Graph may reject
            // it, so the newest are picked client-side.
            // Sent Items only, not the whole mailbox: an extended-property any() filter
            // is an unindexed scan, and across /me/messages it hit 504 then 503
            // ("Cannot query rows in a table") on the first live test. Tracked
            // messages are sent messages, so nothing is lost by narrowing.
            const res = (await withTransientRetry(() =>
              graphClient.makeRequest(
                `/me/mailFolders/sentitems/messages${query({
                  $filter: AWAITING_FILTER,
                  $expand: EXPAND_TRACKING,
                  $select: SENT_SELECT,
                  $top: String(LIST_FETCH),
                })}`
              )
            )) as { value?: MailMessage[]; '@odata.nextLink'?: string };
            tracked = (res?.value ?? []).filter(
              (m) => m.id && (prop(m, STATE_PROPERTY_ID) ?? STATE_AWAITING) === STATE_AWAITING
            );
            moreExist = Boolean(res?.['@odata.nextLink']);
          } catch (error) {
            return jsonResult(
              { error: `Could not list tracked messages: ${(error as Error).message}` },
              true
            );
          }

          tracked.sort((a, b) => ts(b.sentDateTime) - ts(a.sentDateTime));
          const totalSeen = tracked.length;
          const capped = moreExist || tracked.length > LIST_CAP;
          tracked = tracked.slice(0, LIST_CAP);

          const actor = getRequestActor();
          const repliesByMessage = await fetchReplies(tracked);

          const now = Date.now();
          const items = [];
          let awaitingCount = 0;
          let repliedCount = 0;
          for (const msg of tracked) {
            const meta = parseMeta(prop(msg, META_PROPERTY_ID));
            const self = new Set(
              [addr(msg.from), addr(msg.sender), actor?.upn ?? '']
                .map((a) => a.toLowerCase())
                .filter(Boolean)
            );
            const fetched = repliesByMessage.get(msg.id as string);
            if (!fetched || 'error' in fetched) {
              items.push({
                ...describeSent(msg),
                status: 'unknown',
                error: fetched?.error ?? 'Thread could not be checked.',
              });
              continue;
            }
            const replies = fetched.replies
              .filter((r) =>
                isQualifyingReply(r, {
                  trackedId: msg.id,
                  selfAddresses: self,
                  waitingOn: meta.waitingOn,
                })
              )
              .sort((a, b) => ts(a.receivedDateTime) - ts(b.receivedDateTime));
            const status = replies.length ? 'replied' : 'awaiting';
            if (status === 'replied') repliedCount++;
            else awaitingCount++;

            const shown = since
              ? replies.filter((r) => ts(r.receivedDateTime) > Date.parse(since))
              : replies;
            const include =
              status === 'awaiting' || params.includeReplied || (since && shown.length > 0);
            if (!include) continue;

            const repliedBy = new Set(replies.map((r) => addr(r.from).toLowerCase()));
            const overdue =
              status === 'awaiting' && meta.dueBy ? Date.parse(meta.dueBy) < now : undefined;
            items.push({
              ...describeSent(msg),
              status,
              awaitingFrom: meta.waitingOn ?? 'anyone but you',
              ...(meta.waitingOn?.length
                ? { stillWaitingOn: meta.waitingOn.filter((w) => !repliedBy.has(w.toLowerCase())) }
                : {}),
              trackedAt: meta.trackedAt,
              ...(meta.dueBy ? { dueBy: meta.dueBy, overdue: Boolean(overdue) } : {}),
              replies: shown.map((r) => ({
                from: addr(r.from),
                receivedDateTime: r.receivedDateTime,
                messageId: r.id,
              })),
            });
          }

          return jsonResult({
            items,
            summary: {
              examined: tracked.length,
              awaiting: awaitingCount,
              replied: repliedCount,
              ...(since ? { since } : {}),
            },
            ...(capped
              ? {
                  capped: true,
                  note:
                    `Only the ${LIST_CAP} most recently sent tracked messages were checked ` +
                    `(${totalSeen}${moreExist ? '+' : ''} tracked). untrack-reply the answered ` +
                    'ones to bring older ones into view.',
                }
              : {}),
          });
        })
    );
  });

  /**
   * One conversation lookup per tracked message, sent as $batch (20 per envelope).
   * No $orderby: Graph rejects it alongside a conversationId filter; sorted by caller.
   */
  async function fetchReplies(
    tracked: MailMessage[]
  ): Promise<Map<string, { replies: MailMessage[] } | { error: string }>> {
    const out = new Map<string, { replies: MailMessage[] } | { error: string }>();
    const requests = tracked
      .filter((m) => m.conversationId)
      .map((m) => ({
        id: m.id as string,
        method: 'GET',
        url: `/me/messages${query({
          $filter: repliesFilter(
            m.conversationId as string,
            new Date(ts(m.sentDateTime)).toISOString()
          ),
          $select: 'id,from,receivedDateTime,subject,isDraft',
          $top: String(REPLIES_PER_THREAD),
        })}`,
      }));
    for (const m of tracked) {
      if (!m.conversationId) out.set(m.id as string, { error: 'Message has no conversationId.' });
    }

    for (let i = 0; i < requests.length; i += BATCH_MAX) {
      const chunk = requests.slice(i, i + BATCH_MAX);
      try {
        const res = await graphClient.batchRequest(chunk);
        const byId = new Map((res?.responses ?? []).map((r) => [r.id, r]));
        for (const req of chunk) {
          const r = byId.get(req.id);
          if (r && r.status >= 200 && r.status < 300) {
            out.set(req.id, { replies: (r.body as { value?: MailMessage[] })?.value ?? [] });
          } else {
            const code = (r?.body as { error?: { code?: string } } | undefined)?.error?.code;
            out.set(req.id, {
              error: `Thread lookup failed (${r?.status ?? 'no response'}${code ? ` ${code}` : ''}).`,
            });
          }
        }
      } catch (error) {
        logger.warn(`reply-tracking batch failed: ${(error as Error).message}`);
        for (const req of chunk) out.set(req.id, { error: 'Thread lookup failed.' });
      }
    }
    return out;
  }

  // --- untrack-reply -------------------------------------------------------------
  register('untrack-reply', true, () => {
    server.tool(
      'untrack-reply',
      'Stop tracking a sent email for a reply — the follow up is done or no longer needed. ' +
        'It drops out of list-awaiting-replies. Pass the messageId that list-awaiting-replies ' +
        'or track-reply returned. Only the tracking marker on your sent message changes; the ' +
        'email itself is untouched.',
      {
        messageId: z
          .string()
          .min(1)
          .describe('The tracked sent message, from list-awaiting-replies.'),
      },
      {
        title: 'untrack-reply',
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: true,
      },
      async (params) =>
        withUsageLog('untrack-reply', async () => {
          const messageId = params.messageId.trim();
          try {
            // Graph cannot delete a single-value extended property on its own, so the
            // state is overwritten; the list filter only matches 'awaiting'.
            await graphClient.makeRequest(`/me/messages/${encodeURIComponent(messageId)}`, {
              method: 'PATCH',
              body: JSON.stringify({
                singleValueExtendedProperties: [{ id: STATE_PROPERTY_ID, value: STATE_DONE }],
              }),
            });
          } catch (error) {
            return jsonResult(
              { error: `Could not untrack '${messageId}': ${(error as Error).message}` },
              true
            );
          }
          return jsonResult({ untracked: true, messageId });
        })
    );
  });
}

/** Graph's 502/503/504 are worth one more try after a short pause; anything else is not. */
export async function withTransientRetry<T>(fn: () => Promise<T>, delayMs = 2000): Promise<T> {
  try {
    return await fn();
  } catch (error) {
    if (!/\b50[234]\b/.test((error as Error)?.message ?? '')) throw error;
    await new Promise((r) => setTimeout(r, delayMs));
    return fn();
  }
}
