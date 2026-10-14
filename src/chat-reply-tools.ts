import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type GraphClient from './graph-client.js';
import logger from './logger.js';
import { withUsageLog } from './usage-log.js';
import { getRequestActor } from './request-context.js';
import type { RegisterHooks } from './reply-tracking-tools.js';

/**
 * Teams "who replied" (tsq.28): "did anyone answer me in a Teams chat?"
 *
 * Scott Schatz (CIO, 2026-10-14): "I had a session go through the Teams chats my
 * agents posted in over the last week, and six people had replied with nobody
 * noticing." The email answer to that (track-reply, reply-tracking-tools.ts) stores
 * a marker on the sent message; a Teams chat message cannot carry one, and this
 * server stores nothing. So this tool is STATELESS: every call recomputes the answer
 * live from the chats themselves.
 *
 * The rule that makes it work without a marker: in a chat, "a reply to me" is any
 * message from someone else that comes after a message of mine. A chat is
 * 'awaiting-you' when someone posted after my most recent message; 'answered' when
 * a reply came in and I posted again afterwards. An agent posting as the user is the
 * user (from.user.id is the caller's oid), so chats an agent started count too.
 *
 * Cost per call: 1-5 chat-list pages, plus one $batch envelope per 20 chats examined.
 */

/** Default and ceiling for chats examined per call; each is one sub-request. */
export const DEFAULT_TOP = 25;
export const MAX_TOP = 50;
/** Graph's ceiling on sub-requests in one $batch. */
const BATCH_MAX = 20;
/** /me/chats and /chats/{id}/messages both cap $top at 50. */
const PAGE_SIZE = 50;
/** Chat-list pages read before giving up and reporting capped. */
const MAX_CHAT_PAGES = 5;
/** Replies returned per chat; the count is reported in full. */
const REPLIES_SHOWN = 5;
const PREVIEW_CHARS = 120;
const DEFAULT_SINCE_DAYS = 7;

interface ChatIdentity {
  user?: { id?: string; displayName?: string; userPrincipalName?: string; email?: string };
  application?: { id?: string; displayName?: string };
}

export interface ChatMessage {
  id?: string;
  messageType?: string;
  createdDateTime?: string;
  deletedDateTime?: string | null;
  from?: ChatIdentity | null;
  body?: { contentType?: string; content?: string };
  attachments?: unknown[];
}

interface Chat {
  id?: string;
  topic?: string | null;
  chatType?: string;
  webUrl?: string;
  lastUpdatedDateTime?: string;
  lastMessagePreview?: {
    createdDateTime?: string;
    from?: ChatIdentity | null;
  } | null;
}

/** One person the caller asked about, as typed plus whatever the directory said. */
export interface FromFilter {
  input: string;
  /** Entra object id, when an email address resolved through /users. */
  userId?: string;
  /** Lower-cased name tokens to match against from.user.displayName. */
  nameTokens: string[];
  email?: string;
}

/** One chat in the tool's output. */
interface ChatRow {
  chatId?: string;
  topic?: string;
  chatType?: string;
  with?: string[];
  status: 'awaiting-you' | 'answered' | 'unknown';
  error?: string;
  myLastMessageAt?: string;
  replyCount?: number;
  replies?: { from: string; at?: string; messageId?: string; preview: string }[];
  webUrl?: string;
}

function jsonResult(payload: unknown, isError = false) {
  return {
    content: [{ type: 'text' as const, text: JSON.stringify(payload) }],
    ...(isError ? { isError: true } : {}),
  };
}

/** Epoch ms; 0 for missing or unparseable, so such items sort oldest. */
function ts(value: string | undefined | null): number {
  const t = Date.parse(value ?? '');
  return Number.isNaN(t) ? 0 : t;
}

/** When a chat last saw activity: the newer of its last message and its last update. */
export function chatActivity(chat: Chat): number {
  return Math.max(ts(chat.lastMessagePreview?.createdDateTime), ts(chat.lastUpdatedDateTime));
}

const ENTITIES: Record<string, string> = {
  nbsp: ' ',
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
};

/**
 * Plain text from a Teams message body, cut to `max` characters. Teams bodies are
 * almost always HTML (<p>, <at>, <img>, <attachment>), and a preview full of tags
 * wastes the small payload this tool aims for.
 */
export function plainPreview(html: string | undefined, max = PREVIEW_CHARS): string {
  if (!html) return '';
  const text = html
    .replace(/<(script|style)[^>]*>[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<br\s*\/?>|<\/(p|div|li|tr)>/gi, ' ')
    .replace(/<[^>]*>/g, '')
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&([a-z]+);/gi, (m, name) => ENTITIES[name.toLowerCase()] ?? m)
    .replace(/\s+/g, ' ')
    .trim();
  return text.length > max ? `${text.slice(0, max - 1).trimEnd()}…` : text;
}

function nameTokens(value: string): string[] {
  return value
    .toLowerCase()
    .split(/[\s,]+/)
    .filter(Boolean);
}

/**
 * Whether a message's sender is one of the people asked about. Pure; exported for tests.
 *
 * Graph's chatMessage from.user carries id and displayName only — no email or UPN
 * (the identity resource type has neither). So an email address is matched through
 * the user id it resolved to; a name is matched on displayName, case-insensitively,
 * with every typed word required ("scott" or "Schatz, Scott" both find "Scott
 * Schatz"). userPrincipalName/email are compared too, should Graph ever send them.
 */
export function matchesFrom(from: ChatIdentity | null | undefined, filters: FromFilter[]): boolean {
  const user = from?.user;
  if (!user) return false;
  const display = new Set(nameTokens(user.displayName ?? ''));
  const emails = [user.userPrincipalName, user.email]
    .filter(Boolean)
    .map((e) => (e as string).toLowerCase());
  return filters.some((f) => {
    if (f.userId && user.id && f.userId.toLowerCase() === user.id.toLowerCase()) return true;
    if (f.email && emails.includes(f.email)) return true;
    return f.nameTokens.length > 0 && f.nameTokens.every((t) => display.has(t));
  });
}

/** A real person's (or the user's) message: not a system event, not deleted. */
function isUserMessage(m: ChatMessage): boolean {
  return m.messageType === 'message' && !m.deletedDateTime && Boolean(m.from?.user?.id);
}

export interface ChatVerdict {
  status: 'awaiting-you' | 'answered' | 'none';
  myLastMessageAt?: string;
  replies: ChatMessage[];
  /** Display names of everyone but me seen in the fetched messages, newest first. */
  others: string[];
}

/**
 * The core rule, pure so the tests can pin it down. Messages in any order.
 *
 * A message from someone else counts as a reply when (1) it comes after some
 * message of mine, (2) it was posted after `since`, and (3) with a from-filter, its
 * sender matches. If any counted reply comes after my LAST message, the chat is
 * awaiting-you and those unanswered replies are returned; otherwise, if any counted
 * reply exists, I already posted after it and the chat is answered.
 *
 * A chat where I have not posted in the fetched window is 'none': nothing there is
 * a reply to me, and an unprompted message is not what this tool is asked about.
 */
export function judgeChat(
  messages: ChatMessage[],
  opts: { meId: string; since: number; from?: FromFilter[] }
): ChatVerdict {
  const me = opts.meId.toLowerCase();
  const isMe = (m: ChatMessage) => m.from?.user?.id?.toLowerCase() === me;
  const sorted = messages
    .filter(isUserMessage)
    .sort((a, b) => ts(a.createdDateTime) - ts(b.createdDateTime));

  const others: string[] = [];
  for (const m of [...sorted].reverse()) {
    const name = m.from?.user?.displayName;
    if (!isMe(m) && name && !others.includes(name)) others.push(name);
  }

  const mine = sorted.filter(isMe);
  if (!mine.length) return { status: 'none', replies: [], others };
  const firstMine = ts(mine[0].createdDateTime);
  const lastMine = ts(mine[mine.length - 1].createdDateTime);
  const myLastMessageAt = mine[mine.length - 1].createdDateTime;

  const counted = sorted.filter((m) => {
    if (isMe(m)) return false;
    const t = ts(m.createdDateTime);
    if (t <= firstMine || t <= opts.since) return false;
    return !opts.from?.length || matchesFrom(m.from, opts.from);
  });
  const unanswered = counted.filter((m) => ts(m.createdDateTime) > lastMine);
  if (unanswered.length)
    return { status: 'awaiting-you', myLastMessageAt, replies: unanswered, others };
  if (counted.length) return { status: 'answered', myLastMessageAt, replies: counted, others };
  return { status: 'none', myLastMessageAt, replies: [], others };
}

function normalizeSince(value: string | undefined): number | { error: string } {
  if (value === undefined || value.trim() === '') {
    return Date.now() - DEFAULT_SINCE_DAYS * 24 * 60 * 60 * 1000;
  }
  const t = Date.parse(value);
  if (Number.isNaN(t)) return { error: `since is not a date: '${value}'. Use ISO 8601.` };
  return t;
}

/** Path + query of an @odata.nextLink, as makeRequest wants it (graph-tools does the same). */
function nextPath(nextLink: string): string {
  const url = new URL(nextLink);
  return url.pathname.replace(/^\/v1\.0/, '') + url.search;
}

export function registerChatReplyTools(
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

  /** The caller's Entra object id: from the token in HTTP mode, from /me otherwise. */
  async function resolveMeId(): Promise<string | undefined> {
    const oid = getRequestActor()?.oid;
    if (oid) return oid;
    try {
      const me = (await graphClient.makeRequest('/me?$select=id')) as { id?: string };
      return me?.id;
    } catch (error) {
      logger.warn(`list-chat-replies: could not resolve /me (${(error as Error).message})`);
      return undefined;
    }
  }

  /**
   * Turns the typed names/addresses into matchers. An address is looked up in the
   * directory so it can be matched by user id — chat messages carry no email.
   */
  async function resolveFrom(inputs: string[]): Promise<FromFilter[]> {
    return Promise.all(
      inputs.map(async (raw): Promise<FromFilter> => {
        const input = raw.trim();
        if (!input.includes('@')) return { input, nameTokens: nameTokens(input) };
        const email = input.toLowerCase();
        try {
          const user = (await graphClient.makeRequest(
            `/users/${encodeURIComponent(input)}?$select=id,displayName`
          )) as { id?: string; displayName?: string };
          return {
            input,
            email,
            userId: user?.id,
            nameTokens: user?.displayName ? nameTokens(user.displayName) : [],
          };
        } catch (error) {
          logger.info(`list-chat-replies: '${input}' not resolved (${(error as Error).message})`);
          return { input, email, nameTokens: [] };
        }
      })
    );
  }

  /**
   * Chats active since `since`, newest activity first. No $orderby: it was rejected
   * on /me/chats when tried for tsq.28, so pages are read and sorted here. A page
   * with nothing newer than `since` ends the walk early, on the assumption that
   * Teams pages chats roughly by recent activity (to confirm live; if it does not,
   * MAX_CHAT_PAGES still bounds the walk and `capped` says so).
   */
  async function listActiveChats(since: number): Promise<{ chats: Chat[]; moreExist: boolean }> {
    const chats: Chat[] = [];
    let path: string | undefined = `/me/chats?$expand=lastMessagePreview&$top=${PAGE_SIZE}`;
    let pages = 0;
    while (path && pages < MAX_CHAT_PAGES) {
      const res = (await graphClient.makeRequest(path)) as {
        value?: Chat[];
        '@odata.nextLink'?: string;
      };
      pages++;
      const page = (res?.value ?? []).filter((c) => c.id);
      const active = page.filter((c) => chatActivity(c) > since);
      chats.push(...active);
      const next = res?.['@odata.nextLink'];
      path = next && active.length > 0 ? nextPath(next) : undefined;
      if (next && active.length > 0 && pages >= MAX_CHAT_PAGES) {
        return { chats: sortChats(chats), moreExist: true };
      }
    }
    return { chats: sortChats(chats), moreExist: false };
  }

  function sortChats(chats: Chat[]): Chat[] {
    return chats.sort((a, b) => chatActivity(b) - chatActivity(a));
  }

  /**
   * The newest 50 messages of each chat, 20 chats per $batch envelope. Fifty is
   * Graph's page ceiling. No $filter/$orderby: Graph documents a lastModifiedDateTime
   * filter only alongside a matching $orderby, and my last message may predate
   * `since` yet still be what a later reply answers — so the page is taken whole
   * and judgeChat does the sorting and the since-cut.
   */
  async function fetchMessages(
    chats: Chat[]
  ): Promise<Map<string, { messages: ChatMessage[] } | { error: string; status?: number }>> {
    const out = new Map<string, { messages: ChatMessage[] } | { error: string; status?: number }>();
    const requests = chats.map((c, i) => ({
      id: String(i),
      chatId: c.id as string,
      method: 'GET',
      url: `/chats/${encodeURIComponent(c.id as string)}/messages?$top=${PAGE_SIZE}`,
    }));
    for (let i = 0; i < requests.length; i += BATCH_MAX) {
      const chunk = requests.slice(i, i + BATCH_MAX);
      try {
        const res = await graphClient.batchRequest(
          chunk.map(({ id, method, url }) => ({ id, method, url }))
        );
        const byId = new Map((res?.responses ?? []).map((r) => [r.id, r]));
        for (const req of chunk) {
          const r = byId.get(req.id);
          if (r && r.status >= 200 && r.status < 300) {
            out.set(req.chatId, { messages: (r.body as { value?: ChatMessage[] })?.value ?? [] });
          } else {
            const code = (r?.body as { error?: { code?: string } } | undefined)?.error?.code;
            out.set(req.chatId, {
              status: r?.status,
              error: `Chat could not be read (${r?.status ?? 'no response'}${code ? ` ${code}` : ''}).`,
            });
          }
        }
      } catch (error) {
        logger.warn(`list-chat-replies batch failed: ${(error as Error).message}`);
        for (const req of chunk) out.set(req.chatId, { error: 'Chat could not be read.' });
      }
    }
    return out;
  }

  register('list-chat-replies', false, () => {
    server.tool(
      'list-chat-replies',
      // Front-loaded with the words people type: did anyone reply in Teams, unanswered chats.
      'Did anyone reply in Teams? Finds unanswered chats: Teams chats where someone answered ' +
        'you (or an agent posting as you) and you have not responded since — who answered me, ' +
        'when, and a short preview, so you can follow up in Teams chat. Works from live data ' +
        'every time; nothing is stored or tracked in advance, so it covers any 1:1, group or ' +
        'meeting chat you or an agent started or posted in. Teams channels are not included.\n\n' +
        'A reply is a message from someone else after one of yours. since (default: 7 days ' +
        'ago) limits it to replies posted after that time. from limits it to replies from ' +
        'named people: give email addresses (looked up in the directory and matched by user ' +
        "id) or display names — Teams messages carry the sender's name but not their email, " +
        'so a name is matched against the display name, every word required. By default ' +
        "returns only chats awaiting you ('awaiting-you'); onlyUnanswered=false also returns " +
        "chats you already answered ('answered'). chatId checks one chat. Reads the newest 50 " +
        `messages of each of the ${DEFAULT_TOP} (top, max ${MAX_TOP}) most recently active ` +
        'chats; open a reply with get-chat-message.',
      {
        since: z
          .string()
          .optional()
          .describe('Only replies posted after this time (ISO 8601). Default: 7 days ago.'),
        from: z
          .array(z.string().min(1))
          .max(10)
          .optional()
          .describe(
            'Only count replies from these people: email addresses or display names. ' +
              'Omit for anyone but you.'
          ),
        chatId: z.string().optional().describe('Check only this chat (id from list-chats).'),
        onlyUnanswered: z
          .boolean()
          .optional()
          .describe('Default true: only chats where someone replied and you have not since.'),
        top: z
          .number()
          .int()
          .min(1)
          .max(MAX_TOP)
          .optional()
          .describe(`Most recently active chats to examine (default ${DEFAULT_TOP}).`),
      },
      {
        title: 'list-chat-replies',
        readOnlyHint: true,
        openWorldHint: true,
      },
      async (params) =>
        withUsageLog('list-chat-replies', async () => {
          const since = normalizeSince(params.since);
          if (typeof since === 'object') return jsonResult(since, true);
          const onlyUnanswered = params.onlyUnanswered ?? true;
          const top = Math.min(params.top ?? DEFAULT_TOP, MAX_TOP);

          const meId = await resolveMeId();
          if (!meId) {
            return jsonResult(
              { error: 'Could not tell which messages are yours: your user id is unavailable.' },
              true
            );
          }
          const fromFilters = params.from?.length ? await resolveFrom(params.from) : undefined;

          let candidates: Chat[];
          let activeCount: number;
          let moreExist = false;
          if (params.chatId?.trim()) {
            const chatId = params.chatId.trim();
            let chat: Chat = { id: chatId };
            try {
              chat = {
                ...((await graphClient.makeRequest(
                  `/chats/${encodeURIComponent(chatId)}?$select=id,topic,chatType,webUrl`
                )) as Chat),
                id: chatId,
              };
            } catch (error) {
              // The messages call below reports the real failure; topic is cosmetic.
              logger.info(
                `list-chat-replies: chat details unavailable (${(error as Error).message})`
              );
            }
            candidates = [chat];
            activeCount = 1;
          } else {
            try {
              const listed = await listActiveChats(since);
              moreExist = listed.moreExist;
              activeCount = listed.chats.length;
              candidates = listed.chats.slice(0, top);
            } catch (error) {
              return jsonResult(
                { error: `Could not list your chats: ${(error as Error).message}` },
                true
              );
            }
          }

          const fetched = await fetchMessages(candidates);
          const chats: ChatRow[] = [];
          let awaitingYou = 0;
          let answered = 0;
          let unknown = 0;
          for (const chat of candidates) {
            const label = {
              chatId: chat.id,
              ...(chat.topic ? { topic: chat.topic } : {}),
              ...(chat.chatType ? { chatType: chat.chatType } : {}),
            };
            const got = fetched.get(chat.id as string);
            if (!got || 'error' in got) {
              unknown++;
              chats.push({ ...label, status: 'unknown', error: got?.error ?? 'Not checked.' });
              continue;
            }
            const verdict = judgeChat(got.messages, { meId, since, from: fromFilters });
            if (verdict.status === 'none') continue;
            if (verdict.status === 'awaiting-you') awaitingYou++;
            else answered++;
            if (verdict.status === 'answered' && onlyUnanswered) continue;

            const shown = verdict.replies.slice(-REPLIES_SHOWN);
            chats.push({
              ...label,
              // Who is in the chat, from the messages already fetched — no members call.
              ...(chat.topic ? {} : { with: verdict.others.slice(0, 5) }),
              status: verdict.status,
              myLastMessageAt: verdict.myLastMessageAt,
              replyCount: verdict.replies.length,
              replies: shown.map((m) => ({
                from: m.from?.user?.displayName ?? '(unknown)',
                at: m.createdDateTime,
                messageId: m.id,
                preview:
                  plainPreview(m.body?.content) || (m.attachments?.length ? '[attachment]' : ''),
              })),
              ...(chat.webUrl ? { webUrl: chat.webUrl } : {}),
            });
          }
          // Awaiting first, then most recent reply first.
          const order = { 'awaiting-you': 0, answered: 1, unknown: 2 } as Record<string, number>;
          const lastAt = (c: ChatRow) =>
            c.replies?.length ? ts(c.replies[c.replies.length - 1].at) : 0;
          chats.sort((a, b) => order[a.status] - order[b.status] || lastAt(b) - lastAt(a));

          const capped = moreExist || activeCount > candidates.length;
          const unresolved = fromFilters?.filter((f) => f.email && !f.userId).map((f) => f.input);
          return jsonResult({
            chats,
            summary: {
              since: new Date(since).toISOString(),
              chatsActive: activeCount,
              examined: candidates.length,
              awaitingYou,
              answered,
              ...(unknown ? { unknown } : {}),
              ...(fromFilters ? { from: fromFilters.map((f) => f.input) } : {}),
            },
            ...(unresolved?.length
              ? {
                  warning:
                    `Not found in the directory: ${unresolved.join(', ')}. Their replies ` +
                    'cannot be matched by address; try their display name.',
                }
              : {}),
            ...(capped
              ? {
                  capped: true,
                  note:
                    `Only the ${candidates.length} most recently active chats were checked` +
                    `${moreExist ? ' (more chats exist beyond those listed)' : ` of ${activeCount}`}. ` +
                    `Raise top (max ${MAX_TOP}), narrow since, or pass chatId.`,
                }
              : {}),
          });
        })
    );
  });
}
