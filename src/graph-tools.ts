import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import logger from './logger.js';
import {
  dailyLimitRefusalResult,
  describeFailure,
  isDailyLimitRefusal,
  logToolUsage,
  reserveDailyAllowance,
  withUsageLog,
} from './usage-log.js';
import GraphClient from './graph-client.js';
import AuthManager from './auth.js';
import { api } from './generated/client.js';
import { z } from 'zod';
import { readFileSync } from 'fs';
import { createHash } from 'crypto';
import path from 'path';
import { fileURLToPath } from 'url';
import { TOOL_CATEGORIES } from './tool-categories.js';
import { getRequestTokens } from './request-context.js';
import { parseTeamsUrl } from './lib/teams-url-parser.js';
import { registerNotificationTools } from './notification-tools.js';
import { registerFileTextTools } from './file-text-tools.js';
import { registerFileTools } from './file-tools.js';
import { registerReplyTrackingTools, trackingProperties } from './reply-tracking-tools.js';
import { registerChatReplyTools } from './chat-reply-tools.js';
import { registerMeetingTools } from './meeting-tools.js';
import { registerInferenceTools } from './inference-tools.js';
import { buildBM25Index, scoreQuery, tokenize, type BM25Index } from './lib/bm25.js';
export interface DiscoverySearchIndex {
  bm25: BM25Index;
  nameTokens: Map<string, Set<string>>;
}
import { describeToolSchema } from './lib/tool-schema.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

interface EndpointConfig {
  pathPattern: string;
  method: string;
  toolName: string;
  scopes?: string[];
  workScopes?: string[];
  returnDownloadUrl?: boolean;
  supportsTimezone?: boolean;
  supportsExpandExtendedProperties?: boolean;
  llmTip?: string;
  /**
   * Replaces the upstream Graph description entirely.
   *
   * The generated descriptions are scraped Graph reference prose and a number of
   * them describe a DIFFERENT operation than the tool performs — `delete-mail-message`
   * opened with "Delete eventMessage", `list-mail-messages` with "Get an open
   * extension (openTypeExtension object)". Under eager tool loading that was
   * survivable, because the model read the whole string and found the tip. Claude
   * now searches for tools instead of loading them all, and discovery matches on
   * the opening text — so a first sentence about open extensions means the mail
   * tool never surfaces for a mail query, and argues against itself when it does.
   *
   * Write these as a person would search for them: verb first, then the object, in
   * the words a user actually types. Disambiguate near neighbours explicitly.
   */
  description?: string;
  /**
   * Replaces the request-body schema with a permissive object carrying this text.
   *
   * The generated client inlines the full Graph entity schema for every write —
   * `create-chat` expands the entire chat entity, members, every message subtype,
   * and costs ~7,000 tokens on its own. Measured 2026-09-02: the 24 heaviest tools
   * carried 62% of the whole 142,000-token catalogue, and the bulk of that is body
   * schemas nobody needs in full. A model does not need the complete chat entity to
   * start a chat; it needs to know that chatType and members are required.
   *
   * So the body becomes a passthrough object — any shape still reaches Graph, and
   * Graph remains the validator it always was — described by a short, accurate
   * sentence naming the required and common fields. That is cheaper AND clearer
   * than a 200-property schema the model has to infer intent from.
   *
   * Write these as a worked example, not a type definition.
   */
  compactBody?: string;
  skipEncoding?: string[]; // Parameter names that should NOT be URL-encoded (for function-style API calls)
  contentType?: string;
  acceptType?: string; // Custom Accept header for endpoints returning non-JSON content (e.g., text/vtt)
  readOnly?: boolean; // When true, allow this endpoint in read-only mode even if method is not GET
}

const endpointsData = JSON.parse(
  readFileSync(path.join(__dirname, 'endpoints.json'), 'utf8')
) as EndpointConfig[];

/** When set to a positive integer, caps Graph `$top` on list requests (see README). */
function maxTopFromEnv(): number | undefined {
  const raw = process.env.MS365_MCP_MAX_TOP;
  if (raw === undefined || raw === '') return undefined;
  const n = Number.parseInt(raw, 10);
  if (!Number.isFinite(n) || n < 1) {
    logger.warn(
      `Ignoring invalid MS365_MCP_MAX_TOP=${JSON.stringify(raw)} (use a positive integer)`
    );
    return undefined;
  }
  return n;
}

function clampTopQueryParam(queryParams: Record<string, string>): void {
  const cap = maxTopFromEnv();
  if (cap === undefined || queryParams['$top'] === undefined) return;
  const requested = Number.parseInt(queryParams['$top'], 10);
  if (!Number.isFinite(requested) || requested <= cap) return;
  logger.info(`Clamping $top from ${requested} to ${cap} (MS365_MCP_MAX_TOP)`);
  queryParams['$top'] = String(cap);
}

// Backported from upstream #597 (dca6460), which we had not taken at our v0.88.2 pin.
// list-mail-messages carried 12 of the 29 errors in the week to 2026-09-18, and both
// failing shapes below reproduce as a Graph 400.
//
// Outlook message collections only. The path has to be a mailbox owner, optionally some
// mailFolders/childFolders nesting, and then end at messages. Matching the owner prefix and
// the collection name separately would catch /me/chats/{id}/messages, and matching any
// mailFolders descendant would catch list-mail-folders, list-mail-child-folders,
// list-mail-rules and list-mail-attachments, none of which take message KQL. /chats, /teams
// and /planner messages, and directory search, have their own conventions and are untouched.
const OUTLOOK_MAIL_PATH =
  /^\/(?:me|users\/[^/]+)(?:\/(?:mailFolders|childFolders)\/[^/]+)*\/messages(?:\/delta\(\))?$/i;

function isOutlookMailPath(path: string): boolean {
  return OUTLOOK_MAIL_PATH.test(path);
}

/** A quoted run starting at `start` (an opening quote), with escapes preserved. */
function readQuotedSegment(
  expr: string,
  start: number
): { segment: string; end: number } | undefined {
  let j = start + 1;
  let segment = '';
  while (j < expr.length) {
    // Consume an escaped backslash as a unit, otherwise the second slash pairs with a real
    // delimiter behind it and the scan runs off the end of a well-formed string.
    if (expr[j] === '\\' && expr[j + 1] === '\\') {
      segment += '\\\\';
      j += 2;
      continue;
    }
    if (expr[j] === '\\' && expr[j + 1] === '"') {
      segment += '\\"';
      j += 2;
      continue;
    }
    if (expr[j] === '"') return { segment, end: j };
    segment += expr[j];
    j += 1;
  }
  return undefined;
}

// The properties KQL recognises on a message, from the searchable-email-property table at
// learn.microsoft.com/en-us/graph/search-query-parameter. `category` is documented only on
// the Exchange page that table links to, and both spellings of hasAttachment(s) are here
// because that table and its own example disagree. Shape alone is not enough to tell a
// clause from a phrase: "RE: quarterly report" and "Q3: plan.pdf" both look like
// property:value.
const MAIL_SEARCH_PROPERTIES = new Set([
  'attachment',
  'bcc',
  'body',
  'category',
  'cc',
  'from',
  'hasattachment',
  'hasattachments',
  'importance',
  'kind',
  'participants',
  'received',
  'recipients',
  'sent',
  'size',
  'subject',
  'to',
]);

/**
 * `property:` or a comparison — `received>=2024-01-01`, `size>1000`. The value must follow
 * the operator immediately: KQL demotes a restriction with whitespace around the operator
 * to free text, so `from: the desk of the CEO` is a phrase that has to keep its quotes,
 * not a clause to unwrap.
 */
const CLAUSE_HEAD = /^([A-Za-z]+)(?::|<=|>=|<>|=|<|>)\S/;

/** KQL's boolean operators: uppercase and free-standing, per the KQL syntax reference. */
const BOOLEAN_JOIN = /\s(?:AND|OR|NOT)\s/;

/** The recognised `property:`/comparison head of a run, or null if it does not open with one. */
function clauseHead(segment: string): RegExpExecArray | null {
  const head = CLAUSE_HEAD.exec(segment);
  return head && MAIL_SEARCH_PROPERTIES.has(head[1].toLowerCase()) ? head : null;
}

/** Append a slash when the trailing run is odd, so it cannot escape a quote placed after it. */
function balanceTrailingSlashes(text: string): string {
  const slashes = text.length - text.replace(/\\+$/, '').length;
  return slashes % 2 === 1 ? `${text}\\` : text;
}

/**
 * How a quoted run should be emitted once the whole expression gains its enclosing pair.
 *
 * - `phrase` keeps the quotes where they are, escaped as \". A run holding the value of a
 *   restriction is always this, even when its text contains a colon
 *   (subject:"RE: quarterly report"), as is anything that does not open with a recognised
 *   property at all ("quarterly report", "RE: quarterly report").
 * - `clause` drops the quotes: either one bare restriction the caller quoted by mistake
 *   ("from:john" AND subject:meeting), or several joined by boolean operators and quoted as
 *   a group ("from:john AND subject:meeting" OR from:jane). Both are per-clause quoting,
 *   which is the directory convention and a 400 here.
 * - `restriction-value` moves the quotes past the operator: "subject:quarterly report"
 *   becomes subject:\"quarterly report\". Escaping in place would leave `subject:` inside
 *   the phrase as literal text and lose the restriction entirely, and dropping the quotes
 *   would bind only `quarterly` to subject and let `report` float as free text. Only moving
 *   them keeps both the property and the grouping.
 */
type RunKind = 'phrase' | 'clause' | 'restriction-value';

function classifyRun(segment: string, introducedByProperty: boolean): RunKind {
  if (introducedByProperty) return 'phrase';
  const head = clauseHead(segment);
  if (!head) return 'phrase';
  if (BOOLEAN_JOIN.test(segment) || !/\s/.test(segment)) return 'clause';
  return 'restriction-value';
}

/**
 * Rewrite the interior of a mail KQL expression so it can be wrapped in one pair of
 * double quotes. Phrase quotes are escaped as \" by analogy with the rule Microsoft
 * documents for directory search; mail's own docs never show an embedded quote, so that
 * form is inferred rather than published.
 */
function rewriteMailSearchQuotes(expr: string): string {
  let out = '';
  let i = 0;
  while (i < expr.length) {
    if (expr[i] === '\\' && expr[i + 1] === '\\') {
      out += '\\\\';
      i += 2;
      continue;
    }
    if (expr[i] === '\\' && expr[i + 1] === '"') {
      out += '\\"';
      i += 2;
      continue;
    }
    if (expr[i] !== '"') {
      out += expr[i];
      i += 1;
      continue;
    }
    // An unterminated run is read as a missing closing quote rather than a stray opening
    // one; dropping the delimiter would shed the grouping and widen the search. Its tail is
    // balanced first, because the closer synthesized below would otherwise pair with a
    // trailing backslash and let the next quote end the string early.
    const run = readQuotedSegment(expr, i);
    const segment = run ? run.segment : balanceTrailingSlashes(expr.slice(i + 1));
    const introducedByProperty = i > 0 && expr[i - 1] === ':';
    switch (classifyRun(segment, introducedByProperty)) {
      case 'clause':
        out += segment;
        break;
      case 'restriction-value': {
        const head = clauseHead(segment)!;
        const valueAt = head[0].length - 1;
        out += `${segment.slice(0, valueAt)}\\"${segment.slice(valueAt)}\\"`;
        break;
      }
      default:
        out += `\\"${segment}\\"`;
    }
    if (!run) break;
    i = run.end + 1;
  }
  return balanceTrailingSlashes(out.trim());
}

/**
 * Outlook mail wants the whole KQL expression inside one pair of double quotes
 * ($search="from:x AND subject:y"). Models quote each clause instead
 * ($search='"from:x" AND subject:y'), or send a phrase with no enclosing pair
 * ($search='subject:"quarterly report"'). Both are 400s. Normalize to one enclosing
 * pair, mirroring the Body auto-wrap already done in executeGraphTool.
 *
 * Verified against Graph: a bare single term and a correctly wrapped expression both
 * succeed; 'subject:"quarterly report"' and '"quarterly report" AND from:x' are both
 * rejected until the enclosing pair is added.
 */
function normalizeSearchQueryParam(
  queryParams: Record<string, string>,
  path: string,
  toolAlias: string
): CallToolResult | undefined {
  if (!isOutlookMailPath(path)) return;

  const raw = queryParams['$search'];
  if (raw === undefined) return;
  const trimmed = raw.trim();

  // Graph rejects $search with $orderby on messages (SearchWithOrderBy, 400). It
  // was 19 of 23 list-mail-messages failures in the week of 27 September: the
  // model asks for "newest first" out of habit. Search results already come back
  // newest-sent first (learn.microsoft.com/graph/search-query-parameter), which
  // is what the sort nearly always asked for, so drop it rather than fail.
  if (queryParams['$orderby'] !== undefined) {
    logger.info(`Dropped '$orderby' from ${toolAlias}: Graph does not allow it with $search`);
    delete queryParams['$orderby'];
  }

  // Nothing searchable. Deleting $search would widen the request into an unfiltered listing
  // of the whole mailbox and hand it back as though it were the search result, which is a
  // worse answer than the 400 Graph would have returned, so refuse instead.
  const noSearchableText = (): CallToolResult => {
    logger.warn(`Refusing ${toolAlias}: '$search' has no searchable text`);
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            error: 'invalid_search',
            tool: toolAlias,
            message:
              'The $search parameter has no searchable text. Supply a KQL expression such as "from:john" or "subject:budget", or omit $search to list messages unfiltered.',
          }),
        },
      ],
      isError: true,
    };
  };

  if (trimmed === '' || /^["'\s]+$/.test(trimmed)) return noSearchableText();

  // An expression already inside one enclosing pair is unwrapped first, so its interior
  // is judged on its own terms and re-wrapped unchanged. Without this, a correctly
  // wrapped free-text search ("quarterly report") would be read as a phrase and become
  // a phrase search ("\"quarterly report\"").
  let expr = trimmed;
  if (expr.startsWith('"')) {
    const whole = readQuotedSegment(expr, 0);
    // No closing quote at all means the caller dropped it off the enclosing pair, not that
    // they opened a phrase. Reading it as a phrase would send a literal search for the whole
    // expression, which matches nothing and gives the model no error to correct against.
    if (!whole) expr = expr.slice(1);
    else if (whole.end === expr.length - 1) expr = whole.segment;
  }

  const inner = rewriteMailSearchQuotes(expr);
  // Unreachable while the guard above catches every all-quote/whitespace value; kept so a
  // later change to the rewriter cannot quietly send Graph $search="".
  if (inner === '') return noSearchableText();
  const normalized = `"${inner}"`;
  if (normalized !== raw) {
    logger.info(`Auto-corrected parameter '$search': normalized KQL quoting to ${normalized}`);
    queryParams['$search'] = normalized;
  }
}

type TextContent = {
  type: 'text';
  text: string;
  [key: string]: unknown;
};

type ImageContent = {
  type: 'image';
  data: string;
  mimeType: string;
  [key: string]: unknown;
};

type AudioContent = {
  type: 'audio';
  data: string;
  mimeType: string;
  [key: string]: unknown;
};

type ResourceTextContent = {
  type: 'resource';
  resource: {
    text: string;
    uri: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

type ResourceBlobContent = {
  type: 'resource';
  resource: {
    blob: string;
    uri: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

type ResourceContent = ResourceTextContent | ResourceBlobContent;

type ContentItem = TextContent | ImageContent | AudioContent | ResourceContent;

interface CallToolResult {
  content: ContentItem[];
  _meta?: Record<string, unknown>;
  isError?: boolean;

  [key: string]: unknown;
}

function errorResult(payload: Record<string, unknown>): CallToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(payload) }], isError: true };
}

// ---- tsq.28: request-body normalization for mail and calendar writes ---------
//
// Week to 2026-10-09: send-mail 400 RequestBodyRead ×6 (3 users),
// create-calendar-event 400 UnableToDeserializePostBody ×5, create-draft-email ×1.
// Both codes mean Graph could not bind the JSON to its type at all — a shape
// error, not a value error. The shapes models get wrong are few and unambiguous:
// recipients written as plain addresses instead of { emailAddress: { address } },
// a send-mail body missing its `message` wrapper, and a body given as a string.
// Each has exactly one correct reading, so they are fixed here rather than
// returned as a 400 the model has to decode. Times are the exception: see
// checkEventTimes.

const MAIL_RECIPIENT_FIELDS = ['toRecipients', 'ccRecipients', 'bccRecipients', 'replyTo'];

/** Tools whose body is a message, or carries one under `message`. */
const MAIL_BODY_TOOLS = new Set([
  'send-mail',
  'send-shared-mailbox-mail',
  'create-draft-email',
  'reply-mail-message',
  'reply-all-mail-message',
  'forward-mail-message',
  'create-reply-draft',
  'create-reply-all-draft',
  'create-forward-draft',
]);

/** sendMail is the one action whose body must be { message, saveToSentItems }. */
const SEND_MAIL_TOOLS = new Set(['send-mail', 'send-shared-mailbox-mail']);

const EVENT_BODY_TOOLS = new Set([
  'create-calendar-event',
  'create-specific-calendar-event',
  'update-calendar-event',
  'update-specific-calendar-event',
]);

type JsonObject = Record<string, unknown>;

function isPlainObject(value: unknown): value is JsonObject {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

/**
 * `'a@x.com'`, `{ address, name }` or `{ emailAddress: 'a@x.com' }` → the
 * `{ emailAddress: { address, name } }` Graph binds to. Anything else —
 * including an already-correct entry — is returned untouched, so Graph stays
 * the validator for shapes this does not recognise.
 */
function toEmailAddressEntry(entry: unknown): { value: unknown; changed: boolean } {
  if (typeof entry === 'string') {
    return { value: { emailAddress: { address: entry.trim() } }, changed: true };
  }
  if (!isPlainObject(entry)) return { value: entry, changed: false };
  if (typeof entry.emailAddress === 'string') {
    const { emailAddress, ...rest } = entry;
    return { value: { ...rest, emailAddress: { address: emailAddress.trim() } }, changed: true };
  }
  if (entry.emailAddress === undefined && typeof entry.address === 'string') {
    const { address, name, ...rest } = entry;
    const emailAddress: JsonObject = { address: address.trim() };
    if (typeof name === 'string') emailAddress.name = name;
    return { value: { ...rest, emailAddress }, changed: true };
  }
  return { value: entry, changed: false };
}

/** A recipient field may also arrive as one string, or several joined by , or ; */
function normalizeAddressList(value: unknown): { value: unknown; changed: boolean } {
  let list: unknown[];
  let changed = false;
  if (typeof value === 'string') {
    list = value
      .split(/[;,]/)
      .map((s) => s.trim())
      .filter(Boolean);
    changed = true;
  } else if (Array.isArray(value)) {
    list = value;
  } else if (isPlainObject(value)) {
    list = [value];
    changed = true;
  } else {
    return { value, changed: false };
  }
  const mapped = list.map((entry) => {
    const r = toEmailAddressEntry(entry);
    changed ||= r.changed;
    return r.value;
  });
  return { value: mapped, changed };
}

/** A body given as a bare string becomes an itemBody; HTML is sniffed, not assumed. */
function normalizeItemBody(value: unknown): { value: unknown; changed: boolean } {
  if (typeof value !== 'string') return { value, changed: false };
  const contentType = /<[a-z][\s\S]*>/i.test(value) ? 'HTML' : 'Text';
  return { value: { contentType, content: value }, changed: true };
}

/** Applies the recipient and body fixes to one message-shaped object, in place. */
function normalizeMessageFields(message: JsonObject, fixes: string[], where: string): void {
  for (const field of MAIL_RECIPIENT_FIELDS) {
    if (message[field] === undefined) continue;
    const r = normalizeAddressList(message[field]);
    if (r.changed) {
      message[field] = r.value;
      fixes.push(`${where}${field} → [{ emailAddress: { address } }]`);
    }
  }
  if (message.body !== undefined) {
    const r = normalizeItemBody(message.body);
    if (r.changed) {
      message.body = r.value;
      fixes.push(`${where}body string → { contentType, content }`);
    }
  }
}

/**
 * Attendees take the same address fix plus a `type`. Only entries this rewrote
 * get the default 'required' — what Outlook does with a name typed into To —
 * so an attendee already in Graph's shape is sent exactly as the caller wrote it.
 */
function normalizeAttendees(value: unknown): { value: unknown; changed: boolean } {
  const r = normalizeAddressList(value);
  if (!r.changed || !Array.isArray(r.value)) return r;
  const originals = Array.isArray(value) ? value : undefined;
  const mapped = r.value.map((a, i) => {
    const rewritten = !originals || toEmailAddressEntry(originals[i]).changed;
    return rewritten && isPlainObject(a) && a.type === undefined ? { ...a, type: 'required' } : a;
  });
  return { value: mapped, changed: true };
}

/**
 * Calendar start/end must be { dateTime, timeZone }. A bare string
 * ('2026-10-14T15:00') fails with UnableToDeserializePostBody, and unlike the
 * recipient shapes it has no single correct reading: the zone is missing, and
 * guessing UTC or Eastern would put the meeting at the wrong hour on someone's
 * calendar with nothing to show it happened. Refuse with the shape instead.
 */
function checkEventTimes(body: JsonObject, toolAlias: string): CallToolResult | undefined {
  for (const field of ['start', 'end']) {
    if (typeof body[field] !== 'string') continue;
    logger.warn(`Refusing ${toolAlias}: '${field}' is a bare string, not { dateTime, timeZone }`);
    return errorResult({
      error: 'invalid_event_time',
      tool: toolAlias,
      field,
      message:
        `${field} must be an object { dateTime: '2026-09-08T14:00:00', timeZone: 'Eastern ` +
        `Standard Time' }, not a bare string. The time zone has to be stated — use the ` +
        "one the user gave, or their mailbox's (get-mailbox-settings → timeZone). Do not guess.",
    });
  }
  return undefined;
}

/**
 * Returns the body to send (possibly rewritten) or a refusal. Each rewrite is
 * logged, like the existing parameter auto-corrections, so the weekly report
 * can tell whether a fix is firing.
 */
/**
 * Tools that can mark the message they create as awaiting a reply (tsq.28).
 * 'message' = the body IS the message (a draft); 'wrapped' = the message sits
 * under body.message (sendMail, and the reply/forward actions, whose optional
 * `message` carries writable properties onto the reply they create).
 */
const TRACK_AT_SEND_TOOLS: Record<string, 'message' | 'wrapped'> = {
  'send-mail': 'wrapped',
  'create-draft-email': 'message',
  'reply-mail-message': 'wrapped',
  'reply-all-mail-message': 'wrapped',
  'forward-mail-message': 'wrapped',
  'create-reply-draft': 'wrapped',
  'create-reply-all-draft': 'wrapped',
  'create-forward-draft': 'wrapped',
};

/**
 * Adds the reply-tracking marker to the message a send/draft tool creates.
 * Scott, 2026-10-13: "if send-mail and the reply tools could take an option to
 * track at send time, it would not depend on the agent remembering a second
 * call." The marker rides inside the request, so the Sent Items copy is born
 * tracked and list-awaiting-replies finds it like any track-reply'd message.
 */
export function applyTrackAtSend(
  toolAlias: string,
  body: unknown,
  opts: { dueBy?: string; waitingOn?: string[] }
): unknown {
  const shape = TRACK_AT_SEND_TOOLS[toolAlias];
  if (!shape) return body;
  const props = trackingProperties(opts);
  const base: JsonObject = isPlainObject(body) ? { ...body } : {};
  const target: JsonObject =
    shape === 'message'
      ? base
      : isPlainObject(base.message)
        ? { ...(base.message as JsonObject) }
        : {};
  const existing = Array.isArray(target.singleValueExtendedProperties)
    ? (target.singleValueExtendedProperties as JsonObject[]).filter(
        (p) => !props.some((q) => q.id === p.id)
      )
    : [];
  target.singleValueExtendedProperties = [...existing, ...props];
  if (shape === 'wrapped') base.message = target;
  logger.info(`${toolAlias}: tracking this message for a reply (track at send)`);
  return base;
}

export function normalizeWriteBody(
  toolAlias: string,
  body: unknown
): { body: unknown; error?: CallToolResult } {
  if (!isPlainObject(body)) return { body };
  const isMail = MAIL_BODY_TOOLS.has(toolAlias);
  const isEvent = EVENT_BODY_TOOLS.has(toolAlias);
  if (!isMail && !isEvent) return { body };

  const fixes: string[] = [];
  let out: JsonObject = { ...body };

  if (isMail) {
    // { subject, toRecipients, body } at the top level is a message that lost its
    // wrapper. Only for sendMail, whose body is never itself a message; a draft's
    // body IS the message, and reply/forward legitimately have top-level
    // toRecipients beside `comment`.
    if (
      SEND_MAIL_TOOLS.has(toolAlias) &&
      out.message === undefined &&
      out.Message === undefined &&
      ['subject', 'toRecipients', 'body'].some((k) => out[k] !== undefined)
    ) {
      const { saveToSentItems, ...message } = out;
      out = { message, ...(saveToSentItems !== undefined ? { saveToSentItems } : {}) };
      fixes.push('wrapped top-level message fields in { message }');
    }
    // reply/forward/createReply take recipients both beside `comment` and inside
    // `message`; sendMail only inside it; a draft only at the top level.
    normalizeMessageFields(out, fixes, '');
    if (isPlainObject(out.message)) {
      out.message = { ...out.message };
      normalizeMessageFields(out.message as JsonObject, fixes, 'message.');
    }
  }

  if (isEvent) {
    const timeError = checkEventTimes(out, toolAlias);
    if (timeError) return { body, error: timeError };
    if (out.attendees !== undefined) {
      const r = normalizeAttendees(out.attendees);
      if (r.changed) {
        out.attendees = r.value;
        fixes.push("attendees → [{ emailAddress: { address }, type: 'required' }]");
      }
    }
    if (out.body !== undefined) {
      const r = normalizeItemBody(out.body);
      if (r.changed) {
        out.body = r.value;
        fixes.push('body string → { contentType, content }');
      }
    }
  }

  if (fixes.length === 0) return { body };
  logger.info(`Auto-corrected request body for ${toolAlias}: ${fixes.join('; ')}`);
  return { body: out };
}

// ---- tsq.28: query options Graph rejects on specific collections -------------

/** Graph's documented page-size maximum for /me/chats and /chats/{id}/messages. */
const TEAMS_CHAT_MAX_TOP = 50;

/**
 * /chats/{id}/messages supports $orderby on exactly these two, both descending
 * (learn.microsoft.com/graph/api/chat-list-messages, "Optional query parameters").
 */
const CHAT_MESSAGE_ORDERBY = new Set(['lastmodifieddatetime desc', 'createddatetime desc']);

/** The filters list-online-meetings can actually run with delegated permissions. */
const ONLINE_MEETING_FILTER = /\b(?:joinWebUrl|joinMeetingId|videoTeleconferenceId)\b/i;

function clampTo(queryParams: Record<string, string>, max: number, toolAlias: string): void {
  const raw = queryParams['$top'];
  if (raw === undefined) return;
  const n = Number.parseInt(raw, 10);
  if (Number.isFinite(n) && n > max) {
    logger.info(`Clamping $top from ${n} to ${max} on ${toolAlias} (Graph maximum)`);
    queryParams['$top'] = String(max);
  }
}

/**
 * The first property a $filter tests: `receivedDateTime ge …` → receivedDateTime,
 * `contains(subject,'x')` → subject, `from/emailAddress/address eq …` →
 * from/emailAddress/address, `toRecipients/any(…)` → toRecipients. Undefined when
 * the expression opens with something this cannot read, so the caller leaves it
 * alone rather than act on a guess.
 */
export function firstFilterProperty(filter: string): string | undefined {
  let f = filter.trim();
  for (;;) {
    if (f.startsWith('(')) f = f.slice(1).trimStart();
    else if (/^not\s/i.test(f)) f = f.slice(3).trimStart();
    else break;
  }
  const fn = /^(?:contains|startswith|endswith)\s*\(\s*([A-Za-z][\w/]*)/i.exec(f);
  if (fn) return fn[1];
  const prop = /^([A-Za-z][\w/]*)/.exec(f);
  return prop ? prop[1].replace(/\/(?:any|all)$/i, '') : undefined;
}

/**
 * Rewrites query options that a collection is known to reject, returning a
 * refusal when the request cannot succeed as asked, and the $orderby to apply
 * client-side when one was removed but can still be honoured on the page.
 */
export function normalizeCollectionQuery(
  toolAlias: string,
  toolPath: string,
  queryParams: Record<string, string>
): { error?: CallToolResult; clientOrderby?: string } {
  // list-chats 400 ×3 in the week to 2026-10-09. /me/chats rejects $orderby
  // outright ("QueryOptions to order by lastUpdatedDateTime is not supported";
  // the tool's tip already says not to send it) and caps $top at 50.
  if (toolAlias === 'list-chats') {
    if (queryParams['$orderby'] !== undefined) {
      logger.info(`Dropped '$orderby' from ${toolAlias}: Graph does not support it on /me/chats`);
      delete queryParams['$orderby'];
    }
    clampTo(queryParams, TEAMS_CHAT_MAX_TOP, toolAlias);
  }

  // list-chat-messages 400 ×2. Same $top ceiling; $orderby only in the two
  // documented forms. Anything else is dropped — the default order is already
  // newest first, which is what an $orderby here is nearly always asking for.
  if (toolAlias === 'list-chat-messages') {
    clampTo(queryParams, TEAMS_CHAT_MAX_TOP, toolAlias);
    const orderby = queryParams['$orderby'];
    if (
      orderby !== undefined &&
      !CHAT_MESSAGE_ORDERBY.has(orderby.trim().replace(/\s+/g, ' ').toLowerCase())
    ) {
      logger.info(
        `Dropped '$orderby' from ${toolAlias}: Graph supports only lastModifiedDateTime desc or createdDateTime desc`
      );
      delete queryParams['$orderby'];
    }
  }

  // list-mail-rules 400 ErrorInvalidUrlQueryFilter ×1. The messageRules
  // collection takes no $filter, $search or $orderby; the whole set is small
  // enough to fetch and read.
  if (toolAlias === 'list-mail-rules') {
    for (const key of ['$filter', '$search', '$orderby']) {
      if (queryParams[key] !== undefined) {
        logger.info(
          `Dropped '${key}' from ${toolAlias}: Graph does not support it on messageRules`
        );
        delete queryParams[key];
      }
    }
  }

  // list-online-meetings 403 ×2. GET /me/onlineMeetings is not a listing: Graph
  // only answers it with a $filter on the join URL (or the meeting/VTC id), and
  // without one returns a 403 that reads like a permissions problem. Say what
  // it actually needs.
  if (toolAlias === 'list-online-meetings') {
    const filter = queryParams['$filter'];
    if (!filter || !ONLINE_MEETING_FILTER.test(filter)) {
      logger.warn(`Refusing ${toolAlias}: no joinWebUrl filter`);
      return {
        error: errorResult({
          error: 'join_url_required',
          tool: toolAlias,
          message:
            'Graph cannot list online meetings; it can only look one up by its join link. Pass filter: "JoinWebUrl eq \'<url>\'". ' +
            'Get the URL from the calendar event (get-calendar-event → onlineMeeting.joinUrl), ' +
            'or convert a recap/short link with parse-teams-url first.',
        }),
      };
    }
  }

  // list-mail-messages / list-mail-folder-messages 400 InefficientFilter ×2.
  // Graph requires every $orderby property to appear in $filter, first and in
  // the same order (learn.microsoft.com/graph/query-parameters, "Using $filter
  // and $orderby in the same query"). When the leading property differs, the
  // request cannot succeed; drop $orderby and sort the returned page instead,
  // the same trade list-conversation-messages already makes. $search with
  // $orderby is handled (dropped) in normalizeSearchQueryParam.
  if (
    isOutlookMailPath(toolPath) &&
    queryParams['$filter'] !== undefined &&
    queryParams['$orderby'] !== undefined &&
    queryParams['$search'] === undefined
  ) {
    const orderby = queryParams['$orderby'];
    const orderProp = orderby.split(',')[0].trim().split(/\s+/)[0];
    const filterProp = firstFilterProperty(queryParams['$filter']);
    if (filterProp && orderProp && filterProp.toLowerCase() !== orderProp.toLowerCase()) {
      logger.info(
        `Dropped '$orderby' from ${toolAlias}: '${orderProp}' is not the first property in $filter (InefficientFilter); sorting the page client-side`
      );
      delete queryParams['$orderby'];
      // Only a single plain property can be sorted faithfully here.
      if (/^[A-Za-z]\w*(\s+(asc|desc))?$/i.test(orderby.trim())) {
        return { clientOrderby: orderby.trim() };
      }
    }
  }

  return {};
}

/**
 * Sorts the `value` page of a successful list result by a single-property
 * orderby and marks it, or returns the result untouched if it cannot. Shared by
 * list-conversation-messages and the InefficientFilter fallback above.
 */
export function sortResultPage(
  result: CallToolResult,
  orderby: string,
  note: string,
  toolAlias: string
): CallToolResult {
  try {
    const textBlock = result.content?.[0];
    if (!textBlock || textBlock.type !== 'text' || typeof textBlock.text !== 'string') {
      return result;
    }
    const parsed = JSON.parse(textBlock.text);
    if (!parsed || !Array.isArray(parsed.value)) {
      return result;
    }

    const orderbyParts = orderby.trim().split(/\s+/);
    const field = orderbyParts[0];
    const direction = (orderbyParts[1] || 'asc').toLowerCase();
    const ascending = direction !== 'desc';

    parsed.value.sort((a: Record<string, unknown>, b: Record<string, unknown>) => {
      const aVal = a?.[field];
      const bVal = b?.[field];
      if (aVal === undefined || aVal === null) return ascending ? 1 : -1;
      if (bVal === undefined || bVal === null) return ascending ? -1 : 1;
      if (aVal < bVal) return ascending ? -1 : 1;
      if (aVal > bVal) return ascending ? 1 : -1;
      return 0;
    });

    parsed._tsq_clientSorted = {
      field,
      direction: ascending ? 'asc' : 'desc',
      note,
    };

    return {
      ...result,
      content: [{ ...textBlock, text: JSON.stringify(parsed) }],
    };
  } catch (err) {
    logger.warn(`Failed to client-side sort ${toolAlias} result: ${(err as Error).message}`);
    return result;
  }
}

// ---- tsq.28: idempotent message delete ----------------------------------------

/** A single Outlook message, in the mailbox or any folder — not its attachments. */
const OUTLOOK_MESSAGE_ITEM_PATH =
  /^\/(?:me|users\/[^/]+)(?:\/(?:mailFolders|childFolders)\/[^/]+)*\/messages\/[^/]+$/i;

/**
 * delete-mail-message 404 ErrorItemNotFound ×4 after tsq.27 (+29 before). A
 * delete MOVES the message to Deleted Items and gives it a new id, so the old id
 * 404s from then on; nearly every one of these was a second delete of a message
 * the first call had already removed — sometimes our own 5xx retry, where the
 * first attempt landed but its response did not. The user's intent ("this
 * message gone from where it was") is met, so this reports success, flagged so
 * the model can say so if the id was simply wrong.
 */
function alreadyDeletedResult(
  tool: { alias: string; method: string; path: string },
  response: { isError?: boolean; content: { text?: string }[] }
): CallToolResult | undefined {
  if (!response.isError || tool.method.toUpperCase() !== 'DELETE') return undefined;
  if (!OUTLOOK_MESSAGE_ITEM_PATH.test(tool.path)) return undefined;
  const failure = describeFailure(response);
  if (failure.status !== 404 || failure.code !== 'ErrorItemNotFound') return undefined;
  logger.info(`${tool.alias}: 404 ErrorItemNotFound treated as already deleted`);
  return {
    content: [
      {
        type: 'text',
        text: JSON.stringify({
          success: true,
          alreadyDeleted: true,
          note:
            'No message with this id exists any more — it was already deleted or moved (a ' +
            'delete moves it to Deleted Items under a new id). Nothing further to do. Do not ' +
            'retry; if the user expected it to still be there, check Deleted Items.',
        }),
      },
    ],
  };
}

// ---- tsq.28: get-sharepoint-site-by-path ---------------------------------------

/**
 * Path segments that belong to a page or library inside a site, not to the site
 * itself. A pasted browser URL ends in one of these more often than not.
 */
const SITE_PATH_STOP_SEGMENTS = new Set([
  'sitepages',
  'pages',
  '_layouts',
  '_api',
  '_vti_bin',
  'lists',
  'forms',
  'siteassets',
  'shared documents',
  'documents',
]);

function decodeSegment(segment: string): string {
  try {
    return decodeURIComponent(segment);
  } catch {
    return segment;
  }
}

/**
 * Builds the Graph path for a site addressed by hostname and server-relative
 * path: GET /sites/{hostname}:/{server-relative-path}
 * (learn.microsoft.com/graph/api/site-getbypath). Each path segment is encoded
 * on its own; the ':' after the hostname is the address syntax and must stay
 * literal, which is why this is built here and not through the generic
 * path-parameter encoder (it would turn the slashes into %2F).
 *
 * Replaces the generated `/sites/{site-id}/getByPath(path='{path}')` form,
 * which failed every call in 30 days (≈18/18, 6 users): it needs a site id the
 * caller does not have — the whole point of the tool is that they have a URL.
 */
export function resolveSiteByPath(params: Record<string, unknown>): {
  path?: string;
  error?: string;
} {
  const str = (v: unknown) => (typeof v === 'string' ? v.trim() : '');
  let url = str(params.url);
  let hostname = str(params.hostname);
  let sitePath = str(params.path) || str(params.serverRelativePath);
  // A full URL passed as `path`, or a hostname passed in the old site-id slot by
  // a client still holding the pre-tsq.28 schema.
  if (!url && /^https?:\/\//i.test(sitePath)) {
    url = sitePath;
    sitePath = '';
  }
  if (!hostname && /\.sharepoint\./i.test(str(params.siteId))) hostname = str(params.siteId);

  if (url) {
    let parsed: URL;
    try {
      parsed = new URL(/^https?:\/\//i.test(url) ? url : `https://${url}`);
    } catch {
      return { error: `Not a valid URL: ${url}` };
    }
    hostname = parsed.hostname;
    sitePath = parsed.pathname;
  }

  if (!hostname) {
    return {
      error:
        "Pass the site's full URL (url: 'https://contoso.sharepoint.com/sites/marketing'), or " +
        'hostname plus path. To find a site by name instead, use search-sharepoint-sites.',
    };
  }
  if (!/^[a-z0-9-]+(\.[a-z0-9-]+)+$/i.test(hostname)) {
    return { error: `Not a valid SharePoint hostname: ${hostname}` };
  }

  const segments: string[] = [];
  const raw = sitePath.split('/').filter(Boolean);
  // Sharing links put a type marker first: /:f:/r/sites/x/… or /:w:/s/x/…
  if (raw.length >= 2 && /^:[a-z]:$/i.test(raw[0]) && /^[a-z]$/i.test(raw[1])) raw.splice(0, 2);
  for (const segment of raw) {
    const decoded = decodeSegment(segment);
    if (SITE_PATH_STOP_SEGMENTS.has(decoded.toLowerCase()) || /\.aspx$/i.test(decoded)) break;
    segments.push(encodeURIComponent(decoded));
  }

  // No path left means the root site, which Graph addresses by hostname alone.
  return {
    path: segments.length ? `/sites/${hostname}:/${segments.join('/')}` : `/sites/${hostname}`,
  };
}

/**
 * Thin wrapper around the tool execution that records a structured usage record
 * (which tool, by whom, success/error) for periodic tiered-rollout reporting.
 * All exit paths of the impl flow through here, so coverage is uniform.
 */
async function executeGraphTool(
  tool: (typeof api.endpoints)[0],
  config: EndpointConfig | undefined,
  graphClient: GraphClient,
  params: Record<string, unknown>,
  authManager?: AuthManager
): Promise<CallToolResult> {
  // Per-user daily ceiling on deletes and sends (tsq.28, off unless configured).
  // Checked here so every route to a limited tool — direct, execute-tool in
  // discovery mode — passes the same gate, and the refusal is in the usage log.
  const allowance = reserveDailyAllowance(tool.alias);
  if (isDailyLimitRefusal(allowance)) {
    logger.warn(
      `Refusing ${tool.alias}: daily ${allowance.kind} limit of ${allowance.limit} reached`
    );
    const refusal = dailyLimitRefusalResult(tool.alias, allowance);
    logToolUsage(tool.alias, 'error', describeFailure(refusal));
    return refusal;
  }

  let result: CallToolResult;
  try {
    result = await executeGraphToolImpl(tool, config, graphClient, params, authManager);
  } catch (error) {
    allowance?.refund();
    logToolUsage(tool.alias, 'error', describeFailure(error));
    throw error;
  }
  // A failed action did not happen, so it does not use up the day.
  if (result.isError) allowance?.refund();
  logToolUsage(
    tool.alias,
    result.isError ? 'error' : 'success',
    result.isError ? describeFailure(result) : undefined
  );
  return result;
}

async function executeGraphToolImpl(
  tool: (typeof api.endpoints)[0],
  config: EndpointConfig | undefined,
  graphClient: GraphClient,
  params: Record<string, unknown>,
  authManager?: AuthManager
): Promise<CallToolResult> {
  logger.info(`Tool ${tool.alias} called with params: ${JSON.stringify(params)}`);
  try {
    // tsq.28: the generated path for this tool cannot work (see resolveSiteByPath);
    // swap in the hostname:/path form and drop the addressing params so the
    // generic loop below does not try to place them.
    if (tool.alias === 'get-sharepoint-site-by-path') {
      const resolved = resolveSiteByPath(params);
      if (!resolved.path) {
        return errorResult({ error: 'invalid_site_address', message: resolved.error });
      }
      logger.info(`${tool.alias}: resolved to ${resolved.path}`);
      const addressing = new Set(['url', 'hostname', 'path', 'serverRelativePath', 'siteId']);
      tool = { ...tool, path: resolved.path, parameters: [] };
      params = Object.fromEntries(Object.entries(params).filter(([k]) => !addressing.has(k)));
    }

    // Resolve account-specific token if `account` parameter is provided (or auto-resolve for single account).
    // Skip in OAuth/HTTP mode — let the request context drive token selection via GraphClient.
    // Also skip when a request-context token exists (HTTP/OAuth flow where token comes from middleware).
    let accountAccessToken: string | undefined;
    if (authManager && !authManager.isOAuthModeEnabled() && !getRequestTokens()) {
      const accountParam = params.account as string | undefined;
      try {
        accountAccessToken = await authManager.getTokenForAccount(accountParam);
      } catch (err) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: (err as Error).message }),
            },
          ],
          isError: true,
        };
      }
    }

    const parameterDefinitions = tool.parameters || [];

    let path = tool.path;
    const queryParams: Record<string, string> = {};
    const headers: Record<string, string> = {};
    let body: unknown = null;

    for (const [paramName, paramValue] of Object.entries(params)) {
      // Skip control parameters - not part of the Microsoft Graph API
      if (
        [
          'account',
          'fetchAllPages',
          'includeHeaders',
          'excludeResponse',
          'timezone',
          'expandExtendedProperties',
        ].includes(paramName)
      ) {
        continue;
      }

      // Ok, so, MCP clients (such as claude code) doesn't support $ in parameter names,
      // and others might not support __, so we strip them in hack.ts and restore them here
      const odataParams = [
        'filter',
        'select',
        'expand',
        'orderby',
        'skip',
        'top',
        'count',
        'search',
        'format',
      ];
      // Handle both "top" and "$top" formats - strip $ if present, then re-add it
      const normalizedParamName = paramName.startsWith('$') ? paramName.slice(1) : paramName;
      const isOdataParam = odataParams.includes(normalizedParamName.toLowerCase());
      const fixedParamName = isOdataParam ? `$${normalizedParamName.toLowerCase()}` : paramName;
      // Convert kebab-case param names to camelCase for path param matching.
      // endpoints.json uses {message-id} but hack.ts extracts :messageId (camelCase) from the path.
      // LLMs may pass "message-id" (kebab) — we normalize so both forms work.
      const camelCaseParamName = paramName.replace(/-([a-z])/g, (_, c: string) => c.toUpperCase());

      // Look up param definition using normalized name (without $) for OData params,
      // or camelCase equivalent for kebab-case path params
      const paramDef = parameterDefinitions.find(
        (p) =>
          p.name === paramName ||
          p.name === camelCaseParamName ||
          (isOdataParam && p.name === normalizedParamName)
      );

      if (paramDef) {
        switch (paramDef.type) {
          case 'Path': {
            // Check if this parameter should skip URL encoding (for function-style API calls)
            const shouldSkipEncoding = config?.skipEncoding?.includes(paramName) ?? false;
            // Use encodeURIComponent but preserve '=' which is valid in path segments (RFC 3986)
            // and commonly appears in Microsoft Graph base64-encoded resource IDs.
            // Without this, IDs like "AAMk...AAA=" become "AAMk...AAA%3D" causing 404 errors.
            // First we encode, then unencode. Crazy, check out https://github.com/Softeria/ms-365-mcp-server/issues/245
            const encodedValue = shouldSkipEncoding
              ? (paramValue as string)
              : encodeURIComponent(paramValue as string).replace(/%3D/g, '=');

            // Replace both the original param name and the camelCase variant
            // to handle {message-id} (endpoints.json) and :messageId (generated client) formats
            path = path
              .replace(`{${paramName}}`, encodedValue)
              .replace(`:${paramName}`, encodedValue)
              .replace(`{${camelCaseParamName}}`, encodedValue)
              .replace(`:${camelCaseParamName}`, encodedValue);
            break;
          }

          case 'Query':
            if (paramValue !== '' && paramValue != null) {
              queryParams[fixedParamName] = `${paramValue}`;
            }
            break;

          case 'Body':
            if (paramDef.schema) {
              const parseResult = paramDef.schema.safeParse(paramValue);
              if (!parseResult.success) {
                const wrapped = { [paramName]: paramValue };
                const wrappedResult = paramDef.schema.safeParse(wrapped);
                if (wrappedResult.success) {
                  logger.info(
                    `Auto-corrected parameter '${paramName}': AI passed nested field directly, wrapped it as {${paramName}: ...}`
                  );
                  body = wrapped;
                } else {
                  body = paramValue;
                }
              } else {
                body = paramValue;
              }
            } else {
              body = paramValue;
            }
            break;

          case 'Header':
            headers[fixedParamName] = `${paramValue}`;
            break;
        }
      } else if (paramName === 'body') {
        body = paramValue;
        logger.info(`Set body param: ${JSON.stringify(body)}`);
      } else if (
        path.includes(`:${paramName}`) ||
        path.includes(`{${paramName}}`) ||
        path.includes(`:${camelCaseParamName}`) ||
        path.includes(`{${camelCaseParamName}}`)
      ) {
        // Fallback: path param not declared in tool.parameters (generated client omits them).
        // Replace placeholder directly so the URL is valid.
        const encodedValue = encodeURIComponent(paramValue as string).replace(/%3D/g, '=');
        path = path
          .replace(`{${paramName}}`, encodedValue)
          .replace(`:${paramName}`, encodedValue)
          .replace(`{${camelCaseParamName}}`, encodedValue)
          .replace(`:${camelCaseParamName}`, encodedValue);
        logger.info(`Path param fallback: replaced :${camelCaseParamName} with encoded value`);
      }
    }

    // Defense-in-depth for Bug 2: Graph rejects $top on /delta() endpoints with
    // HTTP 400. The user-facing schema for -delta tools strips top/$top, so
    // freshly-connected clients can't send it. But cached/stale clients (and
    // ad-hoc callers) might still try. Drop it server-side before clamping or
    // sending, regardless of where it came from.
    if (tool.alias.endsWith('-delta')) {
      delete queryParams['$top'];
    }

    clampTopQueryParam(queryParams);
    const searchError = normalizeSearchQueryParam(queryParams, tool.path, tool.alias);
    if (searchError) return searchError;
    const collectionQuery = normalizeCollectionQuery(tool.alias, tool.path, queryParams);
    if (collectionQuery.error) return collectionQuery.error;

    const normalizedBody = normalizeWriteBody(tool.alias, body);
    if (normalizedBody.error) return normalizedBody.error;
    body = normalizedBody.body;

    if (params.trackReply === true && TRACK_AT_SEND_TOOLS[tool.alias]) {
      body = applyTrackAtSend(tool.alias, body, {
        dueBy: typeof params.trackDueBy === 'string' ? params.trackDueBy : undefined,
        waitingOn: Array.isArray(params.trackWaitingOn)
          ? (params.trackWaitingOn as unknown[]).filter((v): v is string => typeof v === 'string')
          : undefined,
      });
    }

    const preferValues: string[] = [];

    // Handle timezone parameter for calendar endpoints
    if (config?.supportsTimezone && params.timezone) {
      preferValues.push(`outlook.timezone="${params.timezone}"`);
      logger.info(`Setting timezone preference: outlook.timezone="${params.timezone}"`);
    }

    const bodyFormat = process.env.MS365_MCP_BODY_FORMAT || 'text';
    if (bodyFormat !== 'html' && tool.method.toUpperCase() === 'GET') {
      preferValues.push(`outlook.body-content-type="${bodyFormat}"`);
    }

    if (preferValues.length > 0) {
      headers['Prefer'] = preferValues.join(', ');
    }

    // Handle expandExtendedProperties parameter for calendar endpoints
    if (config?.supportsExpandExtendedProperties && params.expandExtendedProperties === true) {
      const expandValue = 'singleValueExtendedProperties';
      if (queryParams['$expand']) {
        queryParams['$expand'] += `,${expandValue}`;
      } else {
        queryParams['$expand'] = expandValue;
      }
      logger.info(`Adding $expand=${expandValue} for extended properties`);
    }

    if (config?.contentType) {
      headers['Content-Type'] = config.contentType;
      logger.info(`Setting custom Content-Type: ${config.contentType}`);
    }

    if (config?.acceptType) {
      headers['Accept'] = config.acceptType;
      logger.info(`Setting custom Accept: ${config.acceptType}`);
    }

    if (Object.keys(queryParams).length > 0) {
      const queryString = Object.entries(queryParams)
        .map(([key, value]) => `${key}=${encodeURIComponent(value).replace(/%2C/gi, ',')}`)
        .join('&');
      path = `${path}${path.includes('?') ? '&' : '?'}${queryString}`;
    }

    const options: {
      method: string;
      headers: Record<string, string>;
      body?: string;
      rawResponse?: boolean;
      includeHeaders?: boolean;
      excludeResponse?: boolean;
      queryParams?: Record<string, string>;
      accessToken?: string;
    } = {
      method: tool.method.toUpperCase(),
      headers,
    };

    if (options.method !== 'GET' && body) {
      if (config?.contentType === 'text/html') {
        if (typeof body === 'string') {
          options.body = body;
        } else if (typeof body === 'object' && 'content' in body) {
          options.body = (body as { content: string }).content;
        } else {
          options.body = String(body);
        }
      } else {
        options.body = typeof body === 'string' ? body : JSON.stringify(body);
      }
    }

    const isProbablyMediaContent =
      tool.errors?.some((error) => error.description === 'Retrieved media content') ||
      path.endsWith('/content');

    if (config?.returnDownloadUrl && path.endsWith('/content')) {
      path = path.replace(/\/content$/, '');
      logger.info(
        `Auto-returning download URL for ${tool.alias} (returnDownloadUrl=true in endpoints.json)`
      );
    } else if (isProbablyMediaContent) {
      options.rawResponse = true;
    }

    // Set includeHeaders if requested
    if (params.includeHeaders === true) {
      options.includeHeaders = true;
    }

    // Set excludeResponse if requested
    if (params.excludeResponse === true) {
      options.excludeResponse = true;
    }

    // Pass account-resolved token if available
    if (accountAccessToken) {
      options.accessToken = accountAccessToken;
    }

    // Redact accessToken from log output to prevent credential leakage
    const { accessToken: _redacted, ...safeOptions } = options;
    logger.info(
      `Making graph request to ${path} with options: ${JSON.stringify(safeOptions)}${_redacted ? ' [accessToken=REDACTED]' : ''}`
    );

    let response = await graphClient.graphRequest(path, options);

    const fetchAllPages = params.fetchAllPages === true;
    if (fetchAllPages && response?.content?.[0]?.text) {
      try {
        let combinedResponse = JSON.parse(response.content[0].text);
        let allItems = combinedResponse.value || [];
        let nextLink = combinedResponse['@odata.nextLink'];
        let pageCount = 1;
        const maxPages = 100;
        const maxItems = 10_000;

        let deltaLink: string | undefined = combinedResponse['@odata.deltaLink'];
        let bailedOut: 'maxPages' | 'maxItems' | null = null;
        while (nextLink) {
          // Check caps BEFORE fetching the next page so the loop exits cleanly
          // and the last-known nextLink is preserved for caller resume. Previously
          // the cap check was in the while condition, which combined with the
          // unconditional `delete @odata.nextLink` below stranded callers mid-sync
          // (value array returned, but no resume token).
          if (pageCount >= maxPages) {
            bailedOut = 'maxPages';
            break;
          }
          if (allItems.length >= maxItems) {
            bailedOut = 'maxItems';
            break;
          }

          logger.info(`Fetching page ${pageCount + 1} from: ${nextLink}`);

          // Extract path + query string from the nextLink URL.
          // Pass the full path (with query string) as the endpoint so that
          // $skiptoken and other pagination params are preserved.
          // Previously, query params were extracted into nextOptions.queryParams
          // but graphRequest/performRequest never read that field — they were lost.
          const url = new URL(nextLink);
          const nextPath = url.pathname.replace('/v1.0', '') + url.search;
          const nextOptions = { ...options };

          const nextResponse = await graphClient.graphRequest(nextPath, nextOptions);
          if (nextResponse?.content?.[0]?.text) {
            const nextJsonResponse = JSON.parse(nextResponse.content[0].text);
            if (nextJsonResponse.value && Array.isArray(nextJsonResponse.value)) {
              allItems = allItems.concat(nextJsonResponse.value);
            }
            nextLink = nextJsonResponse['@odata.nextLink'];
            if (nextJsonResponse['@odata.deltaLink']) {
              deltaLink = nextJsonResponse['@odata.deltaLink'];
            }
            pageCount++;
          } else {
            break;
          }
        }

        if (bailedOut) {
          logger.warn(
            `fetchAllPages hit cap '${bailedOut}' (pages=${pageCount}/${maxPages}, ` +
              `items=${allItems.length}/${maxItems}); preserving @odata.nextLink for caller resume.`
          );
        }

        combinedResponse.value = allItems;
        if (combinedResponse['@odata.count']) {
          combinedResponse['@odata.count'] = allItems.length;
        }

        // Only strip @odata.nextLink if the loop exited because Graph stopped
        // emitting it. If we bailed out due to caps, preserve the last-known
        // nextLink so the caller can resume.
        if (bailedOut && nextLink) {
          combinedResponse['@odata.nextLink'] = nextLink;
        } else {
          delete combinedResponse['@odata.nextLink'];
        }

        // Carry the @odata.deltaLink from the final page so callers can resume
        // a delta sync. Without this, fetchAllPages on a /delta endpoint silently
        // drops the resume token and forces callers to re-list from scratch.
        if (deltaLink) {
          combinedResponse['@odata.deltaLink'] = deltaLink;
        }

        // Diagnostic marker so callers can detect partial results and the reason.
        if (bailedOut) {
          combinedResponse['_tsq_pagingBailedOut'] = {
            reason: bailedOut,
            pageCount,
            itemCount: allItems.length,
            maxPages,
            maxItems,
            note: 'Returned partial results due to pagination cap. Use @odata.nextLink to resume.',
          };
        }

        response.content[0].text = JSON.stringify(combinedResponse);

        logger.info(
          `Pagination complete: collected ${allItems.length} items across ${pageCount} pages` +
            (bailedOut ? ` (bailed out on ${bailedOut})` : '')
        );
      } catch (e) {
        logger.error(`Error during pagination: ${e}`);
      }
    }

    if (response?.content?.[0]?.text) {
      const responseText = response.content[0].text;
      logger.info(`Response size: ${responseText.length} characters`);

      try {
        const jsonResponse = JSON.parse(responseText);
        if (jsonResponse.value && Array.isArray(jsonResponse.value)) {
          logger.info(`Response contains ${jsonResponse.value.length} items`);
        }
        if (jsonResponse['@odata.nextLink']) {
          logger.info(`Response has pagination nextLink: ${jsonResponse['@odata.nextLink']}`);
        }
      } catch {
        // Non-JSON response
      }
    }

    const alreadyDeleted = alreadyDeletedResult(tool, response);
    if (alreadyDeleted) return alreadyDeleted;

    // Convert McpResponse to CallToolResult with the correct structure
    const content: ContentItem[] = response.content.map((item) => ({
      type: 'text' as const,
      text: item.text,
    }));

    const converted: CallToolResult = {
      content,
      _meta: response._meta,
      isError: response.isError,
    };
    if (collectionQuery.clientOrderby && !converted.isError) {
      return sortResultPage(
        converted,
        collectionQuery.clientOrderby,
        'Sorted client-side by the MCP server because Graph rejects this $orderby with this $filter (InefficientFilter). Only the returned page is sorted; pass fetchAllPages for a complete ordering.',
        tool.alias
      );
    }
    return converted;
  } catch (error) {
    logger.error(`Error in tool ${tool.alias}: ${(error as Error).message}`);
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            error: `Error in tool ${tool.alias}: ${(error as Error).message}`,
          }),
        },
      ],
      isError: true,
    };
  }
}

export function registerGraphTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false,
  enabledToolsPattern?: string,
  orgMode: boolean = false,
  authManager?: AuthManager,
  multiAccount: boolean = false,
  accountNames: string[] = [],
  toolAllowlist?: string[]
): number {
  // An explicit named allowlist (Phase 1 tiered rollout) wins over the
  // ENABLED_TOOLS regex. Exact-match, no patterns.
  const allowSet = toolAllowlist ? new Set(toolAllowlist) : undefined;

  let enabledToolsRegex: RegExp | undefined;
  if (!allowSet && enabledToolsPattern) {
    try {
      enabledToolsRegex = new RegExp(enabledToolsPattern, 'i');
      logger.info(`Tool filtering enabled with pattern: ${enabledToolsPattern}`);
    } catch {
      logger.error(`Invalid tool filter regex pattern: ${enabledToolsPattern}. Ignoring filter.`);
    }
  }

  // Single predicate used everywhere a tool is gated (endpoint tools and the
  // hand-written custom tools below) so the allowlist and the regex never diverge.
  const isToolEnabled = (toolName: string): boolean => {
    if (allowSet) return allowSet.has(toolName);
    if (enabledToolsRegex) return enabledToolsRegex.test(toolName);
    return true;
  };

  if (allowSet) {
    logger.info(`Tool allowlist enabled: ${allowSet.size} tools (exact match)`);
  }

  const registeredNames: string[] = [];
  let registeredCount = 0;
  let skippedCount = 0;
  let failedCount = 0;

  for (const tool of api.endpoints) {
    const endpointConfig = endpointsData.find((e) => e.toolName === tool.alias);
    if (!orgMode && endpointConfig && !endpointConfig.scopes && endpointConfig.workScopes) {
      logger.info(`Skipping work account tool ${tool.alias} - not in org mode`);
      skippedCount++;
      continue;
    }

    const method = tool.method.toUpperCase();
    if (readOnly && method !== 'GET') {
      // Allow POST endpoints that are explicitly marked as readOnly in endpoints.json
      // (e.g. get-schedule, find-meeting-times which are read-only queries via POST).
      // PATCH/DELETE are always blocked in read-only mode.
      if (!(method === 'POST' && endpointConfig?.readOnly)) {
        logger.info(`Skipping write operation ${tool.alias} in read-only mode`);
        skippedCount++;
        continue;
      }
    }

    if (!isToolEnabled(tool.alias)) {
      logger.info(`Skipping tool ${tool.alias} - not in allowlist / doesn't match filter`);
      skippedCount++;
      continue;
    }

    const paramSchema: Record<string, z.ZodTypeAny> = {};
    if (tool.parameters && tool.parameters.length > 0) {
      for (const param of tool.parameters) {
        // A configured compactBody replaces the inlined Graph entity schema; see
        // EndpointConfig.compactBody. Matched on the parameter's Body type, with a
        // name fallback so a client-generator change cannot silently disable it.
        const isBody = param.type === 'Body' || param.name === 'body';
        if (isBody && endpointConfig?.compactBody) {
          paramSchema[param.name] = z.object({}).passthrough().describe(endpointConfig.compactBody);
          continue;
        }
        paramSchema[param.name] = param.schema || z.any();
      }
    }

    // Graph rejects $top on /delta() endpoints (HTTP 400). Page size is controlled
    // internally by Graph or via the Prefer: odata.maxpagesize header (not exposed
    // by this server). Strip $top/top from delta tool schemas so the model never
    // sends it.
    const isDeltaTool = tool.alias.endsWith('-delta');
    if (isDeltaTool) {
      delete paramSchema['top'];
      delete paramSchema['$top'];
    }

    // Extract path parameters from the path pattern (e.g., :todoTaskListId from /me/todo/lists/:todoTaskListId/tasks)
    // The generated client omits these from tool.parameters, so we add them manually.
    const pathParamMatches = tool.path.matchAll(/:([a-zA-Z]+)/g);
    for (const match of pathParamMatches) {
      const pathParamName = match[1];
      if (!(pathParamName in paramSchema)) {
        paramSchema[pathParamName] = z.string().describe(`Path parameter: ${pathParamName}`);
      }
    }

    // tsq.28: the generated schema asks for a site id plus a path, and the site id
    // is exactly what a caller of this tool does not have. The executor builds the
    // real address from a URL (or hostname + path) — see resolveSiteByPath.
    if (tool.alias === 'get-sharepoint-site-by-path') {
      delete paramSchema['siteId'];
      paramSchema['url'] = z
        .string()
        .describe(
          'Full site URL, e.g. https://contoso.sharepoint.com/sites/marketing. A page or library URL inside the site works too.'
        )
        .optional();
      paramSchema['hostname'] = z
        .string()
        .describe('Instead of url: the SharePoint hostname, e.g. contoso.sharepoint.com.')
        .optional();
      paramSchema['path'] = z
        .string()
        .describe('With hostname: the server-relative site path, e.g. /sites/marketing.')
        .optional();
    }

    if (tool.method.toUpperCase() === 'GET' && tool.path.includes('/')) {
      paramSchema['fetchAllPages'] = z
        .boolean()
        .describe('Merge all pages (max 100). Only for a genuine full export.')
        .optional();
    }

    // Override OData parameter descriptions with spec-gap guidance
    if (paramSchema['filter'] !== undefined || paramSchema['$filter'] !== undefined) {
      const key = paramSchema['$filter'] !== undefined ? '$filter' : 'filter';
      paramSchema[key] = z
        .string()
        .describe('OData filter. Not with search. contains() needs count=true.')
        .optional();
    }
    if (paramSchema['search'] !== undefined || paramSchema['$search'] !== undefined) {
      const key = paramSchema['$search'] !== undefined ? '$search' : 'search';
      paramSchema[key] = z
        .string()
        .describe(
          'KQL query in one pair of double quotes; directory (users/groups) instead quotes each clause with no outer pair. Not with filter.'
        )
        .optional();
    }
    if (paramSchema['select'] !== undefined || paramSchema['$select'] !== undefined) {
      const key = paramSchema['$select'] !== undefined ? '$select' : 'select';
      paramSchema[key] = z.string().describe('Fields to return, comma-separated.').optional();
    }
    if (paramSchema['orderby'] !== undefined || paramSchema['$orderby'] !== undefined) {
      const key = paramSchema['$orderby'] !== undefined ? '$orderby' : 'orderby';
      paramSchema[key] = z
        .string()
        .describe('Sort expression, e.g. receivedDateTime desc')
        .optional();
    }
    if (paramSchema['top'] !== undefined || paramSchema['$top'] !== undefined) {
      const key = paramSchema['$top'] !== undefined ? '$top' : 'top';
      paramSchema[key] = z.number().describe('Page size. Start small, 5-15.').optional();
    }
    if (paramSchema['skip'] !== undefined || paramSchema['$skip'] !== undefined) {
      const key = paramSchema['$skip'] !== undefined ? '$skip' : 'skip';
      paramSchema[key] = z.number().describe('Items to skip. Not with search.').optional();
    }
    if (paramSchema['count'] !== undefined || paramSchema['$count'] !== undefined) {
      const countKey = paramSchema['$count'] !== undefined ? '$count' : 'count';
      paramSchema[countKey] = z
        .boolean()
        .describe('Advanced query mode (ConsistencyLevel: eventual).')
        .optional();
    }

    // Add account parameter for multi-account mode.
    // Layer 2: Account names are surfaced in the description (not as a strict enum) so the LLM
    // sees available accounts upfront without a round-trip, but accounts added mid-session via
    // --login are still accepted — getTokenForAccount() handles validation at runtime.
    if (multiAccount) {
      const accountHint =
        accountNames.length > 0 ? `Known accounts: ${accountNames.join(', ')}. ` : '';
      paramSchema['account'] = z
        .string()
        .describe(
          `${accountHint}Microsoft account email to use for this request. ` +
            `Required when multiple accounts are configured. ` +
            `Use the list-accounts tool to discover all currently available accounts.`
        )
        .optional();
    }

    // Add includeHeaders parameter for all tools to capture ETags and other headers
    paramSchema['includeHeaders'] = z
      .boolean()
      .describe('Also return response headers, e.g. ETag.')
      .optional();

    // Add excludeResponse parameter to only return success/failure indication
    paramSchema['excludeResponse'] = z
      .boolean()
      .describe('Return only success or failure, not the body.')
      .optional();

    // Track at send (tsq.28): available on every tool that creates a message.
    if (TRACK_AT_SEND_TOOLS[tool.alias]) {
      paramSchema['trackReply'] = z
        .boolean()
        .describe(
          'Track this message for a reply, in the same call: list-awaiting-replies will report ' +
            'it in any later conversation. Use whenever the user expects an answer.'
        )
        .optional();
      paramSchema['trackDueBy'] = z
        .string()
        .describe('With trackReply: ISO date after which it is reported overdue.')
        .optional();
      paramSchema['trackWaitingOn'] = z
        .array(z.string())
        .describe("With trackReply: only these addresses' replies count. Default: anyone but you.")
        .optional();
    }

    // Add timezone parameter for calendar endpoints that support it
    if (endpointConfig?.supportsTimezone) {
      paramSchema['timezone'] = z
        .string()
        .describe(
          'IANA timezone name (e.g., "America/New_York", "Europe/London", "Asia/Tokyo") for calendar event times. If not specified, times are returned in UTC.'
        )
        .optional();
    }

    // Add expandExtendedProperties parameter for calendar endpoints that support it
    if (endpointConfig?.supportsExpandExtendedProperties) {
      paramSchema['expandExtendedProperties'] = z
        .boolean()
        .describe(
          'When true, expands singleValueExtendedProperties on each event. Use this to retrieve custom extended properties (e.g., sync metadata) stored on calendar events.'
        )
        .optional();
    }

    // Build the tool description, optionally appending LLM tips. A hand-written
    // `description` in endpoints.json wins over the upstream Graph prose — see
    // EndpointConfig.description for why that override exists.
    let toolDescription =
      endpointConfig?.description ||
      tool.description ||
      `Execute ${tool.method.toUpperCase()} request to ${tool.path}`;
    if (endpointConfig?.llmTip) {
      toolDescription += `\n\n💡 TIP: ${endpointConfig.llmTip}`;
    }

    try {
      server.tool(
        tool.alias,
        toolDescription,
        paramSchema,
        {
          title: tool.alias,
          readOnlyHint: tool.method.toUpperCase() === 'GET',
          destructiveHint: ['POST', 'PATCH', 'DELETE'].includes(tool.method.toUpperCase()),
          openWorldHint: true, // All tools call Microsoft Graph API
        },
        async (params) => executeGraphTool(tool, endpointConfig, graphClient, params, authManager)
      );
      registeredNames.push(tool.alias);
      registeredCount++;
    } catch (error) {
      logger.error(`Failed to register tool ${tool.alias}: ${(error as Error).message}`);
      failedCount++;
    }
  }

  if (multiAccount) {
    logger.info('Multi-account mode: "account" parameter injected into all tool schemas');
  }

  // Register parse-teams-url utility tool (no Graph API call)
  if (isToolEnabled('parse-teams-url')) {
    try {
      server.tool(
        'parse-teams-url',
        'Converts any Teams meeting URL format (short /meet/, full /meetup-join/, or recap ?threadId=) into a standard joinWebUrl. Use this before list-online-meetings when the user provides a recap or short URL.',
        {
          url: z.string().describe('Teams meeting URL in any format'),
        },
        {
          title: 'parse-teams-url',
          readOnlyHint: true,
          openWorldHint: false,
        },
        async ({ url }) => {
          try {
            const joinWebUrl = parseTeamsUrl(url);
            return { content: [{ type: 'text', text: joinWebUrl }] };
          } catch (error) {
            return {
              content: [
                { type: 'text', text: JSON.stringify({ error: (error as Error).message }) },
              ],
              isError: true,
            };
          }
        }
      );
      registeredNames.push('parse-teams-url');
      registeredCount++;
    } catch (error) {
      logger.error(`Failed to register tool parse-teams-url: ${(error as Error).message}`);
      failedCount++;
    }
  }

  // list-conversation-messages and list-drafts are registered as custom tools because
  // their Graph URLs collide with existing endpoints.json paths
  // (/me/messages and /me/mailFolders/{id}/messages). The simplified-openapi
  // generator collapses path+method duplicates into a single operationId, so a
  // second endpoints.json entry on the same path would never reach the client.
  // Both delegate to executeGraphTool with a synthetic Endpoint so they inherit
  // pagination, auth, OData translation, and $top clamping.

  const mailListReadOnlyHints = {
    readOnlyHint: true,
    destructiveHint: false,
    openWorldHint: true,
  };

  if (isToolEnabled('list-conversation-messages')) {
    try {
      const conversationMessagesTool = {
        alias: 'list-conversation-messages',
        method: 'get' as const,
        path: '/me/messages',
        requestFormat: 'json' as const,
        parameters: [
          { name: 'filter', type: 'Query' as const, schema: z.string() },
          { name: 'select', type: 'Query' as const, schema: z.string() },
          { name: 'orderby', type: 'Query' as const, schema: z.string() },
          { name: 'top', type: 'Query' as const, schema: z.number() },
          { name: 'count', type: 'Query' as const, schema: z.boolean() },
          { name: 'ConsistencyLevel', type: 'Header' as const, schema: z.string() },
        ],
        response: z.any(),
      };
      const conversationMessagesConfig: EndpointConfig = {
        pathPattern: '/me/messages',
        method: 'get',
        toolName: 'list-conversation-messages',
        scopes: ['Mail.Read'],
        workScopes: ['Mail.Read'],
      };

      server.tool(
        'list-conversation-messages',
        'Lists every message in a single email conversation thread by conversationId. ' +
          'The conversationId is stable across subject line changes — get it from any ' +
          "message's conversationId field (get-mail-message, list-mail-messages, etc.). " +
          "Backed by /me/messages with $filter=conversationId eq '{id}' (the server adds " +
          'this filter for you, so callers must NOT pass their own filter parameter). ' +
          'Always pass `select` to limit returned fields ' +
          '(recommended: id,subject,from,toRecipients,receivedDateTime,bodyPreview,conversationId). ' +
          'To sort results, pass `orderby` (e.g. "receivedDateTime asc" or ' +
          '"receivedDateTime desc"). Microsoft Graph rejects $filter=conversationId + ' +
          '$orderby even in advanced query mode (InefficientFilter), so the server sorts ' +
          'the returned page(s) client-side. For correct sort order across long threads, ' +
          'also pass `fetchAllPages: true` so the full set is sorted, not just the first page. ' +
          "The `count: true` parameter enables Graph's advanced query mode (ConsistencyLevel: " +
          'eventual + $count=true). It is NOT required for orderby (the server handles that ' +
          'client-side) but may be useful when combined with other advanced query features.',
        {
          conversationId: z
            .string()
            .describe(
              "The conversationId to filter by. Read it from any message's conversationId field " +
                '(get-mail-message, list-mail-messages, etc.).'
            ),
          select: z
            .string()
            .optional()
            .describe(
              'Comma-separated fields to return, e.g. id,subject,from,toRecipients,receivedDateTime,bodyPreview,conversationId. ' +
                'Always set this to keep responses small.'
            ),
          orderby: z
            .string()
            .optional()
            .describe(
              "Sort expression, e.g. 'receivedDateTime asc' or 'receivedDateTime desc'. " +
                'The server sorts client-side because Graph rejects $filter + $orderby on /me/messages. ' +
                'For complete sort across long threads, also pass fetchAllPages: true.'
            ),
          top: z
            .number()
            .optional()
            .describe(
              'Page size (Graph $top). Start small (5–15) so responses fit context; raise only if needed.'
            ),
          count: z
            .boolean()
            .optional()
            .describe(
              'Set true to enable Graph advanced query mode (ConsistencyLevel: eventual + $count=true). ' +
                'NOT required for orderby (the server sorts client-side). Useful when combining with other advanced query features.'
            ),
          fetchAllPages: z
            .boolean()
            .optional()
            .describe(
              'Follow @odata.nextLink and merge up to 100 pages into one response. ' +
                'Can return enormous payloads — only when the user explicitly needs the whole thread.'
            ),
        },
        {
          title: 'list-conversation-messages',
          ...mailListReadOnlyHints,
        },
        async (params) => {
          const conversationId = String(params.conversationId ?? '').trim();
          if (!conversationId) {
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify({
                    error: 'conversationId is required and must be non-empty.',
                  }),
                },
              ],
              isError: true,
            };
          }
          // Capture the requested orderby BEFORE we strip it from the outbound call.
          // Graph rejects $filter=conversationId + $orderby with InefficientFilter even
          // in advanced query mode, so we sort client-side after fetch.
          const requestedOrderby = params.orderby;

          // Escape single quotes inside the OData string literal by doubling them.
          const escapedId = conversationId.replace(/'/g, "''");
          const callParams: Record<string, unknown> = {
            filter: `conversationId eq '${escapedId}'`,
          };
          if (params.select !== undefined) callParams.select = params.select;
          // NOTE: deliberately NOT setting orderby on callParams — see comment above.
          if (params.top !== undefined) callParams.top = params.top;
          if (params.count === true) {
            callParams.count = true;
            callParams.ConsistencyLevel = 'eventual';
          }
          if (params.fetchAllPages !== undefined) callParams.fetchAllPages = params.fetchAllPages;

          const result = await executeGraphTool(
            conversationMessagesTool,
            conversationMessagesConfig,
            graphClient,
            callParams,
            authManager
          );

          if (!requestedOrderby) {
            return result;
          }

          return sortResultPage(
            result,
            requestedOrderby,
            'Sorted client-side by the MCP server because Graph rejects this $filter + $orderby combination with InefficientFilter, even in advanced query mode.',
            'list-conversation-messages'
          );
        }
      );
      registeredNames.push('list-conversation-messages');
      registeredCount++;
    } catch (error) {
      logger.error(
        `Failed to register tool list-conversation-messages: ${(error as Error).message}`
      );
      failedCount++;
    }
  }

  if (isToolEnabled('list-drafts')) {
    try {
      const draftsTool = {
        alias: 'list-drafts',
        method: 'get' as const,
        path: '/me/mailFolders/drafts/messages',
        requestFormat: 'json' as const,
        parameters: [
          { name: 'select', type: 'Query' as const, schema: z.string() },
          { name: 'orderby', type: 'Query' as const, schema: z.string() },
          { name: 'top', type: 'Query' as const, schema: z.number() },
        ],
        response: z.any(),
      };
      const draftsConfig: EndpointConfig = {
        pathPattern: '/me/mailFolders/drafts/messages',
        method: 'get',
        toolName: 'list-drafts',
        scopes: ['Mail.Read'],
        workScopes: ['Mail.Read'],
      };

      server.tool(
        'list-drafts',
        "Lists unsent draft emails from the user's Drafts folder. Backed by " +
          '/me/mailFolders/drafts/messages (well-known folder name). Useful for auditing ' +
          'unsent drafts or resuming a draft started earlier. Always pass `select` to keep ' +
          'responses small (recommended: id,subject,toRecipients,createdDateTime,lastModifiedDateTime,bodyPreview). ' +
          'Use orderby=lastModifiedDateTime desc to see most recently edited drafts first.',
        {
          select: z
            .string()
            .optional()
            .describe(
              'Comma-separated fields to return, e.g. id,subject,toRecipients,createdDateTime,lastModifiedDateTime,bodyPreview.'
            ),
          orderby: z
            .string()
            .optional()
            .describe(
              "Sort expression, e.g. 'lastModifiedDateTime desc' or 'createdDateTime desc'."
            ),
          top: z
            .number()
            .optional()
            .describe(
              'Page size (Graph $top). Start small (5–15) so responses fit context; raise only if needed.'
            ),
          fetchAllPages: z
            .boolean()
            .optional()
            .describe(
              'Follow @odata.nextLink and merge up to 100 pages into one response. Only use if the user needs every draft.'
            ),
        },
        {
          title: 'list-drafts',
          ...mailListReadOnlyHints,
        },
        async (params) => {
          const callParams: Record<string, unknown> = {};
          if (params.select !== undefined) callParams.select = params.select;
          if (params.orderby !== undefined) callParams.orderby = params.orderby;
          if (params.top !== undefined) callParams.top = params.top;
          if (params.fetchAllPages !== undefined) callParams.fetchAllPages = params.fetchAllPages;
          return executeGraphTool(draftsTool, draftsConfig, graphClient, callParams, authManager);
        }
      );
      registeredNames.push('list-drafts');
      registeredCount++;
    } catch (error) {
      logger.error(`Failed to register tool list-drafts: ${(error as Error).message}`);
      failedCount++;
    }
  }

  if (isToolEnabled('get-messages-batch')) {
    try {
      server.tool(
        'get-messages-batch',
        'Fetches multiple email messages in a single Graph /$batch request. ' +
          'Much faster than calling get-mail-message N times when you need ' +
          'full bodies for several messages (e.g., daily digests, thread ' +
          'summaries, audit reports). Up to 20 messages per call (Graph $batch limit). ' +
          'Pass `select` to limit returned fields and keep payloads small ' +
          '(recommended: id,subject,from,toRecipients,receivedDateTime,bodyPreview,body). ' +
          'Returns { results, failures, summary }: results[] holds the successful ' +
          'sub-responses each tagged with their messageId; failures[] holds any ' +
          'sub-requests Graph returned non-2xx for (e.g. unknown message ID). The ' +
          'outer call still returns HTTP 200 — Graph processes the batch and ' +
          'returns partial success.',
        {
          messageIds: z
            .array(z.string())
            .min(1)
            .max(20)
            .describe(
              'List of message IDs to fetch. Maximum 20 per call (Graph $batch limit). ' +
                'Get IDs from list-mail-messages, list-conversation-messages, search-query, etc.'
            ),
          select: z
            .string()
            .optional()
            .describe(
              'Comma-separated fields to return for each message. Strongly recommended ' +
                'to keep responses small. Example: id,subject,from,receivedDateTime,body.'
            ),
        },
        {
          title: 'get-messages-batch',
          ...mailListReadOnlyHints,
        },
        async (params) =>
          withUsageLog('get-messages-batch', async () => {
            const messageIds = params.messageIds.filter(
              (id) => typeof id === 'string' && id.trim()
            );
            if (messageIds.length === 0) {
              return {
                content: [
                  {
                    type: 'text',
                    text: JSON.stringify({
                      error: 'messageIds must contain at least one non-empty ID.',
                    }),
                  },
                ],
                isError: true,
              };
            }
            if (messageIds.length > 20) {
              return {
                content: [
                  {
                    type: 'text',
                    text: JSON.stringify({
                      error: 'messageIds limited to 20 per call (Graph $batch limit).',
                    }),
                  },
                ],
                isError: true,
              };
            }

            // Build the per-message URL. Graph $batch URLs are relative to /v1.0 and
            // may include a query string. encodeURIComponent escapes '=' in message
            // IDs to %3D — Graph accepts both forms but unescaping keeps logs readable.
            const selectClause = params.select
              ? `?$select=${encodeURIComponent(params.select).replace(/%2C/gi, ',')}`
              : '';
            const subRequests = messageIds.map((id, idx) => ({
              id: String(idx + 1),
              method: 'GET',
              url: `/me/messages/${encodeURIComponent(id).replace(/%3D/g, '=')}${selectClause}`,
            }));

            try {
              const batchResponse = await graphClient.batchRequest(subRequests);

              const results: Array<{ messageId: string; data: unknown }> = [];
              const failures: Array<{ messageId: string; status: number; error: unknown }> = [];

              for (const subResp of batchResponse.responses ?? []) {
                const idx = parseInt(subResp.id, 10) - 1;
                const messageId = messageIds[idx];
                if (subResp.status >= 200 && subResp.status < 300) {
                  results.push({ messageId, data: subResp.body });
                } else {
                  failures.push({ messageId, status: subResp.status, error: subResp.body });
                }
              }

              return {
                content: [
                  {
                    type: 'text',
                    text: JSON.stringify({
                      results,
                      failures,
                      summary: {
                        total: messageIds.length,
                        succeeded: results.length,
                        failed: failures.length,
                      },
                    }),
                  },
                ],
              };
            } catch (err) {
              logger.error(`get-messages-batch failed: ${(err as Error).message}`);
              return {
                content: [
                  {
                    type: 'text',
                    text: JSON.stringify({ error: (err as Error).message }),
                  },
                ],
                isError: true,
              };
            }
          })
      );
      registeredNames.push('get-messages-batch');
      registeredCount++;
    } catch (error) {
      logger.error(`Failed to register tool get-messages-batch: ${(error as Error).message}`);
      failedCount++;
    }
  }

  // download-mail-attachment writes to OneDrive (staging upload), so it is a
  // write tool — skip it entirely in read-only mode, mirroring the endpoint
  // loop's `readOnly && method !== 'GET'` gate above.
  if (readOnly) {
    logger.info('Skipping write tool download-mail-attachment - read-only mode');
    skippedCount++;
  } else if (isToolEnabled('download-mail-attachment')) {
    try {
      // Simple OneDrive upload (PUT .../content) handles a single file up to
      // 250 MB, which covers every mail attachment (Exchange caps well below
      // this), so we avoid the complexity of chunked upload sessions.
      const MAX_ATTACHMENT_BYTES = 250 * 1024 * 1024;
      server.tool(
        'download-mail-attachment',
        'SUPERSEDED by get-file — prefer that tool. It handles any file, mail attachment or ' +
          'document, and decides for you whether to return the text or a link. Use THIS tool ' +
          'only if you specifically need a staged OneDrive copy.\n\n' +
          'Downloads a mail attachment via OneDrive staging — returns a pre-authed ' +
          'download URL plus metadata, NOT the file bytes themselves. Use this for ' +
          'any attachment large enough that returning base64 through get-mail-attachment ' +
          'would overflow context (realistically anything over ~50KB).\n\n' +
          'NOTE: the returned URL points at the tenant SharePoint host. Some execution ' +
          'environments cannot reach it, so fetching the URL may fail even though staging ' +
          'succeeded — another reason to prefer read-mail-attachment-text for reading.\n\n' +
          'Flow: (1) list-mail-attachments to find the attachmentId; ' +
          '(2) download-mail-attachment to stage the file to OneDrive and get a URL; ' +
          '(3) fetch the URL directly (curl/HTTP) to wherever you need the file.\n\n' +
          'Returns { downloadUrl, name, size, contentType, sha256, expiresInSeconds, stagedDriveItemId }. ' +
          'The downloadUrl is short-lived (~1 hour) and pre-authenticated — no Authorization ' +
          'header needed when fetching it. The bytes are NEVER returned through this tool, so ' +
          'model context is unaffected by attachment size. Only supports fileAttachment; for ' +
          'itemAttachment or referenceAttachment use get-mail-attachment instead. Verify the ' +
          'download with the returned sha256.',
        {
          messageId: z
            .string()
            .min(1)
            .describe('The ID of the mail message that owns the attachment.'),
          attachmentId: z
            .string()
            .min(1)
            .describe('The ID of the attachment to download. Get it from list-mail-attachments.'),
          userId: z
            .string()
            .optional()
            .describe(
              'For shared/other mailboxes: the user id or UPN whose message this is. ' +
                "Omit (or pass 'me') for the signed-in user's own mailbox. The staged " +
                "copy is always written to the signed-in user's OneDrive."
            ),
        },
        {
          title: 'download-mail-attachment',
          readOnlyHint: false,
          destructiveHint: false,
          openWorldHint: true,
        },
        async (params) =>
          withUsageLog('download-mail-attachment', async () => {
            const messageId = (params.messageId ?? '').trim();
            const attachmentId = (params.attachmentId ?? '').trim();
            if (!messageId || !attachmentId) {
              return {
                content: [
                  {
                    type: 'text',
                    text: JSON.stringify({
                      error: 'messageId and attachmentId are both required and must be non-empty.',
                    }),
                  },
                ],
                isError: true,
              };
            }

            const userId = (params.userId ?? '').trim();
            const mailBase =
              userId && userId.toLowerCase() !== 'me'
                ? `/users/${encodeURIComponent(userId)}`
                : '/me';
            const attachmentPath = `${mailBase}/messages/${encodeURIComponent(
              messageId
            )}/attachments/${encodeURIComponent(attachmentId)}`;

            try {
              // 1. Metadata first (no contentBytes — $select keeps it small and
              //    lets us reject non-file attachments before pulling any bytes).
              const metadata = (await graphClient.makeRequest(
                `${attachmentPath}?$select=id,name,contentType,size,isInline`
              )) as {
                '@odata.type'?: string;
                name?: string;
                contentType?: string;
                size?: number;
              };

              const odataType = metadata['@odata.type'];
              if (odataType && odataType !== '#microsoft.graph.fileAttachment') {
                return {
                  content: [
                    {
                      type: 'text',
                      text: JSON.stringify({
                        error: `download-mail-attachment only supports fileAttachment, but this is ${odataType}. Use get-mail-attachment for itemAttachment, or follow the @odata reference for referenceAttachment.`,
                      }),
                    },
                  ],
                  isError: true,
                };
              }

              if (typeof metadata.size === 'number' && metadata.size > MAX_ATTACHMENT_BYTES) {
                return {
                  content: [
                    {
                      type: 'text',
                      text: JSON.stringify({
                        error: `Attachment is ${metadata.size} bytes, which exceeds the ${MAX_ATTACHMENT_BYTES}-byte (250 MB) staging limit.`,
                      }),
                    },
                  ],
                  isError: true,
                };
              }

              // 2. Stream the raw bytes from /$value (server-side only — never
              //    serialized into the tool response).
              const { buffer, contentType: rawContentType } = await graphClient.fetchBinary(
                `${attachmentPath}/$value`
              );
              if (buffer.byteLength > MAX_ATTACHMENT_BYTES) {
                return {
                  content: [
                    {
                      type: 'text',
                      text: JSON.stringify({
                        error: `Attachment is ${buffer.byteLength} bytes, which exceeds the ${MAX_ATTACHMENT_BYTES}-byte (250 MB) staging limit.`,
                      }),
                    },
                  ],
                  isError: true,
                };
              }
              const contentType =
                metadata.contentType || rawContentType || 'application/octet-stream';

              // 3. Build a collision-safe staging filename and upload to a marked
              //    folder under the user's OneDrive root. Using /drive/root (not
              //    special/approot) because the granted scope is Files.ReadWrite,
              //    not Files.ReadWrite.AppFolder — approot is the AppFolder construct
              //    and can 404 under the broad scope. The path auto-creates the folder.
              const rawName =
                (metadata.name && String(metadata.name)) || `attachment-${attachmentId}`;
              const safeName =
                rawName
                  .replace(/[\\/:*?"<>|]/g, '_')
                  .replace(/\s+/g, ' ')
                  .trim() || 'attachment';
              const uniquePrefix = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
              const stagedName = `${uniquePrefix}-${safeName}`;
              const uploadPath = `/me/drive/root:/Apps/TSQ-M365-MCP-staging/${encodeURIComponent(
                stagedName
              )}:/content`;

              const driveItem = (await graphClient.putBinary(uploadPath, buffer, contentType)) as {
                id?: string;
                '@microsoft.graph.downloadUrl'?: string;
              };

              const stagedDriveItemId = driveItem.id;
              let downloadUrl = driveItem['@microsoft.graph.downloadUrl'];

              // The PUT response often omits the downloadUrl; fetch it explicitly
              // from the item if so.
              if (!downloadUrl && stagedDriveItemId) {
                const fetched = (await graphClient.makeRequest(
                  `/me/drive/items/${encodeURIComponent(stagedDriveItemId)}`
                )) as { '@microsoft.graph.downloadUrl'?: string };
                downloadUrl = fetched['@microsoft.graph.downloadUrl'];
              }

              const sha256 = createHash('sha256').update(buffer).digest('hex');

              return {
                content: [
                  {
                    type: 'text',
                    text: JSON.stringify({
                      downloadUrl,
                      name: rawName,
                      size: buffer.byteLength,
                      contentType,
                      sha256,
                      expiresInSeconds: 3600,
                      stagedDriveItemId,
                      note: 'The downloadUrl is pre-authed and short-lived (~1 hour). Fetch it directly via curl/HTTP to download the bytes to your own disk — no Authorization header needed. The bytes are NOT returned through this tool. Verify the download against sha256.',
                    }),
                  },
                ],
              };
            } catch (err) {
              logger.error(`download-mail-attachment failed: ${(err as Error).message}`);
              return {
                content: [
                  {
                    type: 'text',
                    text: JSON.stringify({ error: (err as Error).message }),
                  },
                ],
                isError: true,
              };
            }
          })
      );
      registeredNames.push('download-mail-attachment');
      registeredCount++;
    } catch (error) {
      logger.error(`Failed to register tool download-mail-attachment: ${(error as Error).message}`);
      failedCount++;
    }
  }

  // ---- Change-notification tools --------------------------------------------
  // These wrap the generated /subscriptions endpoints rather than replacing
  // them: the value added here is the notificationUrl plumbing, the clientState
  // secret, the per-user queue, and the long-poll drain.
  //
  // Honest limitation, repeated in every description below: this removes the
  // token cost of polling inside an open session. It cannot wake a closed
  // session — nothing can push into claude.ai.
  registerNotificationTools(server, graphClient, {
    isToolEnabled,
    readOnly,
    push: (name: string) => {
      registeredNames.push(name);
      registeredCount++;
    },
    fail: (name: string, error: Error) => {
      logger.error(`Failed to register tool ${name}: ${error.message}`);
      failedCount++;
    },
    skip: (name: string) => {
      logger.info(`Skipping write tool ${name} - read-only mode`);
      skippedCount++;
    },
  });

  // ---- Document text extraction ---------------------------------------------
  // Read-only: these fetch bytes and return extracted text, writing nothing, so
  // they stay available in read-only mode (unlike download-mail-attachment,
  // which stages a copy to OneDrive and is therefore a write).
  registerFileTextTools(server, graphClient, {
    isToolEnabled,
    push: (name: string) => {
      registeredNames.push(name);
      registeredCount++;
    },
    fail: (name: string, error: Error) => {
      logger.error(`Failed to register tool ${name}: ${error.message}`);
      failedCount++;
    },
  });

  // ---- Unified file handling -------------------------------------------------
  // The meeting tools (list-my-meetings, get-meeting, get-meeting-transcript)
  // are served by TSQ's Note Taker, not Graph. They are registered whether or
  // not Note Taker is configured, so the tool surface is stable; unconfigured,
  // they report themselves as not yet enabled.
  registerMeetingTools(server, graphClient, {
    isToolEnabled,
    push: (name: string) => {
      registeredNames.push(name);
      registeredCount++;
    },
    fail: (name: string, error: Error) => {
      logger.error(`Failed to register tool ${name}: ${error.message}`);
      failedCount++;
    },
  });

  // ---- Delegated inference (Objective 10 pilot, tsq.25) ----------------------
  // Registered only when INFERENCE_ENDPOINT/_DEPLOYMENT are set — absent, the
  // tool does not exist. Deliberately NOT the registered-dark pattern
  // the meeting tools use: a second model in the data path should be
  // invisible until enabled, not discoverable-but-refusing. Pilot list and
  // budget are enforced inside the handler, failing closed.
  registerInferenceTools(server, graphClient, {
    isToolEnabled,
    push: (name: string) => {
      registeredNames.push(name);
      registeredCount++;
    },
    fail: (name: string, error: Error) => {
      logger.error(`Failed to register tool ${name}: ${error.message}`);
      failedCount++;
    },
  });

  // get-file and attach-file supersede the per-delivery tools above. They are
  // registered as writes because both stage bytes into OneDrive or a draft.
  registerFileTools(server, graphClient, {
    isToolEnabled,
    readOnly,
    push: (name: string) => {
      registeredNames.push(name);
      registeredCount++;
    },
    fail: (name: string, error: Error) => {
      logger.error(`Failed to register tool ${name}: ${error.message}`);
      failedCount++;
    },
    skip: (name: string) => {
      logger.info(`Skipping write tool ${name} - read-only mode`);
      skippedCount++;
    },
  });

  // ---- Reply tracking (tsq.27) ------------------------------------------------
  // The state lives on the user's own sent message as an extended property, not
  // on this server: it survives deploys and restarts, any conversation can read
  // it, and TSQ holds nothing. Replaces the need for a live subscription that
  // watch-for-reply has, which only works while its conversation stays open.
  registerReplyTrackingTools(server, graphClient, {
    isToolEnabled,
    readOnly,
    push: (name: string) => {
      registeredNames.push(name);
      registeredCount++;
    },
    fail: (name: string, error: Error) => {
      logger.error(`Failed to register tool ${name}: ${error.message}`);
      failedCount++;
    },
    skip: (name: string) => {
      logger.info(`Skipping write tool ${name} - read-only mode`);
      skippedCount++;
    },
  });

  // ---- Teams replies (tsq.28) --------------------------------------------------
  // Scott, 2026-10-14: six people had answered his agents' Teams chats in a week
  // with nobody noticing. Chat messages cannot carry the reply-tracking marker,
  // so this is computed live from the chats themselves and stores nothing.
  registerChatReplyTools(server, graphClient, {
    isToolEnabled,
    readOnly,
    push: (name: string) => {
      registeredNames.push(name);
      registeredCount++;
    },
    fail: (name: string, error: Error) => {
      logger.error(`Failed to register tool ${name}: ${error.message}`);
      failedCount++;
    },
    skip: (name: string) => {
      logger.info(`Skipping write tool ${name} - read-only mode`);
      skippedCount++;
    },
  });

  // Layer 3 (list-accounts tool) is registered by registerAuthTools in auth-tools.ts.
  // It is the canonical owner of account discovery — no duplicate registration here.

  logger.info(
    `Tool registration complete: ${registeredCount} registered, ${skippedCount} skipped, ${failedCount} failed`
  );
  // Governance audit trail (Phase 1 tiered rollout): record the filter source and
  // the exact, sorted set of tools this connector exposes.
  const filterSource = allowSet
    ? `allowlist (${allowSet.size} tools requested)`
    : enabledToolsPattern
      ? `ENABLED_TOOLS regex: ${enabledToolsPattern}`
      : 'none (all tools)';
  logger.info(`Tool filter source: ${filterSource}`);
  logger.info(
    `Registered tools (${registeredNames.length}): ${[...registeredNames].sort((a, b) => a.localeCompare(b)).join(', ')}`
  );
  return registeredCount;
}

export function buildToolsRegistry(
  readOnly: boolean,
  orgMode: boolean
): Map<string, { tool: (typeof api.endpoints)[0]; config: EndpointConfig | undefined }> {
  const toolsMap = new Map<
    string,
    { tool: (typeof api.endpoints)[0]; config: EndpointConfig | undefined }
  >();

  for (const tool of api.endpoints) {
    const endpointConfig = endpointsData.find((e) => e.toolName === tool.alias);

    if (!orgMode && endpointConfig && !endpointConfig.scopes && endpointConfig.workScopes) {
      continue;
    }

    const method = tool.method.toUpperCase();
    if (readOnly && method !== 'GET') {
      if (!(method === 'POST' && endpointConfig?.readOnly)) {
        continue;
      }
    }

    toolsMap.set(tool.alias, { tool, config: endpointConfig });
  }

  return toolsMap;
}

/**
 * Builds a BM25 index over the tool registry. Name tokens are weighted 3x and llmTip
 * tokens 2x via repetition, so a tool whose name matches the query outranks one that
 * merely mentions the query term in its Microsoft-supplied description.
 */
export function buildDiscoverySearchIndex(
  toolsRegistry: ReturnType<typeof buildToolsRegistry>
): DiscoverySearchIndex {
  // Cap contribution from the `description` and `llmTip` fields so a verbose llmTip
  // (e.g. the KQL search-syntax guide on list-mail-messages, ~300 tokens) doesn't
  // inflate a tool's doc length and crush BM25's length normalization. Names and
  // paths are short and reliable, so they stay uncapped and are repeated to carry
  // the bulk of the ranking signal. Tip excerpt (12 tokens) is enough to capture
  // the first "what this tool does" phrase without swamping the doc.
  const TIP_EXCERPT_TOKENS = 12;
  const DESC_CAP_TOKENS = 40;
  const docs: Array<{ id: string; tokens: string[] }> = [];
  const nameTokens = new Map<string, Set<string>>();
  for (const [name, { tool, config }] of toolsRegistry) {
    const nt = tokenize(name);
    nameTokens.set(name, new Set(nt));
    const pathTokens = tokenize(tool.path);
    const descTokens = tokenize(tool.description).slice(0, DESC_CAP_TOKENS);
    const tipTokens = tokenize(config?.llmTip).slice(0, TIP_EXCERPT_TOKENS);
    const tokens = [
      ...nt,
      ...nt,
      ...nt,
      ...nt,
      ...nt,
      ...pathTokens,
      ...pathTokens,
      ...tipTokens,
      ...descTokens,
    ];
    docs.push({ id: name, tokens });
  }
  return { bm25: buildBM25Index(docs), nameTokens };
}

/**
 * BM25 + a "name precision" bonus: reward tools whose names contain a high fraction
 * of the query tokens (and consist mostly of query-matching tokens). This counteracts
 * cases where a tool with a longer or more off-topic description outranks a tool
 * whose name directly matches — a common problem because many endpoint descriptions
 * are the wrong Graph prose pasted in.
 */
export function scoreDiscoveryQuery(
  query: string,
  index: DiscoverySearchIndex
): Array<{ id: string; score: number }> {
  const queryTokenSet = new Set(tokenize(query));
  if (queryTokenSet.size === 0) return [];
  const ranked = scoreQuery(query, index.bm25);
  const NAME_BONUS_WEIGHT = 2;
  for (const r of ranked) {
    const nt = index.nameTokens.get(r.id);
    if (!nt || nt.size === 0) continue;
    let matchedIdf = 0;
    let matchedCount = 0;
    for (const qt of queryTokenSet) {
      if (nt.has(qt)) {
        matchedCount++;
        matchedIdf += index.bm25.idf.get(qt) ?? 0;
      }
    }
    if (matchedCount === 0) continue;
    const precision = matchedCount / nt.size;
    r.score += precision * matchedIdf * NAME_BONUS_WEIGHT;
  }
  ranked.sort((a, b) => b.score - a.score);
  return ranked;
}

export function registerDiscoveryTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false,
  orgMode: boolean = false,
  authManager?: AuthManager,
  _multiAccount: boolean = false
): void {
  const toolsRegistry = buildToolsRegistry(readOnly, orgMode);
  const searchIndex = buildDiscoverySearchIndex(toolsRegistry);
  logger.info(`Discovery mode: ${toolsRegistry.size} tools available in registry`);

  const categoryNames = Object.keys(TOOL_CATEGORIES).join(', ');

  const toResultEntry = (name: string) => {
    const entry = toolsRegistry.get(name);
    if (!entry) return null;
    const { tool, config } = entry;
    return {
      name,
      method: tool.method.toUpperCase(),
      path: tool.path,
      description: tool.description || `${tool.method.toUpperCase()} ${tool.path}`,
      ...(config?.llmTip ? { llmTip: config.llmTip } : {}),
    };
  };

  server.tool(
    'search-tools',
    `Search through ${toolsRegistry.size} Microsoft Graph API tools. Ranks results by BM25 over tool name, llmTip, description, and path (tokenized on hyphens, camelCase, and whitespace). After picking a tool, call get-tool-schema to see its parameters, then execute-tool to invoke it.`,
    {
      query: z
        .string()
        .describe(
          'Natural-language query. Tokenized and BM25-ranked. E.g. "send email", "create calendar event", "list unread messages".'
        )
        .optional(),
      category: z.string().describe(`Optional pre-filter by category: ${categoryNames}`).optional(),
      limit: z.number().describe('Maximum results (default: 10, max: 50)').optional(),
    },
    {
      title: 'search-tools',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ query, category, limit = 10 }) => {
      const maxLimit = Math.min(Math.max(limit, 1), 50);
      const categoryDef = category ? TOOL_CATEGORIES[category] : undefined;
      const categoryFilter = (name: string) => !categoryDef || categoryDef.pattern.test(name);

      let orderedNames: string[];
      if (query && query.trim().length > 0) {
        const ranked = scoreDiscoveryQuery(query, searchIndex);
        orderedNames = ranked.map((r) => r.id).filter(categoryFilter);
      } else {
        orderedNames = [...toolsRegistry.keys()].filter(categoryFilter);
      }

      const tools = orderedNames.slice(0, maxLimit).map(toResultEntry).filter(Boolean);

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                found: tools.length,
                total: toolsRegistry.size,
                tools,
                tip: 'Call get-tool-schema(tool_name) to see parameters before invoking execute-tool.',
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  server.tool(
    'get-tool-schema',
    'Returns the full parameter schema (name, placement, required, JSON Schema) for a tool discovered via search-tools. Call this before execute-tool so you know what parameters to pass and what enum values are valid.',
    {
      tool_name: z.string().describe('Exact tool name from search-tools (e.g. "send-mail")'),
    },
    {
      title: 'get-tool-schema',
      readOnlyHint: true,
      openWorldHint: false,
    },
    async ({ tool_name }) => {
      const entry = toolsRegistry.get(tool_name);
      if (!entry) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error: `Tool not found: ${tool_name}`,
                tip: 'Use search-tools to find available tools.',
              }),
            },
          ],
          isError: true,
        };
      }
      const schema = describeToolSchema(entry.tool, entry.config?.llmTip);
      return {
        content: [{ type: 'text', text: JSON.stringify(schema, null, 2) }],
      };
    }
  );

  server.tool(
    'execute-tool',
    'Execute a Microsoft Graph API tool by name. Workflow: search-tools → get-tool-schema → execute-tool. Call get-tool-schema first for any tool you have not seen before — passing the wrong shape to parameters will fail validation or return a Graph 400. For list endpoints, prefer modest $top plus $select.',
    {
      tool_name: z.string().describe('Name of the tool to execute (e.g., "list-mail-messages")'),
      parameters: z
        .record(z.any())
        .describe(
          'Parameters shaped per get-tool-schema. Path/query/header params go at the top level; request bodies go under "body".'
        )
        .optional(),
    },
    {
      title: 'execute-tool',
      readOnlyHint: false,
      destructiveHint: true,
      openWorldHint: true,
    },
    async ({ tool_name, parameters = {} }) => {
      const toolData = toolsRegistry.get(tool_name);
      if (!toolData) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error: `Tool not found: ${tool_name}`,
                tip: 'Use search-tools to find available tools.',
              }),
            },
          ],
          isError: true,
        };
      }

      return executeGraphTool(toolData.tool, toolData.config, graphClient, parameters, authManager);
    }
  );

  // Layer 3 (list-accounts) is registered by registerAuthTools — no duplicate here.
}
