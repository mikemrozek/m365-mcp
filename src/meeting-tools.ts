import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type GraphClient from './graph-client.js';
import logger from './logger.js';
import { withUsageLog } from './usage-log.js';
import { getRequestTokens } from './request-context.js';
import { isMeetingThreadId, parseTeamsUrl, threadIdFromTeamsUrl } from './lib/teams-url-parser.js';
import {
  isIdentityVerificationError,
  verifyCallerIdentity,
  type VerifiedIdentity,
} from './lib/verified-identity.js';
import {
  NoteTakerClient,
  NoteTakerError,
  noteTakerConfigFromEnv,
  transcriptsEnabledFromEnv,
  type NoteTakerCode,
  type NoteTakerMeetingRow,
} from './lib/notetaker-client.js';

/**
 * Meeting intelligence — served by TSQ's Note Taker API v1, not Microsoft Graph.
 *
 *   list-my-meetings        GET  /api/v1/meetings
 *   get-meeting             GET  /api/v1/meetings/{id}, or POST /api/v1/meetings/lookup first
 *   get-meeting-transcript  GET  /api/v1/meetings/{id}/transcript (gated off in v1)
 *
 * A Teams conversation id names a SERIES, not a meeting (one id covers 197 of
 * Note Taker's records), so every call after discovery uses Note Taker's own
 * meeting id. Note Taker decides what the caller may see, against the object id
 * we verified via Graph `/me`. No tool takes a user or object id as an argument:
 * who is asking comes from the session only.
 *
 * Nothing here logs meeting content. Note Taker's request ids travel back in the
 * tool result (`requestId`, and `_meta.noteTakerRequestIds`) for reconciliation.
 */

export interface RegisterHooks {
  isToolEnabled: (name: string) => boolean;
  push: (name: string) => void;
  fail: (name: string, error: Error) => void;
}

export interface MeetingToolsDeps {
  /** Test seam. Defaults to a client built from MS365_MCP_NOTETAKER_* env vars. */
  noteTaker?: NoteTakerClient | null;
  /** Test seam. Defaults to MS365_MCP_NOTETAKER_TRANSCRIPTS_ENABLED. */
  transcriptsEnabled?: boolean;
}

export const MEETING_TOOL_NAMES = [
  'list-my-meetings',
  'get-meeting',
  'get-meeting-transcript',
] as const;

/** Same ceilings as get-file's `as: 'text'`: a long transcript can crowd out the conversation. */
export const DEFAULT_MAX_CHARS = 50_000;
export const MAX_CHARS_LIMIT = 200_000;

/** How far either side of a calendar event's start to look for the occurrence. */
const EVENT_WINDOW_MS = 24 * 60 * 60 * 1000;

const REPORT_TO = 'Mike Mrozek';

interface ToolResult {
  [key: string]: unknown;
  content: { type: 'text'; text: string }[];
  isError?: boolean;
}

function result(payload: Record<string, unknown>, isError: boolean, requestIds: string[]) {
  const r: ToolResult = { content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }] };
  if (isError) r.isError = true;
  // Machine-readable copy for usage logging; never content.
  if (requestIds.length) r._meta = { noteTakerRequestIds: requestIds };
  return r;
}

// ---------------------------------------------------------------------------
// Resolving a calendar event / Teams link / thread id to a conversation id
// ---------------------------------------------------------------------------

interface MeetingRef {
  threadId?: string;
  meetingUrl?: string;
  eventId?: string;
}

export interface ResolvedConversation {
  threadId: string;
  via: 'threadId' | 'url' | 'graph-lookup' | 'event';
  /** Calendar event start in UTC, when the reference was an event. */
  eventStart?: string;
}

/**
 * Turn whatever the caller has into the meeting chat thread id — Note Taker's
 * `conversation_id`. Short /meet/ links and calendar events need Graph; the
 * other forms resolve locally. Exported for tests.
 */
export async function resolveThreadId(
  graphClient: Pick<GraphClient, 'makeRequest'>,
  ref: MeetingRef
): Promise<ResolvedConversation> {
  if (ref.threadId) {
    const id = ref.threadId.trim();
    if (!isMeetingThreadId(id)) {
      throw new Error(
        `"${id}" is not a meeting thread id. Expected the form 19:meeting_…@thread.v2 — it is onlineMeeting.chatInfo.threadId, or the id of the meeting's chat.`
      );
    }
    return { threadId: id, via: 'threadId' };
  }

  if (ref.meetingUrl) {
    const normalized = parseTeamsUrl(ref.meetingUrl.trim());
    const local = threadIdFromTeamsUrl(normalized);
    if (local) return { threadId: local, via: 'url' };

    const escaped = normalized.replace(/'/g, "''");
    const found = (await graphClient.makeRequest(
      `/me/onlineMeetings?$filter=joinWebUrl eq '${escaped}'&$select=id,subject,chatInfo`
    )) as { value?: { chatInfo?: { threadId?: string } }[] };
    const threadId = found?.value?.[0]?.chatInfo?.threadId;
    if (!threadId) {
      throw new Error(
        'Could not resolve that Teams link to a meeting. Short /meet/ links only resolve for meetings you can see; try the full join link or the calendar event instead.'
      );
    }
    return { threadId, via: 'graph-lookup' };
  }

  if (ref.eventId) {
    const event = (await graphClient.makeRequest(
      `/me/events/${encodeURIComponent(ref.eventId.trim())}?$select=subject,start,isOnlineMeeting,onlineMeeting`
    )) as {
      subject?: string;
      start?: { dateTime?: string; timeZone?: string };
      onlineMeeting?: { joinUrl?: string };
    };
    const joinUrl = event?.onlineMeeting?.joinUrl;
    if (!joinUrl) {
      throw new Error(
        `Calendar event${event?.subject ? ` "${event.subject}"` : ''} has no Teams meeting link, so there is no meeting to look up.`
      );
    }
    const resolved = await resolveThreadId(graphClient, { meetingUrl: joinUrl });
    return { threadId: resolved.threadId, via: 'event', eventStart: graphStartToUtc(event.start) };
  }

  throw new Error('Identify the meeting with one of meetingId, eventId, meetingUrl or threadId.');
}

/**
 * Graph returns event times as {dateTime, timeZone} with no offset, in UTC
 * unless a Prefer: outlook.timezone header was sent (this server sends none).
 * If some other zone comes back it is still read as UTC: the error is at most
 * 14 hours, inside the ±1 day window this feeds.
 */
function graphStartToUtc(start?: { dateTime?: string; timeZone?: string }): string | undefined {
  const raw = start?.dateTime?.trim();
  if (!raw) return undefined;
  const hasOffset = /(Z|[+-]\d{2}:?\d{2})$/i.test(raw);
  const ms = Date.parse(hasOffset ? raw : `${raw.replace(/(\.\d{3})\d+$/, '$1')}Z`);
  return Number.isFinite(ms) ? utcIso(ms) : undefined;
}

function utcIso(ms: number): string {
  return new Date(ms).toISOString().replace(/\.000Z$/, 'Z');
}

/** Accepts any ISO 8601 date or date-time; returns UTC with a Z, as Note Taker wants. */
export function toUtcIso(value: string, field: string): string {
  const ms = Date.parse(value.trim());
  if (!Number.isFinite(ms)) {
    throw new Error(`${field} "${value}" is not an ISO 8601 date or date-time.`);
  }
  return utcIso(ms);
}

// ---------------------------------------------------------------------------
// Refusals → plain language
// ---------------------------------------------------------------------------

type Fault =
  | 'connector_configuration'
  | 'access'
  | 'not_found'
  | 'switched_off'
  | 'rate_limit'
  | 'temporary';

const CONFIG_FAULT = (what: string) =>
  `The meeting-notes connector is misconfigured: ${what}. This is not something you did and not something you can fix — please report it to ${REPORT_TO}, with the request id if one is shown.`;

/** One outcome per code. Exported for tests. */
export const REFUSALS: Record<NoteTakerCode, { fault: Fault; message: string }> = {
  invalid_token: {
    fault: 'connector_configuration',
    message: CONFIG_FAULT("Note Taker did not accept the connector's own credential"),
  },
  token_expired: {
    fault: 'connector_configuration',
    message: CONFIG_FAULT("Note Taker says the connector's own credential has expired"),
  },
  caller_not_allowed: {
    fault: 'connector_configuration',
    message: CONFIG_FAULT('Note Taker does not recognise the connector as an allowed caller'),
  },
  app_token_required: {
    fault: 'connector_configuration',
    message: CONFIG_FAULT('Note Taker expected an application credential and got something else'),
  },
  missing_role: {
    fault: 'connector_configuration',
    message: CONFIG_FAULT('the connector has not been granted the Note Taker permission it needs'),
  },
  user_identity_missing: {
    fault: 'connector_configuration',
    message: CONFIG_FAULT('the connector did not tell Note Taker who is asking'),
  },
  user_identity_invalid: {
    fault: 'connector_configuration',
    message: CONFIG_FAULT('Note Taker could not read the identity the connector sent'),
  },
  token_unavailable: {
    fault: 'connector_configuration',
    message: CONFIG_FAULT('the connector could not obtain its credential for Note Taker'),
  },
  unexpected_response: {
    fault: 'connector_configuration',
    message: CONFIG_FAULT('Note Taker answered in a form the connector does not understand'),
  },
  not_configured: {
    fault: 'switched_off',
    message:
      'Meeting notes are not enabled on this connector yet. The Note Taker integration is built and waiting for its service to go live.',
  },
  user_not_enabled: {
    fault: 'access',
    message:
      'Your account is not enabled for Note Taker meeting notes yet. Access is being rolled out person by person; ask to be added if you need it. Nothing is wrong with the connector.',
  },
  meeting_not_found: {
    fault: 'not_found',
    message:
      'Note Taker has no meeting by that id that you attended. It deliberately gives the same answer whether the meeting does not exist or you were not in it. Use list-my-meetings to see the meetings it has for you.',
  },
  tenant_transcript_access_disabled: {
    fault: 'switched_off',
    message:
      'Townsquare has turned off programmatic access to meeting transcripts organisation-wide, and Note Taker honours that setting. Nothing can be retrieved until it is turned back on; there is no workaround.',
  },
  api_disabled: {
    fault: 'switched_off',
    message:
      "Note Taker's owner has paused access for this connector. Nothing you did caused it; try again later.",
  },
  transcripts_disabled: {
    fault: 'switched_off',
    message:
      "Note Taker's owner has switched transcript access off for now. Meeting summaries may still be available through get-meeting.",
  },
  rate_limited_user: {
    fault: 'rate_limit',
    message: 'You have made a lot of meeting-notes requests in a short time. Wait, then try again.',
  },
  rate_limited_connector: {
    fault: 'rate_limit',
    message:
      'Note Taker is limiting the connector as a whole right now (everyone shares this limit). Wait, then try again.',
  },
  daily_limit_user: {
    fault: 'rate_limit',
    message: 'You have reached your daily limit for meeting-notes requests. It resets tomorrow.',
  },
  daily_limit_connector: {
    fault: 'rate_limit',
    message:
      'The connector has reached its daily limit with Note Taker for everyone. It resets tomorrow.',
  },
  rate_limited: {
    fault: 'rate_limit',
    message: 'Note Taker is rate-limiting requests. Wait, then try again.',
  },
  upstream_unavailable: {
    fault: 'temporary',
    message:
      'Note Taker is up, but a service it depends on is not answering. Try again in a few minutes.',
  },
  timeout: {
    fault: 'temporary',
    message: 'Note Taker did not answer in time. Try again in a few minutes.',
  },
  unreachable: {
    fault: 'temporary',
    message:
      'Note Taker could not be reached. Try again in a few minutes; if it persists, report it to ' +
      REPORT_TO +
      '.',
  },
};

/** Exported for tests. Never includes Note Taker's own message text in the result. */
export function explainRefusal(e: NoteTakerError): Record<string, unknown> {
  const outcome = REFUSALS[e.code] ?? REFUSALS.unexpected_response;
  const payload: Record<string, unknown> = {
    error: e.code,
    status: e.status,
    fault: outcome.fault,
    message: outcome.message,
  };
  if (outcome.fault === 'connector_configuration') payload.reportTo = REPORT_TO;
  if (e.retryAfterSeconds !== undefined) {
    payload.retryAfterSeconds = e.retryAfterSeconds;
    payload.message = `${outcome.message.replace(/ Wait, then try again\.$/, '')} Try again in about ${e.retryAfterSeconds} seconds.`;
  }
  if (e.serverCode) payload.noteTakerCode = e.serverCode;
  if (e.requestId) payload.requestId = e.requestId;
  return payload;
}

function identityFailure(error: { reason: string }): Record<string, unknown> {
  if (error.reason === 'graph_unavailable') {
    return {
      error: 'identity_unverified',
      reason: error.reason,
      fault: 'temporary',
      message:
        'Could not confirm who is asking (Microsoft Graph did not answer), so nothing was sent to Note Taker. Try again shortly.',
    };
  }
  if (error.reason === 'tenant_mismatch' || error.reason === 'config_unavailable') {
    return {
      error: 'identity_unverified',
      reason: error.reason,
      fault: 'connector_configuration',
      reportTo: REPORT_TO,
      message: CONFIG_FAULT(
        'it could not confirm your sign-in belongs to the configured organisation'
      ),
    };
  }
  return {
    error: 'identity_unverified',
    reason: error.reason,
    fault: 'access',
    message:
      'Could not confirm who is asking, so the request was not sent to Note Taker. Reconnecting the connector usually fixes this.',
  };
}

// ---------------------------------------------------------------------------
// Registration
// ---------------------------------------------------------------------------

export function registerMeetingTools(
  server: McpServer,
  graphClient: GraphClient,
  hooks: RegisterHooks,
  deps: MeetingToolsDeps = {}
): void {
  const { isToolEnabled, push, fail } = hooks;

  const noteTaker =
    deps.noteTaker !== undefined
      ? deps.noteTaker
      : (() => {
          const config = noteTakerConfigFromEnv();
          if (!config) {
            logger.info(
              'Meeting tools registered dark: MS365_MCP_NOTETAKER_API_URL/_AUDIENCE not set'
            );
            return null;
          }
          return new NoteTakerClient(config);
        })();
  const transcriptsEnabled = deps.transcriptsEnabled ?? transcriptsEnabledFromEnv();

  /**
   * The common path: configured? caller credential? verified identity? Then
   * run. Every Note Taker refusal becomes a structured, plain-language result.
   */
  const withCaller =
    (
      run: (
        nt: NoteTakerClient,
        caller: VerifiedIdentity,
        requestIds: string[]
      ) => Promise<{ payload: Record<string, unknown>; isError?: boolean }>
    ) =>
    async (): Promise<ToolResult> => {
      const requestIds: string[] = [];
      if (!noteTaker) {
        return result({ error: 'not_configured', ...REFUSALS.not_configured }, true, requestIds);
      }
      const ctx = getRequestTokens();
      if (!ctx?.accessToken) {
        return result(
          { error: 'no_caller', message: 'No caller credential on this request.' },
          true,
          requestIds
        );
      }
      let caller: VerifiedIdentity;
      try {
        caller = await verifyCallerIdentity(ctx.accessToken);
      } catch (error) {
        if (isIdentityVerificationError(error)) {
          return result(identityFailure(error), true, requestIds);
        }
        throw error;
      }
      try {
        const out = await run(noteTaker, caller, requestIds);
        return result(out.payload, !!out.isError, requestIds);
      } catch (error) {
        if (error instanceof NoteTakerError) {
          if (error.requestId && !requestIds.includes(error.requestId)) {
            requestIds.push(error.requestId);
          }
          return result(explainRefusal(error), true, requestIds);
        }
        throw error;
      }
    };

  const track = (requestIds: string[], id?: string) => {
    if (id) requestIds.push(id);
    return id;
  };

  const annotations = (title: string) => ({
    title,
    readOnlyHint: true,
    destructiveHint: false,
    openWorldHint: true,
  });

  const register = (name: string, fn: () => void) => {
    if (!isToolEnabled(name)) return;
    try {
      fn();
      push(name);
    } catch (error) {
      fail(name, error as Error);
    }
  };

  const isoField = (what: string) =>
    z.string().optional().describe(`${what}, ISO 8601 (e.g. 2026-10-01 or 2026-10-01T14:00:00Z)`);

  // ---- list-my-meetings ------------------------------------------------------
  register('list-my-meetings', () =>
    server.tool(
      'list-my-meetings',
      "List my meetings with notes — Teams meetings the user attended that TSQ's Note Taker recorded, with whether a meeting summary, decisions and action items (recap) are available. " +
        'Use to find a past meeting by date, subject or organizer before reading its notes with get-meeting. ' +
        'Returns rows: Note Taker meeting id, subject, start/end (UTC), duration, organizer, participant count, conversation_id, has_summary, has_transcript, attended. ' +
        'Filters: startAfter/startBefore (ISO date or date-time), subjectContains, organizerEmail; page with limit (max 100) and offset. ' +
        'Only meetings the user attended are returned — Note Taker enforces that, not this tool.',
      {
        startAfter: isoField('Only meetings starting after this time'),
        startBefore: isoField('Only meetings starting before this time'),
        subjectContains: z.string().optional().describe('Text the meeting subject contains'),
        organizerEmail: z.string().optional().describe("Organizer's email address"),
        limit: z
          .number()
          .int()
          .min(1)
          .max(100)
          .optional()
          .describe('Rows per page, 1-100 (Note Taker default if omitted)'),
        offset: z.number().int().min(0).optional().describe('Rows to skip, for paging'),
      },
      annotations('list-my-meetings'),
      async ({ startAfter, startBefore, subjectContains, organizerEmail, limit, offset }) =>
        withUsageLog(
          'list-my-meetings',
          withCaller(async (nt, caller, requestIds) => {
            let filters;
            try {
              filters = {
                start_after: startAfter ? toUtcIso(startAfter, 'startAfter') : undefined,
                start_before: startBefore ? toUtcIso(startBefore, 'startBefore') : undefined,
                subject_contains: subjectContains?.trim() || undefined,
                organizer_email: organizerEmail?.trim() || undefined,
                limit,
                offset,
              };
            } catch (error) {
              return {
                payload: { error: 'invalid_input', message: (error as Error).message },
                isError: true,
              };
            }
            const list = await nt.listMeetings(caller, filters);
            const requestId = track(requestIds, list.requestId);
            const mayHaveMore = limit !== undefined && list.meetings.length === limit;
            return {
              payload: {
                count: list.meetings.length,
                meetings: list.meetings,
                ...list.paging,
                requestId,
                ...(list.meetings.length === 0
                  ? {
                      note: 'Note Taker has no meetings for you matching these filters. It only knows meetings it recorded and that you attended.',
                    }
                  : mayHaveMore
                    ? {
                        note: `There may be more; call again with offset ${(offset ?? 0) + list.meetings.length}.`,
                      }
                    : {}),
              },
            };
          })
        )
    )
  );

  // ---- get-meeting -----------------------------------------------------------
  register('get-meeting', () =>
    server.tool(
      'get-meeting',
      'Get meeting notes: the summary, decisions, action items, key numbers and participants of a Teams meeting the user attended, from TSQ\'s Note Taker — "what did we decide", "recap of Tuesday\'s meeting". ' +
        'Identify the meeting by meetingId (Note Taker id from list-my-meetings — preferred) OR by eventId (calendar event id), meetingUrl (any Teams link) or threadId (19:meeting_…@thread.v2). ' +
        'A Teams link or thread id names a whole recurring series, so give startAfter/startBefore to narrow it; a calendar event is narrowed to ±1 day of its start automatically. ' +
        'If several occurrences match, the tool returns them (with ids) to choose from — call again with meetingId. ' +
        'Channel meetings and some short /meet/ links may not be findable this way; use list-my-meetings instead.',
      {
        meetingId: z
          .string()
          .optional()
          .describe('Note Taker meeting id, from list-my-meetings or an earlier get-meeting'),
        eventId: z
          .string()
          .optional()
          .describe('Calendar event id; the server follows its Teams link and start time'),
        meetingUrl: z
          .string()
          .optional()
          .describe('Any Teams meeting URL (join, /meet/ short link, or recap)'),
        threadId: z.string().optional().describe('Meeting chat thread id: 19:meeting_…@thread.v2'),
        startAfter: isoField('With eventId/meetingUrl/threadId: occurrence starts after this'),
        startBefore: isoField('With eventId/meetingUrl/threadId: occurrence starts before this'),
      },
      annotations('get-meeting'),
      async ({ meetingId, eventId, meetingUrl, threadId, startAfter, startBefore }) =>
        withUsageLog(
          'get-meeting',
          withCaller(async (nt, caller, requestIds) => {
            if (meetingId?.trim()) {
              const detail = await nt.getMeeting(caller, meetingId.trim());
              const requestId = track(requestIds, detail.requestId);
              return { payload: { meeting: detail.meeting, requestId } };
            }

            let resolved: ResolvedConversation;
            let window: { start_after?: string; start_before?: string };
            try {
              resolved = await resolveThreadId(graphClient, { threadId, meetingUrl, eventId });
              const start = resolved.eventStart ? Date.parse(resolved.eventStart) : undefined;
              window = {
                start_after: startAfter
                  ? toUtcIso(startAfter, 'startAfter')
                  : start !== undefined
                    ? utcIso(start - EVENT_WINDOW_MS)
                    : undefined,
                start_before: startBefore
                  ? toUtcIso(startBefore, 'startBefore')
                  : start !== undefined
                    ? utcIso(start + EVENT_WINDOW_MS)
                    : undefined,
              };
            } catch (error) {
              return {
                payload: { error: 'unresolved_meeting', message: (error as Error).message },
                isError: true,
              };
            }

            const found = await nt.lookupMeetings(caller, {
              conversation_id: resolved.threadId,
              ...window,
            });
            track(requestIds, found.requestId);
            const context = {
              resolvedVia: resolved.via,
              conversationId: resolved.threadId,
              ...(window.start_after || window.start_before ? { window } : {}),
            };

            if (found.meetings.length === 0) {
              return {
                payload: {
                  found: 0,
                  ...context,
                  message:
                    'Note Taker has no occurrence of this meeting that you attended' +
                    (window.start_after || window.start_before ? ' in that time window' : '') +
                    '. It only knows meetings it recorded. Channel meetings and meetings joined through short /meet/ links may have no conversation id on its side — try list-my-meetings with a date range or subject instead.',
                  requestIds: [...requestIds],
                },
              };
            }

            if (found.meetings.length > 1) {
              return {
                payload: {
                  found: found.meetings.length,
                  ...context,
                  message:
                    'Several occurrences of this meeting match. Pick one and call get-meeting again with its id as meetingId' +
                    (window.start_after || window.start_before
                      ? '.'
                      : ', or narrow with startAfter/startBefore.'),
                  meetings: found.meetings,
                  requestIds: [...requestIds],
                },
              };
            }

            const row = found.meetings[0] as NoteTakerMeetingRow;
            const detail = await nt.getMeeting(caller, String(row.id));
            track(requestIds, detail.requestId);
            return {
              payload: { meeting: detail.meeting, ...context, requestIds: [...requestIds] },
            };
          })
        )
    )
  );

  // ---- get-meeting-transcript --------------------------------------------------
  register('get-meeting-transcript', () =>
    server.tool(
      'get-meeting-transcript',
      "Get the transcript of a meeting — the full text of what was said — from TSQ's Note Taker, by Note Taker meetingId (from list-my-meetings or get-meeting). " +
        'NOT YET ENABLED: until transcript access is switched on it returns not_enabled; use get-meeting for the summary, decisions and action items instead. ' +
        'Paged: pass the next-page value from the previous response as page. ' +
        `Each page is cut at maxChars (default ${DEFAULT_MAX_CHARS}) and the response says so.`,
      {
        meetingId: z
          .string()
          .describe('Note Taker meeting id, from list-my-meetings or get-meeting'),
        page: z
          .string()
          .optional()
          .describe('Page token or number from the previous response; omit for the first page'),
        maxChars: z
          .number()
          .int()
          .min(500)
          .max(MAX_CHARS_LIMIT)
          .optional()
          .describe(
            `Cap on transcript characters returned. Default ${DEFAULT_MAX_CHARS}, ceiling ${MAX_CHARS_LIMIT}.`
          ),
      },
      annotations('get-meeting-transcript'),
      async ({ meetingId, page, maxChars }) =>
        withUsageLog('get-meeting-transcript', async () => {
          if (!transcriptsEnabled) {
            return result(
              {
                error: 'not_enabled',
                fault: 'switched_off',
                message:
                  'Meeting transcripts are not enabled yet; they follow the meeting summaries in a later phase. Use get-meeting for the summary, decisions and action items.',
              },
              true,
              []
            );
          }
          return withCaller(async (nt, caller, requestIds) => {
            const t = await nt.getTranscript(caller, meetingId.trim(), page?.trim() || undefined);
            const requestId = track(requestIds, t.requestId);
            const cap = maxChars ?? DEFAULT_MAX_CHARS;
            const total = t.text.length;
            const truncated = total > cap;
            return {
              payload: {
                meetingId: meetingId.trim(),
                ...t.meta,
                text: truncated ? t.text.slice(0, cap) : t.text,
                totalChars: total,
                truncated,
                requestId,
                ...(truncated
                  ? {
                      note: `Page cut at ${cap} of ${total} characters. Call again with a higher maxChars (up to ${MAX_CHARS_LIMIT}) if the answer may be later in this page.`,
                    }
                  : {}),
              },
            };
          })();
        })
    )
  );
}
