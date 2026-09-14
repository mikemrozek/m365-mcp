import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type GraphClient from './graph-client.js';
import logger from './logger.js';
import { withUsageLog } from './usage-log.js';
import { getRequestTokens } from './request-context.js';
import { isMeetingThreadId, parseTeamsUrl, threadIdFromTeamsUrl } from './lib/teams-url-parser.js';
import { isIdentityVerificationError, verifyCallerIdentity } from './lib/verified-identity.js';
import {
  NoteTakerClient,
  NoteTakerError,
  noteTakerConfigFromEnv,
  type NoteTakerTranscript,
} from './lib/notetaker-client.js';

/**
 * Meeting intelligence — served by TSQ's Note Taker rather than Microsoft Graph.
 *
 * One tool, one meeting per call, no search over transcript content. That
 * narrowness is deliberate and mirrors Note Taker's own endpoint: it is a
 * lookup, not a dragnet. The caller must have been a participant; Note Taker
 * checks that itself against the identity we verified via Graph `/me`.
 */

export interface RegisterHooks {
  isToolEnabled: (name: string) => boolean;
  push: (name: string) => void;
  fail: (name: string, error: Error) => void;
}

export interface MeetingToolsDeps {
  /** Test seam. Defaults to a client built from MS365_MCP_NOTETAKER_* env vars. */
  noteTaker?: NoteTakerClient | null;
}

/** Same ceilings as get-file's `as: 'text'`, for the same reason: a long transcript can crowd out the conversation. */
export const DEFAULT_MAX_CHARS = 50_000;
export const MAX_CHARS_LIMIT = 200_000;

interface ToolResult {
  [key: string]: unknown;
  content: { type: 'text'; text: string }[];
  isError?: boolean;
}

const ok = (payload: unknown): ToolResult => ({
  content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }],
});
const err = (payload: Record<string, unknown>): ToolResult => ({
  content: [{ type: 'text', text: JSON.stringify(payload, null, 2) }],
  isError: true,
});

interface MeetingRef {
  threadId?: string;
  meetingUrl?: string;
  eventId?: string;
}

/**
 * Turn whatever the caller has into the meeting chat thread id Note Taker keys on.
 * Exported for tests. Short /meet/ links and calendar events need Graph; the
 * other forms resolve locally.
 */
export async function resolveThreadId(
  graphClient: Pick<GraphClient, 'makeRequest'>,
  ref: MeetingRef
): Promise<{ threadId: string; via: 'threadId' | 'url' | 'graph-lookup' | 'event' }> {
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
      `/me/events/${encodeURIComponent(ref.eventId.trim())}?$select=subject,isOnlineMeeting,onlineMeeting`
    )) as { subject?: string; onlineMeeting?: { joinUrl?: string } };
    const joinUrl = event?.onlineMeeting?.joinUrl;
    if (!joinUrl) {
      throw new Error(
        `Calendar event${event?.subject ? ` "${event.subject}"` : ''} has no Teams meeting link, so there is no meeting to look up.`
      );
    }
    const resolved = await resolveThreadId(graphClient, { meetingUrl: joinUrl });
    return { threadId: resolved.threadId, via: 'event' };
  }

  throw new Error('Identify the meeting with one of threadId, meetingUrl or eventId.');
}

/** What to tell the caller for each refusal. Operational faults say so, so nobody retries them differently. */
function explainRefusal(e: NoteTakerError): Record<string, unknown> {
  const base = { error: e.refusal, message: e.message };
  switch (e.refusal) {
    case 'not_participant':
      return {
        ...base,
        note: 'Note Taker only serves a transcript to someone who was in the meeting. If the user was invited under a different address, that is why.',
      };
    case 'transcripts_disabled':
      return {
        ...base,
        note: 'This is the organisation-wide Teams setting for programmatic transcript access, switched off. Nothing can be retrieved until it is re-enabled; there is no workaround through this tool.',
      };
    case 'meeting_not_found':
      return {
        ...base,
        note: 'Note Taker only knows meetings it attended. Check the meeting had Note Taker present, or that the id is the meeting chat thread rather than a channel or 1:1 chat.',
      };
    case 'rate_limited':
      return { ...base, retryAfterSeconds: e.retryAfterSeconds };
    case 'disabled':
      return { ...base, note: 'The endpoint owner has paused it. Nothing to retry.' };
    default:
      return {
        ...base,
        note: 'Server-side problem, not something to retry with different inputs. Report it to IT.',
      };
  }
}

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
              'get-meeting-transcript registered dark: MS365_MCP_NOTETAKER_API_URL/_AUDIENCE not set'
            );
            return null;
          }
          return new NoteTakerClient(config);
        })();

  const name = 'get-meeting-transcript';
  if (!isToolEnabled(name)) return;

  try {
    server.tool(
      name,
      "Get the transcript of a Teams meeting the user attended, served by TSQ's Note Taker (not Microsoft Graph). " +
        'Identify the meeting by ONE of: threadId (the meeting chat thread id, 19:meeting_…@thread.v2 — onlineMeeting.chatInfo.threadId, or the id of the meeting chat), ' +
        'meetingUrl (any Teams link: join, short /meet/, or recap), or eventId (a calendar event id from list-calendar-events / get-calendar-view; the server follows its Teams link). ' +
        'Only participants can retrieve it and Note Taker verifies that. Returns the transcript text with subject, organizer, join URL and times, ' +
        'or a structured reason when no transcript exists (not recorded, too old, ad-hoc call). ' +
        'Long transcripts are cut at maxChars (default 50,000) and the response says so — ask for more with a higher maxChars rather than assuming you saw it all.',
      {
        threadId: z.string().optional().describe('Meeting chat thread id: 19:meeting_…@thread.v2'),
        meetingUrl: z
          .string()
          .optional()
          .describe('Any Teams meeting URL (join, /meet/ short link, or recap)'),
        eventId: z
          .string()
          .optional()
          .describe('Calendar event id; the server resolves its Teams link'),
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
      {
        title: name,
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: true,
      },
      async ({ threadId, meetingUrl, eventId, maxChars }) =>
        withUsageLog(name, async () => {
          if (!noteTaker) {
            return err({
              error: 'not_configured',
              message:
                'Meeting transcripts are not enabled on this server yet. The Note Taker integration is built but waiting on its API endpoint; ask IT for status.',
            });
          }

          const ctx = getRequestTokens();
          if (!ctx?.accessToken) {
            return err({ error: 'no_caller', message: 'No caller credential on this request.' });
          }

          let resolved: Awaited<ReturnType<typeof resolveThreadId>>;
          try {
            resolved = await resolveThreadId(graphClient, { threadId, meetingUrl, eventId });
          } catch (error) {
            return err({ error: 'unresolved_meeting', message: (error as Error).message });
          }

          let caller;
          try {
            caller = await verifyCallerIdentity(ctx.accessToken);
          } catch (error) {
            if (isIdentityVerificationError(error)) {
              return err({
                error: 'identity_unverified',
                reason: error.reason,
                message:
                  error.reason === 'graph_unavailable'
                    ? 'Could not confirm who is asking (Microsoft Graph did not answer). Try again shortly.'
                    : 'Could not confirm who is asking, so the request was not sent to Note Taker.',
              });
            }
            throw error;
          }

          let result: NoteTakerTranscript;
          try {
            result = await noteTaker.getTranscript(resolved.threadId, caller);
          } catch (error) {
            if (error instanceof NoteTakerError) return err(explainRefusal(error));
            throw error;
          }

          const cap = maxChars ?? DEFAULT_MAX_CHARS;
          if (result.transcript.available) {
            const total = result.transcript.content.length;
            const truncated = total > cap;
            return ok({
              ...result,
              resolvedVia: resolved.via,
              transcript: {
                ...result.transcript,
                content: truncated
                  ? result.transcript.content.slice(0, cap)
                  : result.transcript.content,
              },
              totalChars: total,
              truncated,
              ...(truncated
                ? {
                    note: `Transcript cut at ${cap} of ${total} characters. Anything after that point was not returned; call again with a higher maxChars (up to ${MAX_CHARS_LIMIT}) if the answer may be later in the meeting.`,
                  }
                : {}),
            });
          }
          return ok({ ...result, resolvedVia: resolved.via });
        })
    );
    push(name);
  } catch (error) {
    fail(name, error as Error);
  }
}
