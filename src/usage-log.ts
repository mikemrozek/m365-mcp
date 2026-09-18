import winston from 'winston';
import path from 'path';
import fs from 'fs';
import os from 'os';
import { getRequestActor } from './request-context.js';

/**
 * Structured per-invocation usage log: "what tool was used, by whom, and did it
 * succeed." Emitted as one JSON object per line, tagged `type: "m365-usage"`, so
 * periodic reporting is a group-by, not a regex scrape of the chatty server log.
 * Drives the connector packaging decision (which tools users actually use — see
 * config/allowlists/README.md).
 *
 * Two sinks:
 *   - stderr (always, when enabled) — so container platforms capture records
 *     durably. On Azure Container Apps these land in Log Analytics
 *     (`ContainerAppConsoleLogs_CL`, queryable by KQL: `where Log_s has
 *     "m365-usage"`). stderr is used (not stdout) because in stdio MCP mode
 *     stdout carries the JSON-RPC protocol and must not be polluted.
 *   - a rotating local file (`usage.log`) — for local/dev tailing.
 *
 * Disable entirely with USAGE_LOG=off. Override the file location with
 * MS365_MCP_USAGE_LOG_DIR (defaults to the same dir as the main server log).
 */

const enabled = (process.env.USAGE_LOG ?? '').toLowerCase() !== 'off';

const logsDir =
  process.env.MS365_MCP_USAGE_LOG_DIR ||
  process.env.MS365_MCP_LOG_DIR ||
  path.join(os.homedir(), '.ms-365-mcp-server', 'logs');

if (enabled && !fs.existsSync(logsDir)) {
  fs.mkdirSync(logsDir, { recursive: true, mode: 0o700 });
}

const transports: winston.transport[] = [];
if (enabled) {
  transports.push(
    new winston.transports.File({
      filename: path.join(logsDir, 'usage.log'),
      maxsize: 10 * 1024 * 1024, // 10 MB
      maxFiles: 5,
      tailable: true,
    })
  );
  // Route info-level records to stderr (not stdout) so they are captured by the
  // container platform without corrupting stdio-mode MCP traffic.
  transports.push(new winston.transports.Console({ stderrLevels: ['error', 'warn', 'info'] }));
}

const usageLogger = winston.createLogger({
  level: 'info',
  format: winston.format.combine(winston.format.timestamp(), winston.format.json()),
  transports,
  silent: !enabled,
});

export type ToolOutcome = 'success' | 'error';

/**
 * Why a call failed, in two fields that cannot carry message content.
 *
 * Until 2026-09-18 the record said only that a call failed. Asked "what are the
 * errors" against a week holding 29 of them, the only way to answer was to
 * reproduce each shape by hand — the server's own Graph error text goes to a
 * file inside the container and dies with it.
 *
 * Graph's error text is not the fix: it echoes the request back, so a failed
 * search would put the search terms into Log Analytics, and a failed send the
 * recipients. Only a numeric status and a short symbolic code are taken, and
 * the code is reduced to a bare token before it is written. Same discipline as
 * SEC-2026-001 — diagnose from shape, never from payload.
 */
export interface FailureDetail {
  /** HTTP status, when the failure carried one. */
  status?: number;
  /** Graph's own `error.code` ('BadRequest', 'SearchWithFilter'), or our refusal slug. */
  code?: string;
}

/** `Microsoft Graph API error: 400 Bad Request - {...}` and the shapes around it. */
function statusFrom(text: string): number | undefined {
  const fromPrefix = /error:\s*(\d{3})\b/i.exec(text);
  if (fromPrefix) return Number(fromPrefix[1]);
  const fromField = /["']?status(?:Code)?["']?\s*[:=]\s*(\d{3})\b/i.exec(text);
  return fromField ? Number(fromField[1]) : undefined;
}

/**
 * Graph nests its code as `{"error":{"code":"BadRequest"}}`; our own tools refuse
 * with `{"error":"invalid_search"}`. Anything else yields nothing rather than a
 * guess, and whatever is found is filtered to a token so a message can never
 * ride out through this field.
 */
function codeFrom(text: string): string | undefined {
  const start = text.indexOf('{');
  if (start === -1) return undefined;
  try {
    const parsed = JSON.parse(text.slice(start)) as { error?: string | { code?: string } };
    const raw = typeof parsed.error === 'string' ? parsed.error : parsed.error?.code;
    if (typeof raw !== 'string') return undefined;
    return raw.replace(/[^A-Za-z0-9_.:-]/g, '').slice(0, 64) || undefined;
  } catch {
    // Not JSON, or a brace that only looked like the start of it.
    return undefined;
  }
}

/** Reads a thrown error or a tool result; returns {} when it can tell nothing. */
export function describeFailure(source: unknown): FailureDetail {
  let text: string | undefined;
  if (source instanceof Error) text = source.message;
  else if (typeof source === 'string') text = source;
  else {
    const content = (source as { content?: { text?: unknown }[] } | null)?.content;
    const first = Array.isArray(content) ? content[0]?.text : undefined;
    if (typeof first === 'string') text = first;
  }
  if (!text) return {};

  const detail: FailureDetail = {};
  const status = statusFrom(text);
  if (status !== undefined) detail.status = status;
  const code = codeFrom(text);
  if (code !== undefined) detail.code = code;
  return detail;
}

/**
 * Record a single tool invocation. Actor is pulled from the request context
 * (HTTP/OAuth mode); in stdio/CLI mode there is a single local user, so actor
 * fields are simply absent.
 */
export function logToolUsage(tool: string, outcome: ToolOutcome, failure?: FailureDetail): void {
  if (!enabled) return;
  const actor = getRequestActor();
  usageLogger.info('tool-usage', {
    type: 'm365-usage', // stable marker for log queries
    tool,
    outcome,
    // Absent on success, and absent when the failure carried neither.
    status: failure?.status,
    code: failure?.code,
    // Stable per-user key for reporting; upn is the human-readable label.
    oid: actor?.oid,
    upn: actor?.upn,
    tid: actor?.tid,
  });
}

/**
 * Wrap a tool handler so its outcome is recorded on every exit path. Use for
 * custom tools that do not route through the shared Graph executor.
 */
export async function withUsageLog<T extends { isError?: boolean }>(
  tool: string,
  fn: () => Promise<T>
): Promise<T> {
  try {
    const result = await fn();
    logToolUsage(
      tool,
      result.isError ? 'error' : 'success',
      result.isError ? describeFailure(result) : undefined
    );
    return result;
  } catch (error) {
    logToolUsage(tool, 'error', describeFailure(error));
    throw error;
  }
}

export default usageLogger;
