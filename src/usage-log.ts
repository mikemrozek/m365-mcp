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
 * Record a single tool invocation. Actor is pulled from the request context
 * (HTTP/OAuth mode); in stdio/CLI mode there is a single local user, so actor
 * fields are simply absent.
 */
export function logToolUsage(tool: string, outcome: ToolOutcome): void {
  if (!enabled) return;
  const actor = getRequestActor();
  usageLogger.info('tool-usage', {
    type: 'm365-usage', // stable marker for log queries
    tool,
    outcome,
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
    logToolUsage(tool, result.isError ? 'error' : 'success');
    return result;
  } catch (error) {
    logToolUsage(tool, 'error');
    throw error;
  }
}

export default usageLogger;
