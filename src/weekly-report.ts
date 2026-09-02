/**
 * Entry point for the weekly usage report.
 *
 * Runs as a scheduled Azure Container Apps job against the same image as the
 * server, invoked with `node dist/weekly-report.js`. It shares nothing with the
 * running server but the image: no HTTP listener, no MCP session, no user
 * token. Keeping it in this repository rather than in a script on someone's
 * machine is deliberate — the report exists because hand-run queries are the
 * thing that stops happening, and a report that lives on one laptop has the
 * same failure mode it was built to fix.
 *
 * Configuration, all by environment variable so the job definition is the only
 * place the values live:
 *
 *   REPORT_WORKSPACE_ID    Log Analytics workspace holding the usage records
 *   REPORT_TENANT_ID       tenant for the token exchange
 *   REPORT_CLIENT_ID       the reporting app registration (Mail.Send, one mailbox)
 *   REPORT_UAMI_CLIENT_ID  the job's user-assigned managed identity
 *   REPORT_SENDER          mailbox to send as; must match the Exchange access policy
 *   REPORT_RECIPIENTS      comma-separated
 *   REPORT_WINDOW_DAYS     optional, default 7
 *   TOOL_ALLOWLIST         optional; enables the never-used list. Same value the server gets
 *   REPORT_SECRET_EXPIRY   optional ISO date, printed as a countdown in the footer
 *
 * `--dry-run` prints the HTML and sends nothing. Use it to review a change to
 * the report before a live run mails it to anyone.
 */

import logger from './logger.js';
import { loadToolAllowlist } from './tool-allowlist.js';
import { buildReport } from './weekly-report/analyze.js';
import { renderHtml, renderLogLine, renderSubject } from './weekly-report/render.js';
import {
  fetchUsageRows,
  getGraphTokenViaFederation,
  sendReportMail,
} from './weekly-report/sources.js';

function required(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) throw new Error(`${name} is not set. The job definition is incomplete.`);
  return value;
}

function optionalAllowlist(): string[] | undefined {
  const raw = process.env.TOOL_ALLOWLIST?.trim();
  if (!raw) return undefined;
  try {
    return loadToolAllowlist(raw).names;
  } catch (error) {
    // A malformed allowlist must not cost the whole report — it only powers the
    // never-used section, which is the least important part of it.
    logger.warn(`[WEEKLY REPORT] Ignoring TOOL_ALLOWLIST: ${(error as Error).message}`);
    return undefined;
  }
}

export async function runWeeklyReport(argv: string[] = process.argv.slice(2)): Promise<number> {
  const dryRun = argv.includes('--dry-run');
  const windowDays = Number(process.env.REPORT_WINDOW_DAYS ?? '7');
  const managedIdentityClientId = required('REPORT_UAMI_CLIENT_ID');

  // Two windows of data, so every figure can carry its prior-week counterpart.
  const rows = await fetchUsageRows({
    workspaceId: required('REPORT_WORKSPACE_ID'),
    days: windowDays * 2,
    managedIdentityClientId,
  });

  const expiry = process.env.REPORT_SECRET_EXPIRY?.trim();
  const report = buildReport(rows, {
    now: new Date(),
    windowDays,
    allowlist: optionalAllowlist(),
  });

  const html = renderHtml(report, { secretExpiry: expiry ? new Date(expiry) : undefined });
  const subject = renderSubject(report);
  logger.info(`[WEEKLY REPORT] ${renderLogLine(report)}`);

  if (dryRun) {
    process.stdout.write(`${subject}\n\n${html}\n`);
    return 0;
  }

  const token = await getGraphTokenViaFederation({
    tenantId: required('REPORT_TENANT_ID'),
    clientId: required('REPORT_CLIENT_ID'),
    managedIdentityClientId,
  });

  await sendReportMail({
    token,
    senderUpn: required('REPORT_SENDER'),
    recipients: required('REPORT_RECIPIENTS')
      .split(',')
      .map((r) => r.trim())
      .filter(Boolean),
    subject,
    html,
  });
  return 0;
}

// Only run when invoked directly, so the module stays importable from tests.
if (process.argv[1]?.endsWith('weekly-report.js')) {
  runWeeklyReport()
    .then((code) => process.exit(code))
    .catch((error) => {
      // A failed report must be loud in the container log — its whole purpose is
      // to be the thing that notices, so it failing silently is the worst case.
      logger.error(`[WEEKLY REPORT] Failed: ${(error as Error).message}`);
      process.exit(1);
    });
}
