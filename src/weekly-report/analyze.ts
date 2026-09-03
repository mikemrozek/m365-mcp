/**
 * Turns raw usage records into the model the weekly report renders.
 *
 * Deliberately pure: no Azure, no Graph, no clock of its own. Everything the
 * report says is derived here from rows plus an explicit `now`, so the whole
 * report is testable from fabricated data — which matters, because the thing
 * this report exists to catch is the week where nobody thought to look.
 *
 * Why it exists at all (2026-09-01): every figure in the 1 September brief was
 * gathered by hand. One of those hand-run queries reversed a recommendation the
 * previous brief had made, hours before that recommendation would have shipped.
 * Totals alone would not have caught it — the reversal was visible only as a
 * trend against the prior week. So every number here carries its prior-week
 * counterpart, and anomalies are computed rather than left for a reader to spot.
 */

export interface UsageRow {
  /** ISO timestamp of the call. */
  ts: string;
  tool: string;
  /** 'success' or anything else; the log writes 'error' for failures. */
  outcome: string;
  upn: string;
}

export interface Totals {
  calls: number;
  errors: number;
  users: number;
  tools: number;
}

export interface UserLine {
  upn: string;
  calls: number;
  tools: number;
  priorCalls: number;
}

export interface ToolLine {
  tool: string;
  calls: number;
  priorCalls: number;
  errors: number;
}

export interface Report {
  window: { from: Date; to: Date };
  priorWindow: { from: Date; to: Date };
  current: Totals;
  prior: Totals;
  users: UserLine[];
  topTools: ToolLine[];
  errorTools: ToolLine[];
  /** The file-handling split, called out because a retirement decision rides on it. */
  fileSplit: {
    replacement: number;
    superseded: number;
    supersededDetail: ToolLine[];
    /** False once the superseded tools are no longer advertised, which changes the advice. */
    supersededStillAdvertised: boolean;
  };
  /** Allowlisted capabilities with no calls in either window. */
  neverUsed: string[];
  /** Plain sentences, most significant first. Empty when the week was unremarkable. */
  anomalies: string[];
  /** A few words naming the single most notable thing, for the subject line. */
  headline: string;
}

/** The capability `get-file` replaced, and the five it replaced. */
const FILE_REPLACEMENT = 'get-file';
const FILE_SUPERSEDED = [
  'get-mail-attachment',
  'download-mail-attachment',
  'read-mail-attachment-text',
  'read-onedrive-file-text',
  'download-onedrive-file-content',
];

const DAY_MS = 24 * 60 * 60 * 1000;

function isError(row: UsageRow): boolean {
  return row.outcome !== 'success';
}

function totals(rows: UsageRow[]): Totals {
  return {
    calls: rows.length,
    errors: rows.filter(isError).length,
    users: new Set(rows.map((r) => r.upn)).size,
    tools: new Set(rows.map((r) => r.tool)).size,
  };
}

function countBy<T>(rows: UsageRow[], key: (r: UsageRow) => T): Map<T, UsageRow[]> {
  const out = new Map<T, UsageRow[]>();
  for (const row of rows) {
    const k = key(row);
    const bucket = out.get(k);
    if (bucket) bucket.push(row);
    else out.set(k, [row]);
  }
  return out;
}

/** "laura.mirarchi@x.com" -> "Laura", for a subject line that reads like a sentence. */
function friendlyName(upn: string): string {
  const local = upn.split('@')[0] ?? upn;
  const first = local.split(/[._-]/)[0] ?? local;
  return first ? first.charAt(0).toUpperCase() + first.slice(1) : upn;
}

/** A readable multiple: 3.4x, or 'from nothing' when the base was zero. */
function ratio(now: number, before: number): string {
  if (before === 0) return 'from nothing';
  const r = now / before;
  return r >= 1 ? `${r.toFixed(1)}x more` : `${(1 / r).toFixed(1)}x less`;
}

/**
 * Anomalies are the point of the report. Totals tell a reader the system is
 * alive; anomalies tell them where to look. Each rule below exists because a
 * real week would have been misread without it:
 *
 *   - a user appearing or vanishing was invisible behind a flat headcount
 *     (7 users in both weeks either side of 1 September — but not the same 7)
 *   - a capability carrying nearly every error in the system (294 of 311)
 *   - a trend reversing inside one week (the file-tool split)
 */
function findAnomalies(
  current: UsageRow[],
  prior: UsageRow[],
  users: UserLine[],
  tools: ToolLine[],
  fileSplit: Report['fileSplit']
): string[] {
  const out: string[] = [];

  for (const u of users.filter((u) => u.priorCalls === 0 && u.calls > 0)) {
    out.push(
      `${u.upn} is new this week — ${u.calls} calls across ${u.tools} capabilities, none last week.`
    );
  }
  for (const u of users.filter((u) => u.calls === 0 && u.priorCalls > 0)) {
    out.push(`${u.upn} stopped — ${u.priorCalls} calls last week, none this week.`);
  }
  for (const u of users.filter((u) => u.calls > 0 && u.priorCalls >= 10)) {
    const r = u.calls / u.priorCalls;
    if (r >= 3 || r <= 1 / 3) {
      out.push(
        `${u.upn} used it ${ratio(u.calls, u.priorCalls)} than last week (${u.priorCalls} → ${u.calls}).`
      );
    }
  }

  for (const t of tools.filter((t) => t.errors >= 5 && t.errors / t.calls >= 0.2)) {
    const pct = Math.round((t.errors / t.calls) * 100);
    out.push(`${t.tool} failed ${t.errors} of ${t.calls} calls (${pct}%).`);
  }

  const cur = totals(current);
  const pre = totals(prior);
  if (pre.calls >= 50 && (cur.calls / pre.calls >= 2 || cur.calls / pre.calls <= 0.5)) {
    out.push(
      `Overall volume is ${ratio(cur.calls, pre.calls)} than last week (${pre.calls} → ${cur.calls}).`
    );
  }

  // Only advise on the file split while the superseded tools are actually advertised.
  // They were retired in tsq.18, and a report that keeps recommending against a
  // retirement already carried out is worse than one that says nothing: it reads as
  // current advice, and the first edition Scott sees should not argue with a decision
  // taken the night before.
  const fileTotal = fileSplit.replacement + fileSplit.superseded;
  if (fileTotal >= 10 && fileSplit.superseded > fileSplit.replacement) {
    const pct = Math.round((fileSplit.replacement / fileTotal) * 100);
    if (fileSplit.supersededStillAdvertised) {
      out.push(
        `File handling: get-file is the minority path at ${pct}% (${fileSplit.replacement} calls against ${fileSplit.superseded} on the five it replaced). Retiring them would remove the route most callers use.`
      );
    } else if (fileSplit.superseded > 0) {
      out.push(
        `File handling: ${fileSplit.superseded} calls still went to capabilities retired part-way through this window, against ${fileSplit.replacement} on get-file. Expect that to reach zero next week; if it does not, the retirement did not take.`
      );
    }
  }

  return out;
}

/**
 * The single most notable thing, in a few words, for the subject line.
 *
 * Ordered by what would make a reader open the mail first: something failing beats
 * someone leaving, which beats someone arriving, which beats a change in volume. A
 * quiet week says so rather than manufacturing drama.
 */
function buildHeadline(
  users: UserLine[],
  tools: ToolLine[],
  current: Totals,
  prior: Totals
): string {
  const failing = tools
    .filter((t) => t.errors >= 5 && t.errors / t.calls >= 0.2)
    .sort((a, b) => b.errors - a.errors)[0];
  if (failing) {
    return `${failing.tool} failing ${Math.round((failing.errors / failing.calls) * 100)}% of the time`;
  }

  const lapsed = users
    .filter((u) => u.calls === 0 && u.priorCalls >= 10)
    .sort((a, b) => b.priorCalls - a.priorCalls)[0];
  if (lapsed && lapsed.priorCalls >= 100) {
    return `${friendlyName(lapsed.upn)} stopped after ${lapsed.priorCalls} calls`;
  }

  const arrived = users
    .filter((u) => u.priorCalls === 0 && u.calls >= 10)
    .sort((a, b) => b.calls - a.calls)[0];
  if (arrived) {
    return `${friendlyName(arrived.upn)} is new, at ${arrived.calls} calls`;
  }

  const grown = users
    .filter((u) => u.priorCalls >= 10 && u.calls / u.priorCalls >= 3)
    .sort((a, b) => b.calls / b.priorCalls - a.calls / a.priorCalls)[0];
  if (grown) {
    return `${friendlyName(grown.upn)} up ${(grown.calls / grown.priorCalls).toFixed(1)}x`;
  }

  if (lapsed) return `${friendlyName(lapsed.upn)} stopped`;

  if (prior.calls >= 50) {
    const ratio = current.calls / prior.calls;
    if (ratio <= 0.5) return `volume down ${(1 / ratio).toFixed(1)}x`;
    if (ratio >= 2) return `volume up ${ratio.toFixed(1)}x`;
  }

  return 'nothing unusual';
}

export function buildReport(
  rows: UsageRow[],
  options: { now: Date; windowDays?: number; allowlist?: string[] }
): Report {
  const days = options.windowDays ?? 7;
  const to = options.now;
  const from = new Date(to.getTime() - days * DAY_MS);
  const priorFrom = new Date(from.getTime() - days * DAY_MS);

  const at = (r: UsageRow) => new Date(r.ts).getTime();
  const current = rows.filter((r) => at(r) >= from.getTime() && at(r) < to.getTime());
  const prior = rows.filter((r) => at(r) >= priorFrom.getTime() && at(r) < from.getTime());

  const priorByUser = countBy(prior, (r) => r.upn);
  const currentByUser = countBy(current, (r) => r.upn);
  const userNames = new Set([...currentByUser.keys(), ...priorByUser.keys()]);
  const users: UserLine[] = [...userNames]
    .map((upn) => ({
      upn,
      calls: currentByUser.get(upn)?.length ?? 0,
      tools: new Set((currentByUser.get(upn) ?? []).map((r) => r.tool)).size,
      priorCalls: priorByUser.get(upn)?.length ?? 0,
    }))
    .sort((a, b) => b.calls - a.calls || b.priorCalls - a.priorCalls);

  const priorByTool = countBy(prior, (r) => r.tool);
  const currentByTool = countBy(current, (r) => r.tool);
  const toolLine = (tool: string): ToolLine => {
    const calls = currentByTool.get(tool) ?? [];
    return {
      tool,
      calls: calls.length,
      priorCalls: priorByTool.get(tool)?.length ?? 0,
      errors: calls.filter(isError).length,
    };
  };
  const allTools = [...new Set([...currentByTool.keys(), ...priorByTool.keys()])].map(toolLine);

  const topTools = [...allTools].sort((a, b) => b.calls - a.calls).slice(0, 10);
  const errorTools = allTools.filter((t) => t.errors > 0).sort((a, b) => b.errors - a.errors);

  const advertised = options.allowlist ? new Set(options.allowlist) : undefined;
  const fileSplit = {
    replacement: toolLine(FILE_REPLACEMENT).calls,
    superseded: FILE_SUPERSEDED.reduce((n, t) => n + toolLine(t).calls, 0),
    supersededDetail: FILE_SUPERSEDED.map(toolLine),
    // With no allowlist to check against, assume they are still advertised — the
    // cautious reading, since it only ever adds a line rather than removing one.
    supersededStillAdvertised: advertised ? FILE_SUPERSEDED.some((t) => advertised.has(t)) : true,
  };

  const seen = new Set(rows.map((r) => r.tool));
  const neverUsed = (options.allowlist ?? []).filter((t) => !seen.has(t)).sort();

  const anomalies = findAnomalies(current, prior, users, allTools, fileSplit);

  return {
    window: { from, to },
    priorWindow: { from: priorFrom, to: from },
    current: totals(current),
    prior: totals(prior),
    users,
    topTools,
    errorTools,
    fileSplit,
    neverUsed,
    anomalies,
    headline: buildHeadline(users, allTools, totals(current), totals(prior)),
  };
}
