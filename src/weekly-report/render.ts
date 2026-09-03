/**
 * Renders the report model as an email.
 *
 * Inline styles only, no external CSS or images — mail clients strip them, and
 * a report nobody can read is worse than no report because it looks like it
 * worked. Anomalies come first and totals second, deliberately: the reader is
 * assumed to be skimming on a phone, and the reason this exists is to catch the
 * thing nobody thought to look for.
 */

import type { Report, ToolLine, UserLine } from './analyze.js';

/**
 * Segoe UI 11pt throughout, tables included.
 *
 * The family and size are repeated on the table and on every cell rather than left to
 * inherit: Outlook and most webmail reset fonts inside table elements, so a single
 * declaration on the wrapper produces a document that looks right in a browser and
 * wrong in the client the report is actually read in.
 */
// The family is single-quoted, so every style attribute below must be double-quoted.
// A single-quoted style attribute holding a single-quoted font family terminates at
// the first inner quote and drops the rest of the declaration — silently, and only in
// the mail client rather than in any preview. Hence double quotes on every attribute.
const FONT_STACK = "font-family:'Segoe UI',Segoe,Tahoma,Arial,sans-serif;font-size:11pt";
const FONT = `${FONT_STACK};line-height:1.5;color:#202020`;
const TABLE = `border-collapse:collapse;margin:8px 0;${FONT_STACK}`;
const CELL = `border:1px solid #d0d0d0;padding:6px 10px;text-align:left;${FONT_STACK}`;
const HEADING = `${FONT_STACK};font-weight:600;margin:16px 0 4px`;

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function day(date: Date): string {
  return date.toLocaleDateString('en-GB', { day: 'numeric', month: 'long' });
}

/** A signed change with an arrow, or an em dash when there is nothing to compare. */
function delta(now: number, before: number): string {
  if (before === 0 && now === 0) return '—';
  if (before === 0) return `new (${now})`;
  const diff = now - before;
  const sign = diff > 0 ? '+' : '';
  return `${before} → ${now} (${sign}${diff})`;
}

function rate(errors: number, calls: number): string {
  if (calls === 0) return '—';
  return `${((errors / calls) * 100).toFixed(1)}%`;
}

function table(headers: string[], rows: string[][]): string {
  const head = headers.map((h) => `<th style="${CELL};background:#f4f4f4'>${h}</th>`).join('');
  const body = rows
    .map((r) => `<tr>${r.map((c) => `<td style="${CELL}">${c}</td>`).join('')}</tr>`)
    .join('');
  return `<table style="${TABLE}"><tr>${head}</tr>${body}</table>`;
}

function userRows(users: UserLine[]): string[][] {
  return users.map((u) => [
    escapeHtml(u.upn),
    String(u.calls),
    String(u.tools),
    delta(u.calls, u.priorCalls),
  ]);
}

function toolRows(tools: ToolLine[]): string[][] {
  return tools.map((t) => [
    escapeHtml(t.tool),
    String(t.calls),
    delta(t.calls, t.priorCalls),
    t.errors ? String(t.errors) : '—',
  ]);
}

/** "31 August" -> "August 31", the way the subject line reads. */
function monthDay(date: Date): string {
  return date.toLocaleDateString('en-US', { month: 'long', day: 'numeric' });
}

export function renderSubject(report: Report): string {
  return `M365 Connector: ${monthDay(report.window.from)} to ${monthDay(report.window.to)}, ${report.headline}`;
}

export function renderHtml(report: Report, options: { secretExpiry?: Date } = {}): string {
  const { current, prior } = report;
  const parts: string[] = [];

  parts.push(`<div style="${FONT}">`);
  parts.push(
    `<p style="color:#666;${FONT_STACK}">Week of ${day(report.window.from)} to ${day(report.window.to)}, compared with ${day(report.priorWindow.from)} to ${day(report.priorWindow.to)}.</p>`
  );

  if (report.anomalies.length) {
    parts.push(`<h3 style="${HEADING}">Worth a look</h3><ul>`);
    for (const line of report.anomalies) parts.push(`<li>${escapeHtml(line)}</li>`);
    parts.push('</ul>');
  } else {
    parts.push(
      `<p><strong>Nothing unusual this week.</strong> No new or lapsed users, no capability failing at a rate worth chasing, and no sharp change in volume.</p>`
    );
  }

  parts.push(`<h3 style="${HEADING}">The week in totals</h3>`);
  parts.push(
    table(
      ['', 'This week', 'Last week'],
      [
        ['Calls', String(current.calls), String(prior.calls)],
        ['Errors', String(current.errors), String(prior.errors)],
        ['Error rate', rate(current.errors, current.calls), rate(prior.errors, prior.calls)],
        ['People', String(current.users), String(prior.users)],
        ['Capabilities used', String(current.tools), String(prior.tools)],
      ]
    )
  );

  parts.push(`<h3 style="${HEADING}">Who used it</h3>`);
  parts.push(table(['Person', 'Calls', 'Capabilities', 'Change'], userRows(report.users)));

  parts.push(`<h3 style="${HEADING}">Most used</h3>`);
  parts.push(table(['Capability', 'Calls', 'Change', 'Errors'], toolRows(report.topTools)));

  if (report.errorTools.length) {
    parts.push(`<h3 style="${HEADING}">Where the errors were</h3>`);
    parts.push(table(['Capability', 'Calls', 'Change', 'Errors'], toolRows(report.errorTools)));
  }

  const fileTotal = report.fileSplit.replacement + report.fileSplit.superseded;
  parts.push(`<h3 style="${HEADING}">File handling</h3>`);
  parts.push(
    `<p>get-file: <strong>${report.fileSplit.replacement}</strong> calls. The five capabilities it replaced: <strong>${report.fileSplit.superseded}</strong>. ` +
      (fileTotal
        ? `That is ${Math.round((report.fileSplit.replacement / fileTotal) * 100)}% on the replacement.`
        : 'No file activity this week.') +
      ` A retirement decision rides on this line, which is why it is reported every week rather than checked when someone remembers.</p>`
  );

  if (report.neverUsed.length) {
    parts.push(`<h3 style="${HEADING}">Never used (${report.neverUsed.length})</h3>`);
    parts.push(
      `<p style="color:#666;${FONT_STACK}">Advertised to everyone, called by nobody in this period: ${report.neverUsed.map(escapeHtml).join(', ')}</p>`
    );
  }

  if (options.secretExpiry) {
    const days = Math.round((options.secretExpiry.getTime() - Date.now()) / 86400000);
    parts.push(
      `<p style="color:#666;${FONT_STACK}">Connector credential expires ${options.secretExpiry.toISOString().slice(0, 10)} — ${days} days from now.</p>`
    );
  }

  parts.push('</div>');
  return parts.join('');
}

/** One line for the container log, so a run leaves evidence even if mail fails. */
export function renderLogLine(report: Report): string {
  return (
    `calls=${report.current.calls} (prior ${report.prior.calls}) ` +
    `errors=${report.current.errors} users=${report.current.users} ` +
    `tools=${report.current.tools} anomalies=${report.anomalies.length}`
  );
}
