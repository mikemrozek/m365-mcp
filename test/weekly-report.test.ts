import { describe, expect, it } from 'vitest';
import { buildReport, type UsageRow } from '../src/weekly-report/analyze.js';
import { renderHtml, renderLogLine, renderSubject } from '../src/weekly-report/render.js';

const NOW = new Date('2026-09-08T12:00:00Z');
const DAY = 24 * 60 * 60 * 1000;

/** A row `daysAgo` before NOW, so tests read as "last week" / "the week before". */
function row(daysAgo: number, over: Partial<UsageRow> = {}): UsageRow {
  return {
    ts: new Date(NOW.getTime() - daysAgo * DAY).toISOString(),
    tool: 'list-mail-messages',
    outcome: 'success',
    upn: 'someone@townsquaremedia.com',
    ...over,
  };
}

const build = (rows: UsageRow[], allowlist?: string[]) =>
  buildReport(rows, { now: NOW, windowDays: 7, allowlist });

describe('buildReport windows', () => {
  it('splits current from prior and ignores anything older', () => {
    const report = build([row(1), row(3), row(9), row(20)]);
    expect(report.current.calls).toBe(2);
    expect(report.prior.calls).toBe(1);
  });

  it('counts an outcome other than success as an error', () => {
    const report = build([row(1), row(2, { outcome: 'error' })]);
    expect(report.current.calls).toBe(2);
    expect(report.current.errors).toBe(1);
  });
});

describe('anomalies', () => {
  it('names a user who appeared this week', () => {
    const report = build([row(9, { upn: 'old@x.com' }), row(1, { upn: 'new@x.com' })]);
    expect(report.anomalies.join(' ')).toContain('new@x.com is new this week');
  });

  it('names a user who stopped', () => {
    const report = build([row(9, { upn: 'gone@x.com' }), row(1, { upn: 'still@x.com' })]);
    expect(report.anomalies.join(' ')).toContain('gone@x.com stopped');
  });

  it('flags a large swing for an established user', () => {
    const rows = [
      ...Array.from({ length: 30 }, () => row(9, { upn: 'busy@x.com' })),
      ...Array.from({ length: 3 }, () => row(1, { upn: 'busy@x.com' })),
    ];
    expect(build(rows).anomalies.join(' ')).toContain('10.0x less');
  });

  it('ignores a swing on a tiny base, because noise is not news', () => {
    const rows = [
      row(9, { upn: 'quiet@x.com' }),
      ...Array.from({ length: 9 }, () => row(1, { upn: 'quiet@x.com' })),
    ];
    expect(build(rows).anomalies.join(' ')).not.toContain('quiet@x.com used it');
  });

  it('flags a capability failing at a rate worth chasing', () => {
    const rows = Array.from({ length: 10 }, (_, i) =>
      row(1, { tool: 'move-mail-message', outcome: i < 6 ? 'error' : 'success' })
    );
    expect(build(rows).anomalies.join(' ')).toContain(
      'move-mail-message failed 6 of 10 calls (60%)'
    );
  });

  it('stops advising against a retirement once the tools are no longer advertised', () => {
    const rows = [
      ...Array.from({ length: 11 }, () => row(1, { tool: 'get-file' })),
      ...Array.from({ length: 18 }, () => row(1, { tool: 'download-mail-attachment' })),
    ];
    // Allowlist without the superseded tools: they shipped as retired in tsq.18.
    const after = buildReport(rows, {
      now: NOW,
      windowDays: 7,
      allowlist: ['get-file', 'put-file', 'attach-file'],
    });
    const text = after.anomalies.join(' ');
    expect(text).not.toContain('Retiring them would remove');
    expect(text).toContain('retired part-way through this window');
    expect(after.fileSplit.supersededStillAdvertised).toBe(false);
  });

  it('flags the file-handling reversal that a totals-only report would hide', () => {
    const rows = [
      ...Array.from({ length: 11 }, () => row(1, { tool: 'get-file' })),
      ...Array.from({ length: 8 }, () => row(1, { tool: 'download-mail-attachment' })),
      ...Array.from({ length: 6 }, () => row(2, { tool: 'get-mail-attachment' })),
      ...Array.from({ length: 4 }, () => row(2, { tool: 'read-mail-attachment-text' })),
    ];
    const report = build(rows);
    expect(report.fileSplit.replacement).toBe(11);
    expect(report.fileSplit.superseded).toBe(18);
    expect(report.anomalies.join(' ')).toContain('get-file is the minority path at 38%');
  });

  it('says so plainly when the week was unremarkable', () => {
    const rows = [
      ...Array.from({ length: 5 }, () => row(9)),
      ...Array.from({ length: 5 }, () => row(1)),
    ];
    expect(build(rows).anomalies).toEqual([]);
  });
});

describe('never-used capabilities', () => {
  it('lists allowlisted capabilities nobody called in either window', () => {
    const report = build(
      [row(1, { tool: 'send-mail' })],
      ['send-mail', 'create-calendar', 'list-drives']
    );
    expect(report.neverUsed).toEqual(['create-calendar', 'list-drives']);
  });

  it('is empty when no allowlist is supplied, rather than claiming everything is unused', () => {
    expect(build([row(1)]).neverUsed).toEqual([]);
  });
});

describe('rendering', () => {
  const report = build([
    row(9, { upn: 'gone@x.com' }),
    row(1, { upn: 'here@x.com', tool: 'send-mail' }),
    row(2, { upn: 'here@x.com', tool: 'send-mail', outcome: 'error' }),
  ]);

  it('names the window and the single most notable thing in the subject', () => {
    const busy = build([
      ...Array.from({ length: 400 }, () => row(9, { upn: 'laura.mirarchi@x.com' })),
      ...Array.from({ length: 3 }, () => row(1, { upn: 'here@x.com' })),
    ]);
    expect(renderSubject(busy)).toBe(
      'M365 Connector: September 1 to September 8, Laura stopped after 400 calls'
    );
  });

  it('prefers a failing capability over a departed user in the headline', () => {
    const rows = [
      ...Array.from({ length: 400 }, () => row(9, { upn: 'laura.mirarchi@x.com' })),
      ...Array.from({ length: 10 }, (_, i) =>
        row(1, { tool: 'move-mail-message', outcome: i < 6 ? 'error' : 'success' })
      ),
    ];
    expect(renderSubject(build(rows))).toContain('move-mail-message failing 60% of the time');
  });

  it('says nothing unusual when there is nothing unusual', () => {
    const quiet = build([
      ...Array.from({ length: 5 }, () => row(9)),
      ...Array.from({ length: 5 }, () => row(1)),
    ]);
    expect(renderSubject(quiet)).toContain('nothing unusual');
    expect(renderHtml(quiet)).toContain('Nothing unusual this week');
  });

  it('sets Segoe UI 11pt on the body, the table and every cell', () => {
    const html = renderHtml(report);
    // Mail clients reset fonts inside tables, so inheritance is not enough.
    expect(html).toContain("font-family:'Segoe UI'");
    const cells = html.match(/<td style="[^"]*"/g) ?? [];
    expect(cells.length).toBeGreaterThan(0);
    for (const cell of cells) {
      expect(cell).toContain('font-size:11pt');
      expect(cell).toContain("'Segoe UI'");
      // A single-quoted attribute would be terminated by the quoted family name.
      expect(cell.startsWith('<td style="')).toBe(true);
    }
    expect(html).not.toContain('font-size:13px');
    expect(html).not.toContain('font-size:14px');
  });

  it('renders tables and the file-handling line', () => {
    const html = renderHtml(report);
    expect(html).toContain('<table');
    expect(html).toContain('Who used it');
    expect(html).toContain('File handling');
  });

  it('escapes anything that came out of the log', () => {
    const nasty = build([row(1, { upn: '<script>alert(1)</script>' })]);
    const html = renderHtml(nasty);
    expect(html).toContain('&lt;script&gt;');
    expect(html).not.toContain('<script>');
  });

  it('prints the credential countdown when an expiry is given', () => {
    const html = renderHtml(report, { secretExpiry: new Date('2028-04-20T00:00:00Z') });
    expect(html).toContain('2028-04-20');
  });

  it('leaves a one-line summary for the container log', () => {
    expect(renderLogLine(report)).toMatch(/calls=\d+ \(prior \d+\) errors=\d+/);
  });
});
