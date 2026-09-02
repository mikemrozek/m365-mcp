import { describe, expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn() },
}));

/**
 * Guards the defect class found on 2026-09-02: 20 of the 136 advertised capabilities
 * opened with scraped Graph reference prose describing a DIFFERENT operation
 * (`delete-mail-message` — "Delete eventMessage"), or with a phrase carrying no
 * information at all ("The events in the calendar.", "Invoke action setReaction").
 *
 * This matters more than it used to. Claude no longer loads every tool definition
 * up front — it searches, and discovery matches on the opening text. A first
 * sentence about open extensions means the mail tool never surfaces for a mail
 * query, and argues against itself if it does.
 *
 * The failure mode has no symptom: nothing errors, the tool simply is not found,
 * and the model silently picks something worse. Hence a test rather than a note.
 */

const BAD_OPENERS: { pattern: RegExp; why: string }[] = [
  { pattern: /open extension/i, why: 'describes openTypeExtension, not the operation' },
  { pattern: /openTypeExtension/i, why: 'describes openTypeExtension, not the operation' },
  {
    pattern: /eventMessage/i,
    why: 'eventMessage is the meeting-invite subclass, not a mail message',
  },
  {
    pattern: /multi-value extended propert/i,
    why: 'describes extended properties, not the operation',
  },
  {
    pattern: /^Invoke action/i,
    why: 'names the Graph action verb and nothing a user would search for',
  },
  {
    pattern: /^The (events|messages|calendar view)\b/i,
    why: 'a noun phrase, not a description of what the tool does',
  },
];

async function registerAllowlistedTools(): Promise<{ name: string; description: string }[]> {
  const { registerGraphTools } = await import('../src/graph-tools.js');
  const allowlist: string[] = JSON.parse(
    readFileSync('config/allowlists/everyday.json', 'utf8')
  ).tools;

  const captured: { name: string; description: string }[] = [];
  const server = {
    tool: (...args: unknown[]) => {
      captured.push({
        name: String(args[0]),
        description: typeof args[1] === 'string' ? args[1] : '',
      });
    },
  };
  const graphClient = { makeRequest: async () => ({}), fetchBinary: async () => ({}) };

  registerGraphTools(
    server as never,
    graphClient as never,
    false,
    undefined,
    true,
    undefined,
    false,
    [],
    allowlist
  );
  return captured;
}

describe('advertised tool descriptions', () => {
  it('registers the whole everyday allowlist', async () => {
    const tools = await registerAllowlistedTools();
    expect(tools.length).toBe(124);
  });

  it('never opens with prose describing a different operation', async () => {
    const tools = await registerAllowlistedTools();
    const offenders: string[] = [];

    for (const tool of tools) {
      const opening = tool.description.split(/\n|(?<=\.)\s/)[0]?.trim() ?? '';
      for (const { pattern, why } of BAD_OPENERS) {
        if (pattern.test(opening)) {
          offenders.push(`${tool.name}: "${opening.slice(0, 60)}" — ${why}`);
          break;
        }
      }
    }

    expect(
      offenders,
      `Add a hand-written "description" in endpoints.json for:\n${offenders.join('\n')}`
    ).toEqual([]);
  });

  it('gives every capability an opening sentence long enough to match a query', async () => {
    const tools = await registerAllowlistedTools();
    const tooShort = tools
      .filter((t) => (t.description.split(/\n|(?<=\.)\s/)[0]?.trim().length ?? 0) < 25)
      .map((t) => `${t.name}: "${t.description.slice(0, 40)}"`);
    expect(tooShort).toEqual([]);
  });

  it('leads with a verb on the capabilities that were rewritten', async () => {
    const tools = await registerAllowlistedTools();
    const rewritten = [
      'list-mail-messages',
      'delete-mail-message',
      'create-draft-email',
      'insert-excel-range',
    ];
    for (const name of rewritten) {
      const tool = tools.find((t) => t.name === name);
      expect(tool, `${name} should be registered`).toBeDefined();
      expect(tool!.description).toMatch(
        /^(Search|Delete|Compose|React|Get|Update|Create|List|Insert|Forward|Remove)/
      );
    }
  });
});
