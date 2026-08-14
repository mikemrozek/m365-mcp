import { describe, it, expect } from 'vitest';
import { readFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { api } from '../src/generated/client.js';

/**
 * Guards the Everyday allowlist — the list that decides what the connector
 * advertises in production.
 *
 * The failure this prevents is silent: a name in the allowlist that doesn't
 * match a real tool is simply never registered, so the capability disappears
 * with no error anywhere. That is exactly how a rollout breaks someone's
 * workflow on a Monday morning with nothing in the logs to explain it.
 */

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));

const load = (name: string) =>
  JSON.parse(readFileSync(path.join(root, 'config/allowlists', name), 'utf8')) as {
    name: string;
    tools: string[];
  };

/** Hand-registered tools, which don't appear in the generated client. */
const CUSTOM_TOOLS = [
  'parse-teams-url',
  'list-conversation-messages',
  'list-drafts',
  'get-messages-batch',
  'download-mail-attachment',
  'subscribe-to-changes',
  'list-my-subscriptions',
  'unsubscribe-from-changes',
  'check-notifications',
  'wait-for-notifications',
  'read-mail-attachment-text',
  'read-onedrive-file-text',
  'get-file',
  'attach-file',
];

const universe = new Set<string>([
  ...api.endpoints.map((e) => e.alias as string),
  ...CUSTOM_TOOLS,
]);

describe('Everyday allowlist', () => {
  const everyday = load('everyday.json');

  it('names only tools that actually exist', () => {
    const phantom = everyday.tools.filter((t) => !universe.has(t));
    expect(phantom, `these would silently never register: ${phantom.join(', ')}`).toEqual([]);
  });

  it('contains no duplicates', () => {
    const dupes = everyday.tools.filter((t, i) => everyday.tools.indexOf(t) !== i);
    expect(dupes).toEqual([]);
  });

  it('is sorted, so diffs stay readable when it is re-cut', () => {
    expect(everyday.tools).toEqual([...everyday.tools].sort());
  });

  it('includes every capability shipped in tsq.13 and tsq.14', () => {
    // These are new, so no usage history would justify them — they have to be
    // added deliberately or the release ships dark.
    const shipped = [
      'subscribe-to-changes',
      'check-notifications',
      'wait-for-notifications',
      'list-my-subscriptions',
      'unsubscribe-from-changes',
      'get-file',
      'attach-file',
    ];
    const missing = shipped.filter((t) => !everyday.tools.includes(t));
    expect(missing, `shipped but not advertised: ${missing.join(', ')}`).toEqual([]);
  });

  it('is a real reduction against the 166 tools advertised today', () => {
    expect(everyday.tools.length).toBeLessThan(166);
    // Sanity floor: if this ever collapses to a handful, something built it wrong.
    expect(everyday.tools.length).toBeGreaterThan(100);
  });

  it('keeps the four tier allowlists internally valid too', () => {
    for (const file of ['core.json', 'docs-excel.json', 'collab-admin.json', 'mailbox-calendar.json']) {
      const tier = load(file);
      const phantom = tier.tools.filter((t) => !universe.has(t));
      expect(phantom, `${file} names non-existent tools: ${phantom.join(', ')}`).toEqual([]);
    }
  });
});
