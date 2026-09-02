import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { parseArgs } from '../src/cli.js';
import { registerGraphTools } from '../src/graph-tools.js';
import type { GraphClient } from '../src/graph-client.js';

vi.mock('../src/cli.js', () => {
  const parseArgsMock = vi.fn();
  return {
    parseArgs: parseArgsMock,
  };
});

vi.mock('../src/generated/client.js', () => {
  return {
    api: {
      endpoints: [
        {
          alias: 'list-mail-messages',
          method: 'get',
          path: '/me/messages',
          parameters: [],
        },
        {
          alias: 'send-mail',
          method: 'post',
          path: '/me/sendMail',
          parameters: [],
        },
        {
          alias: 'delete-mail-message',
          method: 'delete',
          path: '/me/messages/{message-id}',
          parameters: [],
        },
        {
          alias: 'get-schedule',
          method: 'post',
          path: '/me/calendar/getSchedule',
          parameters: [],
        },
        {
          alias: 'update-mail-folder',
          method: 'patch',
          path: '/me/mailFolders/{mailFolder-id}',
          parameters: [],
        },
      ],
    },
  };
});

vi.mock('../src/logger.js', () => {
  return {
    default: {
      info: vi.fn(),
      error: vi.fn(),
    },
  };
});

/**
 * Custom (hand-registered) tools added by registerGraphTools on top of the
 * generated endpoints. Split by whether they survive read-only mode, so these
 * tests assert on identity rather than a count that rots whenever a tool is
 * added.
 */
const CUSTOM_TOOLS_READ_ONLY = [
  'parse-teams-url',
  'list-conversation-messages',
  'list-drafts',
  'get-messages-batch',
  'list-my-subscriptions',
  'check-notifications',
  'wait-for-notifications',
  'read-mail-attachment-text',
  'read-onedrive-file-text',
];

/** The above plus the ones that write, and so are skipped in read-only mode. */
const CUSTOM_TOOLS_ALL = [
  ...CUSTOM_TOOLS_READ_ONLY,
  'download-mail-attachment',
  'subscribe-to-changes',
  'unsubscribe-from-changes',
  'get-file',
  'attach-file',
  'put-file',
];

const sorted = (names: unknown[]) => [...(names as string[])].sort();

describe('Read-Only Mode', () => {
  let mockServer: { tool: ReturnType<typeof vi.fn> };

  beforeEach(() => {
    vi.clearAllMocks();

    delete process.env.READ_ONLY;

    mockServer = {
      tool: vi.fn(),
    };
  });

  afterEach(() => {
    vi.resetAllMocks();
  });

  it('should respect --read-only flag from CLI', () => {
    vi.mocked(parseArgs).mockReturnValue({ readOnly: true } as ReturnType<typeof parseArgs>);

    const options = parseArgs();
    expect(options.readOnly).toBe(true);

    registerGraphTools(mockServer, {} as GraphClient, options.readOnly);

    const toolCalls = mockServer.tool.mock.calls.map((call: unknown[]) => call[0]);
    // Only the GET endpoint survives read-only, alongside the non-writing customs.
    expect(sorted(toolCalls)).toEqual(sorted(['list-mail-messages', ...CUSTOM_TOOLS_READ_ONLY]));
    expect(toolCalls).toContain('list-mail-messages');
    expect(toolCalls).not.toContain('send-mail');
    expect(toolCalls).not.toContain('delete-mail-message');
  });

  it('should register all endpoints when not in read-only mode', () => {
    vi.mocked(parseArgs).mockReturnValue({ readOnly: false } as ReturnType<typeof parseArgs>);

    const options = parseArgs();
    expect(options.readOnly).toBe(false);

    registerGraphTools(mockServer, {} as GraphClient, options.readOnly);

    const toolCalls = mockServer.tool.mock.calls.map((call: unknown[]) => call[0]);
    // All 4 mocked endpoints (get-schedule needs orgMode) plus every custom tool.
    expect(sorted(toolCalls)).toEqual(
      sorted([
        'list-mail-messages',
        'send-mail',
        'delete-mail-message',
        'update-mail-folder',
        ...CUSTOM_TOOLS_ALL,
      ])
    );
    expect(toolCalls).toContain('list-mail-messages');
    expect(toolCalls).toContain('send-mail');
    expect(toolCalls).toContain('delete-mail-message');
    expect(toolCalls).toContain('update-mail-folder');
  });

  it('should allow POST endpoints with readOnly: true in endpoints.json in read-only mode', () => {
    // get-schedule is a POST endpoint with "readOnly": true in endpoints.json,
    // but it only has workScopes so orgMode must be enabled for it to be considered.
    const readOnly = true;
    const enabledToolsPattern = undefined;
    const orgMode = true;

    registerGraphTools(mockServer, {} as GraphClient, readOnly, enabledToolsPattern, orgMode);

    const toolCalls = mockServer.tool.mock.calls.map((call: unknown[]) => call[0]);

    // GET endpoint should be registered
    expect(toolCalls).toContain('list-mail-messages');
    // POST endpoint with readOnly: true should be registered
    expect(toolCalls).toContain('get-schedule');
    // Regular POST endpoint (no readOnly flag) should still be skipped
    expect(toolCalls).not.toContain('send-mail');
    // DELETE endpoint should still be skipped
    expect(toolCalls).not.toContain('delete-mail-message');
    // PATCH endpoint should still be skipped (readOnly bypass is POST-only)
    expect(toolCalls).not.toContain('update-mail-folder');

    // Exactly the two permitted Graph tools, plus the non-writing customs.
    expect(sorted(toolCalls)).toEqual(
      sorted(['list-mail-messages', 'get-schedule', ...CUSTOM_TOOLS_READ_ONLY])
    );
  });

  it('should block PATCH and DELETE endpoints in read-only mode regardless of readOnly flag', () => {
    // The readOnly: true bypass in endpoints.json only applies to POST methods.
    // PATCH and DELETE must always be blocked in read-only mode.
    const readOnly = true;
    const enabledToolsPattern = undefined;
    const orgMode = true;

    registerGraphTools(mockServer, {} as GraphClient, readOnly, enabledToolsPattern, orgMode);

    const toolCalls = mockServer.tool.mock.calls.map((call: unknown[]) => call[0]);

    // PATCH is always blocked in read-only mode
    expect(toolCalls).not.toContain('update-mail-folder');
    // DELETE is always blocked in read-only mode
    expect(toolCalls).not.toContain('delete-mail-message');
  });
});
