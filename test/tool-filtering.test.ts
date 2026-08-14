import { beforeEach, describe, expect, it, vi } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { registerGraphTools } from '../src/graph-tools.js';
import GraphClient from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    error: vi.fn(),
  },
}));

vi.mock('../src/generated/client.js', () => ({
  api: {
    endpoints: [
      {
        alias: 'list-mail-messages',
        method: 'GET',
        path: '/me/messages',
        description: 'List mail messages',
      },
      { alias: 'send-mail', method: 'POST', path: '/me/sendMail', description: 'Send mail' },
      {
        alias: 'list-calendar-events',
        method: 'GET',
        path: '/me/events',
        description: 'List calendar events',
      },
      {
        alias: 'list-excel-worksheets',
        method: 'GET',
        path: '/workbook/worksheets',
        description: 'List Excel worksheets',
      },
      { alias: 'get-current-user', method: 'GET', path: '/me', description: 'Get current user' },
    ],
  },
}));

/**
 * Custom (hand-registered) tools that `registerGraphTools` adds on top of the
 * generated endpoints. Listed explicitly rather than baked into a count so that
 * adding a tool fails these tests with a readable diff of names, and has to be
 * accounted for deliberately.
 */
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

const MOCKED_ENDPOINTS = [
  'list-mail-messages',
  'send-mail',
  'list-calendar-events',
  'list-excel-worksheets',
  'get-current-user',
];

describe('Tool Filtering', () => {
  let server: McpServer;
  let graphClient: GraphClient;
  let toolSpy: ReturnType<typeof vi.spyOn>;

  /** Names of every tool registered, sorted — asserting on these beats counting. */
  const registered = () => toolSpy.mock.calls.map((c) => c[0] as string).sort();
  const sorted = (names: string[]) => [...names].sort();

  beforeEach(() => {
    server = new McpServer({ name: 'test', version: '1.0.0' });
    graphClient = {} as GraphClient;
    toolSpy = vi.spyOn(server, 'tool').mockImplementation(() => {});
  });

  it('should register all tools when no filter is provided', () => {
    registerGraphTools(server, graphClient, false);

    // Every mocked endpoint plus every custom tool, with nothing filtered out.
    expect(registered()).toEqual(sorted([...MOCKED_ENDPOINTS, ...CUSTOM_TOOLS]));
    expect(toolSpy).toHaveBeenCalledWith(
      'list-mail-messages',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'send-mail',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'list-calendar-events',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'list-excel-worksheets',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'get-current-user',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
  });

  it('should filter tools by regex pattern - mail only', () => {
    registerGraphTools(server, graphClient, false, 'mail');

    // Two mocked endpoints match 'mail', as do two custom tools.
    expect(registered()).toEqual(
      sorted([
        'list-mail-messages',
        'send-mail',
        'download-mail-attachment',
        'read-mail-attachment-text',
      ])
    );
  });

  it('should filter tools by regex pattern - calendar or excel', () => {
    registerGraphTools(server, graphClient, false, 'calendar|excel');

    // No custom tool name contains 'calendar' or 'excel'.
    expect(registered()).toEqual(sorted(['list-calendar-events', 'list-excel-worksheets']));
    expect(toolSpy).toHaveBeenCalledWith(
      'list-calendar-events',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'list-excel-worksheets',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
  });

  it('should handle invalid regex patterns gracefully', () => {
    registerGraphTools(server, graphClient, false, '[invalid regex');

    // An unusable pattern must fall back to registering everything, not nothing.
    expect(registered()).toEqual(sorted([...MOCKED_ENDPOINTS, ...CUSTOM_TOOLS]));
  });

  it('should combine read-only and filtering correctly', () => {
    registerGraphTools(server, graphClient, true, 'mail');

    // Read-only drops the writes: send-mail (POST) and download-mail-attachment
    // (stages a copy to OneDrive). Text extraction writes nothing, so it stays.
    expect(registered()).toEqual(sorted(['list-mail-messages', 'read-mail-attachment-text']));
  });

  it('should register no tools when pattern matches nothing', () => {
    registerGraphTools(server, graphClient, false, 'nonexistent');

    expect(toolSpy).toHaveBeenCalledTimes(0);
  });
});
