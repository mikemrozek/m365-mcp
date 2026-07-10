import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { mkdtempSync, writeFileSync, rmSync } from 'fs';
import { tmpdir } from 'os';
import path from 'path';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { loadToolAllowlist } from '../src/tool-allowlist.js';
import { registerGraphTools } from '../src/graph-tools.js';
import GraphClient from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    debug: vi.fn(),
  },
}));

// Mock endpoint set: a couple of Graph endpoints. The five custom tools
// (parse-teams-url, list-conversation-messages, list-drafts, get-messages-batch,
// download-mail-attachment) are registered unconditionally by registerGraphTools
// and are also subject to the allowlist.
vi.mock('../src/generated/client.js', () => ({
  api: {
    endpoints: [
      { alias: 'get-current-user', method: 'GET', path: '/me', description: 'Get current user' },
      {
        alias: 'list-mail-messages',
        method: 'GET',
        path: '/me/messages',
        description: 'List mail messages',
      },
      { alias: 'send-mail', method: 'POST', path: '/me/sendMail', description: 'Send mail' },
    ],
  },
}));

describe('loadToolAllowlist', () => {
  let tmpDir: string;

  beforeEach(() => {
    tmpDir = mkdtempSync(path.join(tmpdir(), 'allowlist-'));
  });

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true });
  });

  it('parses an inline JSON array', () => {
    const { names, source } = loadToolAllowlist('["get-current-user","list-mail-messages"]');
    expect(names).toEqual(['get-current-user', 'list-mail-messages']);
    expect(source).toContain('inline JSON');
    expect(source).toContain('2 tools');
  });

  it('parses an inline JSON object with a tools array', () => {
    const { names } = loadToolAllowlist('{"name":"core","tools":["b-tool","a-tool"]}');
    // de-duplicated and sorted
    expect(names).toEqual(['a-tool', 'b-tool']);
  });

  it('de-duplicates and trims names', () => {
    const { names } = loadToolAllowlist('[" x ","x","y"]');
    expect(names).toEqual(['x', 'y']);
  });

  it('reads a JSON array from a file path', () => {
    const file = path.join(tmpDir, 'core.json');
    writeFileSync(file, JSON.stringify(['get-current-user']));
    const { names, source } = loadToolAllowlist(file);
    expect(names).toEqual(['get-current-user']);
    expect(source).toContain(file);
  });

  it('throws on empty input', () => {
    expect(() => loadToolAllowlist('   ')).toThrow(/empty/i);
  });

  it('throws on invalid JSON', () => {
    expect(() => loadToolAllowlist('[not json')).toThrow(/not valid JSON/i);
  });

  it('throws when the JSON is neither array nor {tools:[]}', () => {
    expect(() => loadToolAllowlist('{"foo":1}')).toThrow(/must be a JSON array/i);
  });

  it('throws on a non-string entry', () => {
    expect(() => loadToolAllowlist('["ok", 42]')).toThrow(/non-string or empty/i);
  });

  it('throws on an unreadable file path', () => {
    expect(() => loadToolAllowlist('/no/such/allowlist/file.json')).toThrow(/could not be read/i);
  });
});

describe('registerGraphTools with an allowlist', () => {
  let server: McpServer;
  let graphClient: GraphClient;
  let toolSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    server = new McpServer({ name: 'test', version: '1.0.0' });
    graphClient = {} as GraphClient;
    toolSpy = vi.spyOn(server, 'tool').mockImplementation(() => {});
  });

  const registeredNames = () => toolSpy.mock.calls.map((c) => c[0] as string).sort();

  it('registers only the exact tools on the allowlist (endpoint + custom)', () => {
    registerGraphTools(
      server,
      graphClient,
      false, // readOnly
      undefined, // enabledToolsPattern
      false, // orgMode
      undefined, // authManager
      false, // multiAccount
      [], // accountNames
      ['get-current-user', 'list-conversation-messages']
    );

    expect(registeredNames()).toEqual(['get-current-user', 'list-conversation-messages']);
  });

  it('exact-matches — a substring does not leak a tool in', () => {
    // 'mail' would match list-mail-messages/send-mail/download-mail-attachment under
    // the old regex; as an allowlist entry it matches nothing.
    registerGraphTools(server, graphClient, false, undefined, false, undefined, false, [], ['mail']);

    expect(toolSpy).toHaveBeenCalledTimes(0);
  });

  it('ignores the ENABLED_TOOLS regex when an allowlist is supplied', () => {
    registerGraphTools(
      server,
      graphClient,
      false,
      'mail', // regex that would match mail tools — must be ignored
      false,
      undefined,
      false,
      [],
      ['get-current-user']
    );

    expect(registeredNames()).toEqual(['get-current-user']);
  });

  it('an empty allowlist registers nothing', () => {
    registerGraphTools(server, graphClient, false, undefined, false, undefined, false, [], []);
    expect(toolSpy).toHaveBeenCalledTimes(0);
  });

  it('still honours read-only mode for allowlisted write tools', () => {
    registerGraphTools(
      server,
      graphClient,
      true, // readOnly
      undefined,
      false,
      undefined,
      false,
      [],
      ['get-current-user', 'send-mail'] // send-mail is a POST write, must be dropped
    );

    expect(registeredNames()).toEqual(['get-current-user']);
  });
});
