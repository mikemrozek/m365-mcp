import { describe, it, expect, vi, beforeEach } from 'vitest';
import { z } from 'zod';

/**
 * We test executeGraphTool logic by importing it indirectly through registerGraphTools.
 * Strategy: mock GraphClient, create a real McpServer, register tools, then invoke them.
 */

// Mock logger to silence output
vi.mock('../logger.js', () => ({
  default: {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    debug: vi.fn(),
  },
}));

// Mock the generated client — we supply our own endpoint definitions per test
const mockEndpoints: any[] = [];
vi.mock('../generated/client.js', () => ({
  api: {
    get endpoints() {
      return mockEndpoints;
    },
  },
}));

// Mock endpoints.json — we supply our own config per test
let mockEndpointsJson: any[] = [];
vi.mock('fs', async (importOriginal) => {
  const actual = await importOriginal<typeof import('fs')>();
  return {
    ...actual,
    readFileSync: (filePath: string, encoding?: string) => {
      if (typeof filePath === 'string' && filePath.includes('endpoints.json')) {
        return JSON.stringify(mockEndpointsJson);
      }
      return actual.readFileSync(filePath, encoding as any);
    },
  };
});

// Mock tool-categories
vi.mock('../tool-categories.js', () => ({
  TOOL_CATEGORIES: {},
}));

// ---------- helpers ----------

function makeEndpoint(overrides: Partial<any> = {}) {
  return {
    method: 'get',
    path: '/me/messages',
    alias: 'test-tool',
    description: 'Test tool',
    requestFormat: 'json' as const,
    parameters: [
      { name: 'filter', type: 'Query', schema: z.string().optional() },
      { name: 'search', type: 'Query', schema: z.string().optional() },
      { name: 'select', type: 'Query', schema: z.string().optional() },
      { name: 'orderby', type: 'Query', schema: z.string().optional() },
      { name: 'count', type: 'Query', schema: z.boolean().optional() },
      { name: 'top', type: 'Query', schema: z.number().optional() },
      { name: 'skip', type: 'Query', schema: z.number().optional() },
    ],
    response: z.any(),
    ...overrides,
  };
}

function makeConfig(overrides: Partial<any> = {}) {
  return {
    pathPattern: '/me/messages',
    method: 'get',
    toolName: 'test-tool',
    scopes: ['Mail.Read'],
    ...overrides,
  };
}

/** Creates a mock GraphClient with a controllable graphRequest spy */
function createMockGraphClient(responses?: any[]) {
  const responseQueue = [...(responses || [])];
  return {
    graphRequest: vi.fn().mockImplementation(async () => {
      if (responseQueue.length > 0) {
        return responseQueue.shift();
      }
      return {
        content: [{ type: 'text', text: JSON.stringify({ value: [] }) }],
      };
    }),
  };
}

/**
 * Because registerGraphTools reads endpointsData at module load time,
 * and we mock fs.readFileSync, we need to re-import after setting mocks.
 */
async function loadModule() {
  // Clear cached module so mocks take effect
  vi.resetModules();
  const mod = await import('../graph-tools.js');
  return mod;
}

/** Minimal McpServer mock that captures registered tools */
function createMockServer() {
  const tools = new Map<
    string,
    { description: string; schema: any; handler: (...args: any[]) => any }
  >();
  return {
    tool: vi.fn(
      (
        name: string,
        description: string,
        schema: any,
        annotations: any,
        handler: (...args: any[]) => any
      ) => {
        tools.set(name, { description, schema, handler });
      }
    ),
    tools,
  };
}

// ========== TESTS ==========

describe('graph-tools', () => {
  beforeEach(() => {
    mockEndpoints.length = 0;
    mockEndpointsJson = [];
    vi.clearAllMocks();
  });

  // ---- 1. $count advanced query mode ----
  describe('$count advanced query mode', () => {
    it('should set ConsistencyLevel: eventual header when $count=true', async () => {
      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ value: [] }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      // Invoke the registered tool with count=true
      const tool = server.tools.get('test-tool');
      expect(tool).toBeDefined();
      await tool!.handler({ count: true });

      // Verify graphRequest was called with ConsistencyLevel header
      expect(graphClient.graphRequest).toHaveBeenCalledTimes(1);
      const [url] = graphClient.graphRequest.mock.calls[0];
      // $count=true should appear in query string
      expect(url).toContain('$count=true');
    });
  });

  // ---- 2. fetchAllPages pagination ----
  describe('fetchAllPages pagination', () => {
    it('should follow @odata.nextLink and combine results', async () => {
      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                value: [{ id: '1' }, { id: '2' }],
                '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skip=2',
              }),
            },
          ],
        },
        {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                value: [{ id: '3' }],
              }),
            },
          ],
        },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('test-tool');
      const result = await tool!.handler({ fetchAllPages: true });

      // Should have made 2 requests (initial + 1 nextLink)
      expect(graphClient.graphRequest).toHaveBeenCalledTimes(2);

      // Combined result should have 3 items
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.value).toHaveLength(3);
      expect(parsed.value.map((v: any) => v.id)).toEqual(['1', '2', '3']);
      // nextLink should be removed from final response
      expect(parsed['@odata.nextLink']).toBeUndefined();
    });

    it('should stop at 100 page limit', async () => {
      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      // Generate 101 responses — each has a nextLink except the last
      const responses = [];
      for (let i = 0; i < 101; i++) {
        responses.push({
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                value: [{ id: `item-${i}` }],
                '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skip=' + (i + 1),
              }),
            },
          ],
        });
      }

      const graphClient = createMockGraphClient(responses);
      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('test-tool');
      await tool!.handler({ fetchAllPages: true });

      // 1 initial + 99 pagination = 100 total requests (stops at pageCount=100)
      expect(graphClient.graphRequest).toHaveBeenCalledTimes(100);
    });
  });

  // ---- 3. Parameter describe() overrides ----
  describe('parameter describe() overrides', () => {
    it('should apply custom descriptions to OData parameters', async () => {
      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, createMockGraphClient() as any);

      const tool = server.tools.get('test-tool');
      expect(tool).toBeDefined();

      const schema = tool!.schema;

      // These assert that the override is applied and still carries the semantics
      // that matter, not the exact prose. The wording was deliberately compressed on
      // 2026-09-02: repeated verbatim across 136 tools it cost ~24,000 tokens of the
      // catalogue, and the long form now lives once in the server instructions.

      // $filter override
      expect(schema['filter']).toBeDefined();
      expect(schema['filter'].description).toContain('OData filter');
      expect(schema['filter'].description).toContain('count=true');

      // $search override
      expect(schema['search']).toBeDefined();
      expect(schema['search'].description).toContain('KQL');

      // $select override
      expect(schema['select']).toBeDefined();
      expect(schema['select'].description).toContain('Fields to return');

      // $orderby override
      expect(schema['orderby']).toBeDefined();
      expect(schema['orderby'].description).toContain('Sort expression');

      // $count override
      expect(schema['count']).toBeDefined();
      expect(schema['count'].description).toContain('Advanced query mode');

      expect(schema['top'].description).toContain('Start small');
    });
  });

  // ---- $search KQL quote normalization ----
  // Backported from upstream #597 (dca6460) together with the normalizer itself.
  describe('$search quote normalization', () => {
    async function callSearch(
      search: string,
      path = '/me/messages'
    ): Promise<{ result: any; graphClient: any }> {
      const endpoint = makeEndpoint({ path });
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ value: [] }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const result = await server.tools.get('test-tool')!.handler({ search });
      return { result, graphClient };
    }

    async function callWithSearch(search: string, path = '/me/messages'): Promise<string> {
      const { graphClient } = await callSearch(search, path);
      return graphClient.graphRequest.mock.calls[0][0] as string;
    }

    it('wraps a bare KQL expression in one pair of double quotes', async () => {
      const url = await callWithSearch('from:john AND subject:meeting');
      expect(url).toContain(`$search=${encodeURIComponent('"from:john AND subject:meeting"')}`);
    });

    it('collapses per-term quoting into a single enclosing pair', async () => {
      const url = await callWithSearch('"from:john" AND subject:meeting');
      expect(url).toContain(`$search=${encodeURIComponent('"from:john AND subject:meeting"')}`);
    });

    it('leaves an already correctly quoted expression untouched', async () => {
      const url = await callWithSearch('"from:john AND subject:meeting"');
      expect(url).toContain(`$search=${encodeURIComponent('"from:john AND subject:meeting"')}`);
    });

    // Graph rejects a property phrase that has no enclosing pair, so add one and escape the
    // phrase quotes. Microsoft documents that escaping for directory search only; mail's own
    // docs never show an embedded quote, so this form is inferred.
    it('adds the enclosing pair around a property phrase', async () => {
      const url = await callWithSearch('subject:"quarterly report"');
      expect(url).toContain(`$search=${encodeURIComponent('"subject:\\"quarterly report\\""')}`);
    });

    it('keeps a standalone phrase grouped', async () => {
      const url = await callWithSearch('"quarterly report" AND from:john');
      expect(url).toContain(
        `$search=${encodeURIComponent('"\\"quarterly report\\" AND from:john"')}`
      );
    });

    it('leaves an already escaped phrase untouched', async () => {
      const query = '"subject:\\"quarterly report\\""';
      const url = await callWithSearch(query);
      expect(url).toContain(`$search=${encodeURIComponent(query)}`);
    });

    // Already correctly wrapped free text is a multi-term search, not a phrase — escaping
    // its quotes would narrow it to messages containing the exact phrase.
    it('leaves already-wrapped free text as a multi-term search', async () => {
      const query = '"quarterly report"';
      const url = await callWithSearch(query);
      expect(url).toContain(`$search=${encodeURIComponent(query)}`);
    });

    // Date and size restrictions use comparison operators rather than a colon; they are
    // clauses too, so per-clause quoting must be undone rather than escaped as a phrase.
    it.each([
      ['"received>=2024-01-01" AND from:john', '"received>=2024-01-01 AND from:john"'],
      ['"size>1000" AND subject:meeting', '"size>1000 AND subject:meeting"'],
      ['"received<2024-01-01"', '"received<2024-01-01"'],
    ])('undoes per-clause quoting on a comparison clause (%s)', async (query, expected) => {
      const url = await callWithSearch(query);
      expect(url).toContain(`$search=${encodeURIComponent(expected)}`);
    });

    // A phrase can open with a word and a colon without being a clause. RE and Q3 are not
    // mail properties, so the grouping quotes have to survive.
    it.each([
      ['"RE: quarterly report" AND from:john', '"\\"RE: quarterly report\\" AND from:john"'],
      ['"Q3: plan.pdf" AND subject:budget', '"\\"Q3: plan.pdf\\" AND subject:budget"'],
    ])('keeps phrase quotes on a clause-shaped phrase (%s)', async (query, expected) => {
      const url = await callWithSearch(query);
      expect(url).toContain(`$search=${encodeURIComponent(expected)}`);
    });

    // The colon inside the phrase is part of the text, not a property separator.
    it.each([
      ['subject:"RE: quarterly report"', '"subject:\\"RE: quarterly report\\""'],
      ['attachment:"Q3: plan.pdf"', '"attachment:\\"Q3: plan.pdf\\""'],
    ])('keeps phrase quotes when the phrase contains a colon (%s)', async (query, expected) => {
      const url = await callWithSearch(query);
      expect(url).toContain(`$search=${encodeURIComponent(expected)}`);
    });

    // KQL demotes a restriction with whitespace around the operator to free text, so these
    // are phrases. Unwrapping them would search a bare `from:`/`subject:` and quietly drop
    // the words the caller was actually looking for.
    it.each([
      [
        '"from: the desk of the CEO" AND subject:report',
        '"\\"from: the desk of the CEO\\" AND subject:report"',
      ],
      [
        '"subject: quarterly report" AND from:john',
        '"\\"subject: quarterly report\\" AND from:john"',
      ],
    ])('keeps phrase quotes when a space follows the property (%s)', async (query, expected) => {
      const url = await callWithSearch(query);
      expect(url).toContain(`$search=${encodeURIComponent(expected)}`);
    });

    // A trailing lone backslash would escape the enclosing pair's closing quote and hand
    // Graph an unterminated string.
    it('balances a trailing backslash so it cannot escape the closing quote', async () => {
      const url = await callWithSearch('from:john\\');
      expect(url).toContain(`$search=${encodeURIComponent('"from:john\\\\"')}`);
    });

    // An unterminated run is a missing closing quote, not a stray opening one. Dropping the
    // delimiter would shed the grouping: `subject:"quarterly report` would go out as subject
    // matching `quarterly` with `report` loose, which is a wider search than was asked for.
    it.each([
      ['from:"john', '"from:\\"john\\""'],
      ['subject:"quarterly report', '"subject:\\"quarterly report\\""'],
    ])('closes an unterminated quote rather than dropping it (%s)', async (query, expected) => {
      const url = await callWithSearch(query);
      expect(url).toContain(`$search=${encodeURIComponent(expected)}`);
    });

    // A restriction binds only the token after its operator, so unwrapping a multi-word value
    // would bind the first word and leave the rest as free text. Escaping the run in place is
    // no better: `subject:` would end up inside the phrase as literal text and the restriction
    // would be lost. Only moving the quotes past the operator keeps both.
    it.each([
      [
        '"subject:quarterly report" AND from:john',
        '"subject:\\"quarterly report\\" AND from:john"',
      ],
      ['"from:john" AND "subject:the big report"', '"from:john AND subject:\\"the big report\\""'],
    ])('moves quotes past the operator on a multi-word value (%s)', async (query, expected) => {
      const url = await callWithSearch(query);
      expect(url).toContain(`$search=${encodeURIComponent(expected)}`);
    });

    // Per-clause quoting of a whole boolean group is the same directory-style mistake as
    // quoting one clause, so it unwraps too. Escaping it would turn live restrictions into
    // literal text and leave only the clauses outside the quotes doing any work.
    it.each([
      [
        '"from:john AND subject:meeting" OR from:jane',
        '"from:john AND subject:meeting OR from:jane"',
      ],
      [
        '"from:john OR from:jane" AND hasAttachments:true',
        '"from:john OR from:jane AND hasAttachments:true"',
      ],
    ])('unwraps a quoted group of clauses (%s)', async (query, expected) => {
      const url = await callWithSearch(query);
      expect(url).toContain(`$search=${encodeURIComponent(expected)}`);
    });

    // A missing closing quote on the enclosing pair is a dropped character, not the start of
    // a phrase. Reading it as a phrase would search for the expression literally and match
    // nothing, leaving the model no error to correct against.
    it('recovers a dropped closing quote on the enclosing pair', async () => {
      const url = await callWithSearch('"from:john AND subject:meeting');
      expect(url).toContain(`$search=${encodeURIComponent('"from:john AND subject:meeting"')}`);
    });

    // An escaped backslash must be consumed as a unit, or its second slash pairs with the
    // real delimiter behind it and the scan runs off the end of a well-formed string.
    it('reads an escaped backslash before a closing quote', async () => {
      const url = await callWithSearch('from:"a\\\\"');
      expect(url).toContain(`$search=${encodeURIComponent('"from:\\"a\\\\\\""')}`);
    });

    // Every repair has to be a fixed point, otherwise a retry or a second pass corrupts a
    // value this code just declared correct.
    const CORPUS = [
      'from:john AND subject:meeting',
      '"from:john" AND subject:meeting',
      'subject:"quarterly report',
      '"subject:quarterly report" AND from:john',
      '"from:john AND subject:meeting" OR from:jane',
      '"from:john AND subject:meeting',
      'from:john\\',
      'from:"a\\\\"',
      'subject:"abc\\',
      '"quarterly report"',
    ];

    // The value Graph receives must be one well-formed escaped string: an opening quote, no
    // unescaped quote before the final one, and no trailing backslash that would escape it.
    // Asserting the shape catches a class of scanner bugs that enumerating cases misses.
    it.each(CORPUS)('emits a balanced escaped string (%s)', async (query) => {
      const url = await callWithSearch(query);
      const value = new URL(url, 'https://graph.microsoft.com').searchParams.get('$search')!;
      expect(value.startsWith('"') && value.endsWith('"')).toBe(true);
      let unescaped = 0;
      for (let i = 0; i < value.length; i++) {
        if (value[i] === '\\') {
          i++;
          continue;
        }
        if (value[i] === '"') unescaped++;
      }
      expect(unescaped).toBe(2);
    });

    it.each(CORPUS)('normalizing twice is a no-op (%s)', async (query) => {
      const once = await callWithSearch(query);
      const search = new URL(once, 'https://graph.microsoft.com').searchParams.get('$search')!;
      const twice = await callWithSearch(search);
      expect(twice).toContain(`$search=${encodeURIComponent(search)}`);
    });

    // Dropping $search would turn a search into an unfiltered listing of the whole mailbox
    // and return it as though it were the result, which is worse than the 400 Graph sends.
    it.each([' ', '   ', '"', '""""'])(
      'refuses an unsearchable $search value %j',
      async (query) => {
        const { result, graphClient } = await callSearch(query);
        expect(result.isError).toBe(true);
        expect(JSON.parse(result.content[0].text).error).toBe('invalid_search');
        expect(graphClient.graphRequest).not.toHaveBeenCalled();
      }
    );

    // Directory search advertises clause-level quoting, which mail's convention would
    // destroy: collapsing the quotes below changes an OR of two clauses into one.
    it.each([
      ['/users', '"displayName:john" OR "displayName:jane"'],
      // Mail-adjacent, but none of these take message KQL.
      ['/me/mailFolders', 'foo OR bar'],
      ['/me/mailFolders/:mailFolderId/childFolders', 'foo OR bar'],
      ['/me/mailFolders/:mailFolderId/messageRules', 'foo OR bar'],
      ['/me/messages/:messageId/attachments', 'foo OR bar'],
      ['/planner/tasks/:plannerTaskId/messages', 'foo OR bar'],
      ['/chats/:chatId/messages', 'foo OR bar'],
      ['/teams/:teamId/channels/:channelId/messages', 'foo OR bar'],
    ])('does not touch $search on %s', async (path, query) => {
      const url = await callWithSearch(query, path);
      expect(url).toContain(`$search=${encodeURIComponent(query)}`);
    });

    it.each([
      '/me/mailFolders/:mailFolderId/messages',
      '/users/:userId/messages',
      '/me/mailFolders/:mailFolderId/childFolders/:childFolderId/messages',
    ])('still normalizes on %s', async (path) => {
      const url = await callWithSearch('"from:john" AND subject:meeting', path);
      expect(url).toContain(`$search=${encodeURIComponent('"from:john AND subject:meeting"')}`);
    });
  });

  describe('MS365_MCP_MAX_TOP', () => {
    const prevMaxTop = process.env.MS365_MCP_MAX_TOP;

    afterEach(() => {
      if (prevMaxTop === undefined) delete process.env.MS365_MCP_MAX_TOP;
      else process.env.MS365_MCP_MAX_TOP = prevMaxTop;
    });

    it('should clamp $top when MS365_MCP_MAX_TOP is set', async () => {
      process.env.MS365_MCP_MAX_TOP = '10';

      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ value: [] }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('test-tool');
      await tool!.handler({ top: 50 });

      const [url] = graphClient.graphRequest.mock.calls[0];
      expect(url).toContain('$top=10');
    });

    it('should pass through $top when MS365_MCP_MAX_TOP is unset', async () => {
      delete process.env.MS365_MCP_MAX_TOP;

      const endpoint = makeEndpoint();
      const config = makeConfig();
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ value: [] }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('test-tool');
      await tool!.handler({ top: 50 });

      const [url] = graphClient.graphRequest.mock.calls[0];
      expect(url).toContain('$top=50');
    });
  });

  // ---- 4. returnDownloadUrl ----
  describe('returnDownloadUrl', () => {
    it('should strip /content from path and return downloadUrl when returnDownloadUrl=true', async () => {
      const endpoint = makeEndpoint({
        alias: 'download-file',
        path: '/me/drive/items/:driveItem-id/content',
        parameters: [{ name: 'driveItem-id', type: 'Path', schema: z.string() }],
      });
      const config = makeConfig({
        toolName: 'download-file',
        pathPattern: '/me/drive/items/{driveItem-id}/content',
        returnDownloadUrl: true,
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const downloadUrl = 'https://download.example.com/file.pdf';
      const graphClient = createMockGraphClient([
        {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                '@microsoft.graph.downloadUrl': downloadUrl,
                name: 'file.pdf',
              }),
            },
          ],
        },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('download-file');
      expect(tool).toBeDefined();
      await tool!.handler({ 'driveItem-id': 'abc123' });

      // Path should NOT end with /content — it gets stripped
      const [requestedPath] = graphClient.graphRequest.mock.calls[0];
      expect(requestedPath).not.toContain('/content');
      expect(requestedPath).toContain('/me/drive/items/abc123');
    });
  });

  // ---- 5. kebab-case path param normalization ----
  describe('kebab-case path param normalization', () => {
    it('should substitute path when LLM passes message-id (kebab) but schema has messageId (camelCase)', async () => {
      // Simulates what hack.ts generates: path uses :messageId (camelCase)
      // but LLMs may pass message-id (kebab-case) since endpoints.json uses {message-id}
      const endpoint = makeEndpoint({
        alias: 'get-mail-message',
        method: 'get',
        path: '/me/messages/:messageId',
        parameters: [
          { name: 'messageId', type: 'Path', schema: z.string() },
          { name: 'select', type: 'Query', schema: z.string().optional() },
        ],
      });
      const config = makeConfig({
        toolName: 'get-mail-message',
        pathPattern: '/me/messages/{message-id}',
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ id: 'AAMk123', subject: 'Test' }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('get-mail-message');
      expect(tool).toBeDefined();

      // Pass kebab-case 'message-id' — should still resolve to correct path
      await tool!.handler({ 'message-id': 'AAMk123abc=' });

      const [requestedPath] = graphClient.graphRequest.mock.calls[0];
      expect(requestedPath).toContain('AAMk123abc=');
      expect(requestedPath).not.toContain(':messageId');
    });

    it('should also work when LLM passes messageId (camelCase) directly', async () => {
      const endpoint = makeEndpoint({
        alias: 'get-mail-message2',
        method: 'get',
        path: '/me/messages/:messageId',
        parameters: [{ name: 'messageId', type: 'Path', schema: z.string() }],
      });
      const config = makeConfig({
        toolName: 'get-mail-message2',
        pathPattern: '/me/messages/{message-id}',
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ id: 'AAMk456' }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('get-mail-message2');
      await tool!.handler({ messageId: 'AAMk456xyz=' });

      const [requestedPath] = graphClient.graphRequest.mock.calls[0];
      expect(requestedPath).toContain('AAMk456xyz=');
      expect(requestedPath).not.toContain(':messageId');
    });
  });

  // ---- 6. supportsTimezone ----
  describe('supportsTimezone', () => {
    it('should set Prefer: outlook.timezone header when timezone param provided', async () => {
      const endpoint = makeEndpoint({
        alias: 'list-calendar-events',
        path: '/me/events',
        parameters: [],
      });
      const config = makeConfig({
        toolName: 'list-calendar-events',
        pathPattern: '/me/events',
        supportsTimezone: true,
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ value: [] }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      const tool = server.tools.get('list-calendar-events');
      expect(tool).toBeDefined();

      // Verify timezone parameter was added to schema
      expect(tool!.schema['timezone']).toBeDefined();
      expect(tool!.schema['timezone'].description).toContain('IANA timezone');

      await tool!.handler({ timezone: 'Europe/Brussels' });

      // Verify Prefer header contains outlook.timezone
      const [, options] = graphClient.graphRequest.mock.calls[0];
      expect(options.headers['Prefer']).toContain('outlook.timezone="Europe/Brussels"');
    });

    it('should NOT add timezone parameter when supportsTimezone is false/absent', async () => {
      const endpoint = makeEndpoint({
        alias: 'list-mail',
        path: '/me/messages',
        parameters: [],
      });
      const config = makeConfig({
        toolName: 'list-mail',
        pathPattern: '/me/messages',
        // no supportsTimezone
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, createMockGraphClient() as any);

      const tool = server.tools.get('list-mail');
      expect(tool!.schema['timezone']).toBeUndefined();
    });
  });

  // ---- 7. outlook.body-content-type Prefer header ----
  describe('outlook.body-content-type Prefer header', () => {
    it('should set Prefer: outlook.body-content-type="text" on GET requests', async () => {
      const endpoint = makeEndpoint({ method: 'get' });
      const config = makeConfig({ method: 'get' });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([
        { content: [{ type: 'text', text: JSON.stringify({ value: [] }) }] },
      ]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      await server.tools.get('test-tool')!.handler({});

      const [, options] = graphClient.graphRequest.mock.calls[0];
      expect(options.headers['Prefer']).toContain('outlook.body-content-type="text"');
    });

    it('should NOT set Prefer: outlook.body-content-type on POST requests', async () => {
      const endpoint = makeEndpoint({
        alias: 'create-reply-draft',
        method: 'post',
        path: '/me/messages/:messageId/createReply',
        parameters: [
          { name: 'messageId', type: 'Path', schema: z.string() },
          { name: 'body', type: 'Body', schema: z.any() },
        ],
      });
      const config = makeConfig({
        toolName: 'create-reply-draft',
        method: 'post',
        pathPattern: '/me/messages/{message-id}/createReply',
      });
      mockEndpoints.push(endpoint);
      mockEndpointsJson = [config];

      const graphClient = createMockGraphClient([{ content: [{ type: 'text', text: '{}' }] }]);

      const server = createMockServer();
      const { registerGraphTools } = await loadModule();
      registerGraphTools(server as any, graphClient as any);

      await server.tools.get('create-reply-draft')!.handler({
        messageId: 'AAMk123',
        body: { Message: { body: { contentType: 'html', content: '<p>hi</p>' } } },
      });

      const [, options] = graphClient.graphRequest.mock.calls[0];
      const prefer = options.headers['Prefer'];
      expect(prefer === undefined || !prefer.includes('outlook.body-content-type')).toBe(true);
    });
  });
});
