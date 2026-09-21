import { describe, it, expect, vi, beforeEach } from 'vitest';
import {
  registerInferenceTools,
  denyListCheck,
  resetInferenceBudget,
} from '../src/inference-tools.js';
import {
  FoundryError,
  inferenceConfigFromEnv,
  type ChatResult,
  type InferenceConfig,
} from '../src/lib/foundry-client.js';

/**
 * tsq.25 delegate-analysis. The properties under test are the boundary ones:
 * nothing leaves the server unless the caller is on the pilot list, the content
 * passes the pre-send policy check, and the budget holds — and the telemetry
 * that records all of it can never carry content. The failure that matters is
 * a silent one (content reaching a second model, or a record quoting it), so
 * the assertions here are mostly about absence.
 */

const { infoSpy, actorRef } = vi.hoisted(() => ({
  infoSpy: vi.fn(),
  actorRef: { current: undefined as { oid?: string; upn?: string; tid?: string } | undefined },
}));

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), warn: vi.fn(), error: vi.fn() },
}));
vi.mock('../src/usage-log.js', () => ({
  withUsageLog: (_name: string, fn: () => Promise<unknown>) => fn(),
  logToolUsage: vi.fn(),
  describeFailure: () => ({}),
  default: { info: infoSpy },
}));
vi.mock('../src/request-context.js', () => ({
  getRequestActor: () => actorRef.current,
  getRequestTokens: () => undefined,
}));

const PILOT_OID = 'pilot-oid-1';

const makeConfig = (over: Partial<InferenceConfig> = {}): InferenceConfig => ({
  endpoint: 'https://unit.test',
  deployment: 'deepseek-v4-flash',
  pilotOids: [PILOT_OID],
  maxInputChars: 200_000,
  dailyTokenBudget: 2_000_000,
  ...over,
});

const OK_RESULT: ChatResult = {
  content: 'ANALYSIS-OUTPUT',
  usage: { promptTokens: 100, completionTokens: 20, totalTokens: 120 },
};

type Handler = (params: Record<string, unknown>) => Promise<{
  content: { type: string; text: string }[];
  isError?: boolean;
}>;

function setup(configOver: Partial<InferenceConfig> = {}, chatImpl?: () => Promise<ChatResult>) {
  const tools = new Map<string, Handler>();
  const server = {
    tool: vi.fn(
      (name: string, _d: string, _s: unknown, _a: unknown, handler: Handler) =>
        void tools.set(name, handler)
    ),
  };
  const graphClient = { makeRequest: vi.fn(), fetchBinary: vi.fn() };
  const chat = vi.fn(chatImpl ?? (async () => OK_RESULT));
  const hooks = { isToolEnabled: () => true, push: vi.fn(), fail: vi.fn() };
  registerInferenceTools(server as never, graphClient as never, hooks, {
    config: makeConfig(configOver),
    chat,
  });
  return { server, graphClient, chat, hooks, handler: tools.get('delegate-analysis')! };
}

const inferenceRecords = () =>
  infoSpy.mock.calls
    .filter((c) => (c[1] as Record<string, unknown>)?.type === 'm365-inference')
    .map((c) => c[1] as Record<string, unknown>);

const parse = (r: { content: { text: string }[] }) => JSON.parse(r.content[0].text);

beforeEach(() => {
  infoSpy.mockClear();
  actorRef.current = { oid: PILOT_OID, upn: 'pilot@townsquaremedia.com', tid: 'tid-1' };
  resetInferenceBudget();
});

// ---- registration gating ----------------------------------------------------

describe('registration', () => {
  it('does not register when the env is absent', () => {
    delete process.env.INFERENCE_ENDPOINT;
    delete process.env.INFERENCE_DEPLOYMENT;
    const server = { tool: vi.fn() };
    const hooks = { isToolEnabled: () => true, push: vi.fn(), fail: vi.fn() };
    registerInferenceTools(server as never, {} as never, hooks, {});
    expect(server.tool).not.toHaveBeenCalled();
    expect(hooks.push).not.toHaveBeenCalled();
  });

  it('does not register when endpoint is set but deployment is missing', () => {
    process.env.INFERENCE_ENDPOINT = 'https://unit.test';
    delete process.env.INFERENCE_DEPLOYMENT;
    try {
      const server = { tool: vi.fn() };
      const hooks = { isToolEnabled: () => true, push: vi.fn(), fail: vi.fn() };
      registerInferenceTools(server as never, {} as never, hooks, {});
      expect(server.tool).not.toHaveBeenCalled();
    } finally {
      delete process.env.INFERENCE_ENDPOINT;
    }
  });

  it('registers when configured, and reports the name through hooks', () => {
    const { server, hooks } = setup();
    expect(server.tool).toHaveBeenCalledTimes(1);
    expect(hooks.push).toHaveBeenCalledWith('delegate-analysis');
  });
});

// ---- the pilot gate -----------------------------------------------------------

describe('pilot gate', () => {
  it('rejects a caller who is not on the pilot list, without calling anything', async () => {
    actorRef.current = { oid: 'stranger-oid', upn: 'stranger@x.com' };
    const { handler, chat, graphClient } = setup();
    const result = await handler({ task: 'summarize', text: 'hello' });
    expect(result.isError).toBe(true);
    expect(parse(result).error).toBe('not_enabled');
    expect(chat).not.toHaveBeenCalled();
    expect(graphClient.makeRequest).not.toHaveBeenCalled();
    expect(inferenceRecords()).toEqual([]);
  });

  it('fails closed when there is no caller identity at all', async () => {
    actorRef.current = undefined;
    const { handler, chat } = setup();
    const result = await handler({ task: 'summarize', text: 'hello' });
    expect(parse(result).error).toBe('not_enabled');
    expect(chat).not.toHaveBeenCalled();
  });

  it('fails closed on an empty pilot list', async () => {
    const { handler, chat } = setup({ pilotOids: [] });
    const result = await handler({ task: 'summarize', text: 'hello' });
    expect(parse(result).error).toBe('not_enabled');
    expect(chat).not.toHaveBeenCalled();
  });
});

// ---- the pre-send policy check --------------------------------------------------

describe('denyListCheck', () => {
  it('blocks an SSN-shaped value', () => {
    expect(denyListCheck('employee record 123-45-6789 attached')).toEqual({
      blocked: true,
      rule: 'ssn',
    });
  });

  it('blocks a Luhn-valid card number', () => {
    expect(denyListCheck('card 4111 1111 1111 1111 on file')).toEqual({
      blocked: true,
      rule: 'card_number',
    });
  });

  it('does not block a 16-digit run that fails Luhn — ids are not cards', () => {
    expect(denyListCheck('invoice 1234 5678 9012 3456 paid')).toEqual({ blocked: false });
  });

  it('blocks a private key marker', () => {
    expect(denyListCheck('-----BEGIN RSA PRIVATE KEY-----\nabc')).toEqual({
      blocked: true,
      rule: 'private_key',
    });
  });

  it('passes ordinary business text', () => {
    expect(denyListCheck('Q3 revenue was $4.2M against a plan of $3.9M.')).toEqual({
      blocked: false,
    });
  });
});

describe('policy enforcement in the handler', () => {
  it('blocks before the provider is called, and logs a content-free record', async () => {
    const { handler, chat } = setup();
    const result = await handler({ task: 'summarize', text: 'ssn is 123-45-6789' });
    expect(result.isError).toBe(true);
    expect(parse(result).error).toBe('blocked_policy');
    expect(parse(result).rule).toBe('ssn');
    expect(chat).not.toHaveBeenCalled();

    const records = inferenceRecords();
    expect(records).toHaveLength(1);
    expect(records[0].outcome).toBe('blocked_policy');
    expect(records[0].code).toBe('ssn');
    expect(JSON.stringify(records[0])).not.toContain('123-45-6789');
  });
});

// ---- caps and assembly -----------------------------------------------------------

describe('input handling', () => {
  it('refuses over-limit content rather than truncating silently', async () => {
    const { handler, chat } = setup({ maxInputChars: 50 });
    const result = await handler({ task: 'summarize', text: 'x'.repeat(80) });
    expect(parse(result).error).toBe('input_too_large');
    expect(chat).not.toHaveBeenCalled();
    expect(inferenceRecords()).toEqual([]);
  });

  it('refuses a call with nothing to analyze', async () => {
    const { handler, chat } = setup();
    const result = await handler({ task: 'summarize' });
    expect(parse(result).error).toBe('nothing_to_analyze');
    expect(chat).not.toHaveBeenCalled();
  });

  it('fetches messages with the text-body preference, deduplicated, and passes their content on', async () => {
    const { handler, chat, graphClient } = setup();
    graphClient.makeRequest.mockResolvedValue({
      subject: 'S1',
      receivedDateTime: '2026-09-21T12:00:00Z',
      from: { emailAddress: { address: 'a@x.com', name: 'A' } },
      body: { content: 'MAILBODY-ONE' },
    });

    await handler({ task: 'summarize', messageIds: ['m1', 'm2', 'm1'] });

    expect(graphClient.makeRequest).toHaveBeenCalledTimes(2);
    const [url, options] = graphClient.makeRequest.mock.calls[0];
    expect(url).toContain('/me/messages/m1');
    expect((options as { headers: Record<string, string> }).headers.Prefer).toContain(
      'outlook.body-content-type'
    );
    const sent = chat.mock.calls[0][0] as { user: string };
    expect(sent.user).toContain('MAILBODY-ONE');
    expect(inferenceRecords()[0].nMessages).toBe(2);
  });

  it('fails the whole call when a message fetch fails, rather than analyzing a partial set', async () => {
    const { handler, chat, graphClient } = setup();
    graphClient.makeRequest.mockRejectedValue(new Error('404'));
    const result = await handler({ task: 'summarize', messageIds: ['gone'] });
    expect(parse(result).error).toBe('fetch_failed');
    expect(chat).not.toHaveBeenCalled();
  });

  it('extracts text from a drive item and counts it', async () => {
    const { handler, chat, graphClient } = setup();
    graphClient.makeRequest.mockResolvedValue({
      name: 'notes.txt',
      size: 9,
      file: { mimeType: 'text/plain' },
    });
    graphClient.fetchBinary.mockResolvedValue({ buffer: Buffer.from('FILE-TEXT') });

    await handler({ task: 'summarize', driveItemIds: ['f1'] });

    const sent = chat.mock.calls[0][0] as { user: string };
    expect(sent.user).toContain('FILE-TEXT');
    expect(inferenceRecords()[0].nFiles).toBe(1);
  });
});

// ---- budget -----------------------------------------------------------------------

describe('budget', () => {
  it('rejects once the daily budget is spent, and recovers on reset', async () => {
    // The gate is a PRE-call check — a call's cost is unknowable beforehand —
    // so the ceiling is soft by at most one call: the first call (120 tokens)
    // is allowed against a budget of 100, and the second is refused.
    const { handler, chat } = setup({ dailyTokenBudget: 100 });

    const first = await handler({ task: 'summarize', text: 'fine' });
    expect(first.isError).toBeUndefined();

    const second = await handler({ task: 'summarize', text: 'fine again' });
    expect(parse(second).error).toBe('over_budget');
    expect(chat).toHaveBeenCalledTimes(1);
    expect(inferenceRecords().at(-1)?.outcome).toBe('over_budget');

    resetInferenceBudget();
    const third = await handler({ task: 'summarize', text: 'after reset' });
    expect(third.isError).toBeUndefined();
  });
});

// ---- the provider call and its record ------------------------------------------------

describe('provider call', () => {
  it('asserts the human caller in user_security_context, snake_case', async () => {
    const { handler, chat } = setup();
    await handler({ task: 'summarize', text: 'fine' });
    const sent = chat.mock.calls[0][0] as { userContext: Record<string, string> };
    expect(sent.userContext).toEqual({
      application_name: 'tsq-m365-mcp',
      end_user_id: PILOT_OID,
      end_user_tenant_id: 'tid-1',
    });
  });

  it('returns the output verbatim behind a provenance line', async () => {
    const { handler } = setup();
    const result = await handler({ task: 'summarize', text: 'fine' });
    expect(result.content[0].text).toMatch(
      /^\[Machine-generated by delegated model deepseek-v4-flash/
    );
    expect(result.content[0].text).toContain('\n\nANALYSIS-OUTPUT');
  });

  it('maps a provider failure to a clean error and a status/code record', async () => {
    const { handler } = setup({}, async () => {
      throw new FoundryError('HTTP 503', 'provider_error', 503, 'ServerBusy');
    });
    const result = await handler({ task: 'summarize', text: 'fine' });
    expect(parse(result)).toMatchObject({ error: 'provider_error', status: 503 });

    const record = inferenceRecords()[0];
    expect(record).toMatchObject({ outcome: 'provider_error', status: 503, code: 'ServerBusy' });
    expect(typeof record.latencyMs).toBe('number');
  });

  it('surfaces retry-after on rate limiting', async () => {
    const { handler } = setup({}, async () => {
      throw new FoundryError('HTTP 429', 'rate_limited', 429, undefined, 30);
    });
    const result = await handler({ task: 'summarize', text: 'fine' });
    expect(parse(result)).toMatchObject({ error: 'rate_limited', retryAfterSeconds: 30 });
  });
});

// ---- telemetry is content-free ------------------------------------------------------

describe('telemetry', () => {
  it('records tokens and shape, never task or content', async () => {
    const { handler } = setup();
    await handler({ task: 'summarize SECRET-TASK-PHRASE', text: 'SECRET-BODY-PHRASE etc' });

    const records = inferenceRecords();
    expect(records).toHaveLength(1);
    const record = records[0];
    expect(record).toMatchObject({
      type: 'm365-inference',
      outcome: 'success',
      model: 'deepseek-v4-flash',
      oid: PILOT_OID,
      inputTokens: 100,
      outputTokens: 20,
      nMessages: 0,
      nFiles: 0,
      inlineChars: 'SECRET-BODY-PHRASE etc'.length,
    });
    const serialized = JSON.stringify(record);
    expect(serialized).not.toContain('SECRET-TASK-PHRASE');
    expect(serialized).not.toContain('SECRET-BODY-PHRASE');
    expect(serialized).not.toContain('ANALYSIS-OUTPUT');
    expect('task' in record).toBe(false);
    expect('text' in record).toBe(false);
    expect('content' in record).toBe(false);
  });
});

// ---- config parsing -------------------------------------------------------------------

describe('inferenceConfigFromEnv', () => {
  it('is undefined without the endpoint or without the deployment', () => {
    expect(inferenceConfigFromEnv({})).toBeUndefined();
    expect(inferenceConfigFromEnv({ INFERENCE_ENDPOINT: 'https://x' })).toBeUndefined();
  });

  it('applies defaults, trims, and strips the trailing slash', () => {
    const config = inferenceConfigFromEnv({
      INFERENCE_ENDPOINT: 'https://x.services.ai.azure.com/',
      INFERENCE_DEPLOYMENT: ' deepseek-v4-flash ',
    })!;
    expect(config.endpoint).toBe('https://x.services.ai.azure.com');
    expect(config.deployment).toBe('deepseek-v4-flash');
    expect(config.pilotOids).toEqual([]);
    expect(config.maxInputChars).toBe(200_000);
    expect(config.dailyTokenBudget).toBe(2_000_000);
  });

  it('treats a malformed pilot list as EMPTY, never as everyone', () => {
    const bad = inferenceConfigFromEnv({
      INFERENCE_ENDPOINT: 'https://x',
      INFERENCE_DEPLOYMENT: 'm',
      INFERENCE_PILOT_OIDS: 'not json at all',
    })!;
    expect(bad.pilotOids).toEqual([]);

    const wrongShape = inferenceConfigFromEnv({
      INFERENCE_ENDPOINT: 'https://x',
      INFERENCE_DEPLOYMENT: 'm',
      INFERENCE_PILOT_OIDS: '{"oids": ["a"]}',
    })!;
    expect(wrongShape.pilotOids).toEqual([]);
  });

  it('parses a valid pilot list', () => {
    const config = inferenceConfigFromEnv({
      INFERENCE_ENDPOINT: 'https://x',
      INFERENCE_DEPLOYMENT: 'm',
      INFERENCE_PILOT_OIDS: '["oid-a", " oid-b "]',
    })!;
    expect(config.pilotOids).toEqual(['oid-a', 'oid-b']);
  });
});
