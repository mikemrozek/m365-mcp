import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { registerNotifyTools, isWorkflowWebhook, workflowPayload } from '../src/notify-tools.js';
import type GraphClient from '../src/graph-client.js';

type Handler = (p: Record<string, unknown>) => Promise<{
  content: { text: string }[];
  isError?: boolean;
}>;

function harness(makeRequest: ReturnType<typeof vi.fn>, readOnly = false) {
  const handlers: Record<string, Handler> = {};
  const schemas: Record<string, Record<string, unknown>> = {};
  const server = {
    tool: (name: string, _d: string, schema: Record<string, unknown>, _a: unknown, h: Handler) => {
      handlers[name] = h;
      schemas[name] = schema;
    },
  } as unknown as Parameters<typeof registerNotifyTools>[0];
  const skipped: string[] = [];
  registerNotifyTools(server, { makeRequest } as unknown as GraphClient, {
    isToolEnabled: () => true,
    readOnly,
    push: () => {},
    fail: (n, e) => {
      throw new Error(`${n}: ${e.message}`);
    },
    skip: (n) => skipped.push(n),
  });
  return { handlers, schemas, skipped };
}

const call = async (h: Handler, p: Record<string, unknown>) => {
  const r = await h(p);
  return { payload: JSON.parse(r.content[0].text), isError: r.isError };
};

const HOOK =
  'https://prod-12.westus.logic.azure.com:443/workflows/abc/triggers/manual/paths/invoke?sig=x';

describe('notify-me', () => {
  let makeRequest: ReturnType<typeof vi.fn>;
  let fetchSpy: ReturnType<typeof vi.fn>;
  beforeEach(() => {
    makeRequest = vi.fn(async (url: string) =>
      url.startsWith('/me?') ? { mail: 'me@tsq.com' } : undefined
    );
    fetchSpy = vi.fn();
    vi.stubGlobal('fetch', fetchSpy);
  });
  afterEach(() => vi.unstubAllGlobals());

  it('has no recipient parameter — it can only reach the user', () => {
    const { schemas } = harness(makeRequest);
    const keys = Object.keys(schemas['notify-me']);
    expect(keys.sort()).toEqual(['link', 'message', 'teamsWebhookUrl', 'title']);
  });

  it('accepts only Power Automate webhook hosts', () => {
    expect(isWorkflowWebhook(HOOK)).toBe(true);
    expect(isWorkflowWebhook('https://abc.environment.api.powerplatform.com/powerautomate/x')).toBe(
      true
    );
    expect(isWorkflowWebhook('https://evil.example.com/logic.azure.com')).toBe(false);
    expect(isWorkflowWebhook('http://prod-1.westus.logic.azure.com/x')).toBe(false);
    expect(isWorkflowWebhook('not a url')).toBe(false);
  });

  it('posts an Adaptive Card to the webhook and sends no email', async () => {
    fetchSpy.mockResolvedValue({ ok: true, status: 202 });
    const { handlers } = harness(makeRequest);
    const { payload } = await call(handlers['notify-me'], {
      message: '2 replies came in',
      title: 'Replies',
      teamsWebhookUrl: HOOK,
    });
    expect(payload).toEqual({ delivered: true, channel: 'teams' });
    const sent = JSON.parse(fetchSpy.mock.calls[0][1].body);
    expect(sent.attachments[0].contentType).toBe('application/vnd.microsoft.card.adaptive');
    expect(makeRequest).not.toHaveBeenCalled();
  });

  it('refuses a non-Workflows URL and sends nothing', async () => {
    const { handlers } = harness(makeRequest);
    const { isError, payload } = await call(handlers['notify-me'], {
      message: 'x',
      teamsWebhookUrl: 'https://example.com/hook',
    });
    expect(isError).toBe(true);
    expect(payload.error).toBe('not_a_workflow_webhook');
    expect(fetchSpy).not.toHaveBeenCalled();
    expect(makeRequest).not.toHaveBeenCalled();
  });

  it('emails only the user’s own address when no webhook is given', async () => {
    const { handlers } = harness(makeRequest);
    const { payload } = await call(handlers['notify-me'], { message: 'Overdue: Q4 budget' });
    expect(payload).toMatchObject({ delivered: true, channel: 'email', to: 'me@tsq.com' });
    const [url, opts] = makeRequest.mock.calls[1];
    expect(url).toBe('/me/sendMail');
    const body = JSON.parse(opts.body);
    expect(body.message.toRecipients).toEqual([{ emailAddress: { address: 'me@tsq.com' } }]);
    expect(body.message.subject).toContain('[Claude]');
  });

  it('falls back to email when the webhook fails, and says so', async () => {
    fetchSpy.mockResolvedValue({ ok: false, status: 404 });
    const { handlers } = harness(makeRequest);
    const { payload } = await call(handlers['notify-me'], { message: 'x', teamsWebhookUrl: HOOK });
    expect(payload.channel).toBe('email');
    expect(payload.teamsError).toContain('404');
  });

  it('escapes HTML in the email body', async () => {
    const { handlers } = harness(makeRequest);
    await call(handlers['notify-me'], { message: '<script>x</script>' });
    const body = JSON.parse(makeRequest.mock.calls[1][1].body);
    expect(body.message.body.content).not.toContain('<script>');
  });

  it('is skipped in read-only mode', () => {
    const { handlers, skipped } = harness(makeRequest, true);
    expect(handlers['notify-me']).toBeUndefined();
    expect(skipped).toEqual(['notify-me']);
  });

  it('builds a card with a title, text and optional link', () => {
    const card = workflowPayload('T', 'body', 'https://x.test').attachments[0].content;
    expect(card.body).toHaveLength(2);
    expect(card.actions[0].url).toBe('https://x.test');
  });
});
