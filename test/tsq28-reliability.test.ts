import { beforeEach, afterEach, describe, expect, it, vi } from 'vitest';
import { z } from 'zod';

/**
 * tsq.28 items 3–8 and 10, through the real executor: each test drives a
 * registered tool handler and asserts on what reached Graph (or that nothing
 * did). Evidence for every case is the week to 2026-10-09 — see the comments
 * beside each fix in src/graph-tools.ts.
 */

const { usageSpy } = vi.hoisted(() => ({ usageSpy: vi.fn() }));

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn() },
}));

// The usage log's winston sink, captured so the m365-usage and m365-limit
// records can be asserted on directly.
vi.mock('winston', () => {
  class Transport {}
  return {
    default: {
      createLogger: () => ({ info: usageSpy }),
      format: {
        combine: (...parts: unknown[]) => parts,
        timestamp: () => ({}),
        json: () => ({}),
      },
      transports: { File: Transport, Console: Transport },
    },
  };
});

vi.mock('../src/request-context.js', () => ({
  getRequestTokens: () => undefined,
  getRequestActor: () => ({ oid: 'oid-1', upn: 'someone@townsquaremedia.com', tid: 'tid-1' }),
}));

vi.mock('../src/generated/client.js', () => {
  const q = (name: string, schema: z.ZodTypeAny = z.string()) => ({
    name,
    type: 'Query',
    schema: schema.optional(),
  });
  // Without '$': the executor matches an OData param to its definition by the bare name.
  const listParams = [q('top', z.number()), q('filter'), q('orderby'), q('search')];
  const body = { name: 'body', type: 'Body', schema: z.any() };
  return {
    api: {
      endpoints: [
        { alias: 'delete-mail-message', method: 'delete', path: '/me/messages/:messageId' },
        {
          alias: 'delete-mail-attachment',
          method: 'delete',
          path: '/me/messages/:messageId/attachments/:attachmentId',
        },
        {
          alias: 'get-sharepoint-site-by-path',
          method: 'get',
          path: `/sites/:siteId/getByPath(path=':path')`,
        },
        { alias: 'send-mail', method: 'post', path: '/me/sendMail', parameters: [body] },
        { alias: 'create-draft-email', method: 'post', path: '/me/messages', parameters: [body] },
        {
          alias: 'reply-mail-message',
          method: 'post',
          path: '/me/messages/:messageId/reply',
          parameters: [body],
        },
        { alias: 'create-calendar-event', method: 'post', path: '/me/events', parameters: [body] },
        { alias: 'list-chats', method: 'get', path: '/me/chats', parameters: listParams },
        {
          alias: 'list-chat-messages',
          method: 'get',
          path: '/chats/:chatId/messages',
          parameters: listParams,
        },
        {
          alias: 'list-online-meetings',
          method: 'get',
          path: '/me/onlineMeetings',
          parameters: listParams,
        },
        {
          alias: 'list-mail-rules',
          method: 'get',
          path: '/me/mailFolders/:mailFolderId/messageRules',
          parameters: listParams,
        },
        {
          alias: 'list-mail-messages',
          method: 'get',
          path: '/me/messages',
          parameters: listParams,
        },
        {
          alias: 'send-chat-message',
          method: 'post',
          path: '/chats/:chatId/messages',
          parameters: [body],
        },
      ],
    },
  };
});

const { registerGraphTools, normalizeWriteBody, firstFilterProperty, resolveSiteByPath } =
  await import('../src/graph-tools.js');
const { resetDailyCountsForTests } = await import('../src/usage-log.js');

type Handler = (params: Record<string, unknown>) => Promise<{
  content: { text: string }[];
  isError?: boolean;
}>;

let graphRequest: ReturnType<typeof vi.fn>;
let handlers: Record<string, Handler>;

function setup() {
  graphRequest = vi.fn().mockResolvedValue({
    content: [{ type: 'text', text: JSON.stringify({ value: [] }) }],
  });
  handlers = {};
  const server = {
    tool: (...args: unknown[]) => {
      handlers[args[0] as string] = args[args.length - 1] as Handler;
    },
  };
  registerGraphTools(server as never, { graphRequest } as never, false, undefined, true);
}

const sentPath = (i = 0) => graphRequest.mock.calls[i][0] as string;
const sentBody = (i = 0) => JSON.parse(graphRequest.mock.calls[i][1].body as string);
const payloadOf = (r: { content: { text: string }[] }) => JSON.parse(r.content[0].text);
const graphError = (status: number, statusText: string, code: string) => ({
  content: [
    {
      type: 'text',
      text: JSON.stringify({
        error: `Microsoft Graph API error: ${status} ${statusText} - {"error":{"code":"${code}","message":"x"}}`,
      }),
    },
  ],
  isError: true,
});

beforeEach(() => {
  vi.clearAllMocks();
  resetDailyCountsForTests();
  setup();
});

// ---- Item 3 --------------------------------------------------------------------

describe('delete-mail-message is idempotent', () => {
  it('reports 404 ErrorItemNotFound as already deleted, not as a failure', async () => {
    graphRequest.mockResolvedValueOnce(graphError(404, 'Not Found', 'ErrorItemNotFound'));
    const result = await handlers['delete-mail-message']({ messageId: 'AAMk=' });
    expect(result.isError).toBeFalsy();
    const payload = payloadOf(result);
    expect(payload.success).toBe(true);
    expect(payload.alreadyDeleted).toBe(true);
  });

  it('leaves other delete failures as errors', async () => {
    graphRequest.mockResolvedValueOnce(graphError(403, 'Forbidden', 'ErrorAccessDenied'));
    const result = await handlers['delete-mail-message']({ messageId: 'AAMk=' });
    expect(result.isError).toBe(true);
  });

  it('is narrow to messages — an attachment 404 is still an error', async () => {
    graphRequest.mockResolvedValueOnce(graphError(404, 'Not Found', 'ErrorItemNotFound'));
    const result = await handlers['delete-mail-attachment']({
      messageId: 'm',
      attachmentId: 'a',
    });
    expect(result.isError).toBe(true);
  });
});

// ---- Item 4 --------------------------------------------------------------------

describe('get-sharepoint-site-by-path', () => {
  it('turns a full site URL into /sites/{hostname}:/{path}, keeping the colon literal', async () => {
    await handlers['get-sharepoint-site-by-path']({
      url: 'https://contoso.sharepoint.com/sites/marketing',
    });
    expect(sentPath()).toBe('/sites/contoso.sharepoint.com:/sites/marketing');
  });

  it('accepts hostname plus path', async () => {
    await handlers['get-sharepoint-site-by-path']({
      hostname: 'contoso.sharepoint.com',
      path: 'teams/sales',
    });
    expect(sentPath()).toBe('/sites/contoso.sharepoint.com:/teams/sales');
  });

  it('refuses rather than calling Graph when it cannot build an address', async () => {
    const result = await handlers['get-sharepoint-site-by-path']({ path: '/sites/marketing' });
    expect(result.isError).toBe(true);
    expect(payloadOf(result).error).toBe('invalid_site_address');
    expect(graphRequest).not.toHaveBeenCalled();
  });
});

describe('resolveSiteByPath', () => {
  it('trims a page or library URL back to the site', () => {
    expect(
      resolveSiteByPath({
        url: 'https://contoso.sharepoint.com/sites/marketing/SitePages/Home.aspx',
      }).path
    ).toBe('/sites/contoso.sharepoint.com:/sites/marketing');
    expect(
      resolveSiteByPath({
        url: 'https://contoso.sharepoint.com/sites/marketing/Shared%20Documents/Forms/AllItems.aspx',
      }).path
    ).toBe('/sites/contoso.sharepoint.com:/sites/marketing');
  });

  it('keeps a subsite, and encodes each segment on its own', () => {
    expect(
      resolveSiteByPath({ url: 'https://contoso.sharepoint.com/sites/hr/Team%20Site/Lists/x' }).path
    ).toBe('/sites/contoso.sharepoint.com:/sites/hr/Team%20Site');
  });

  it('strips the sharing-link marker', () => {
    expect(
      resolveSiteByPath({
        url: 'https://contoso.sharepoint.com/:f:/r/sites/ops/Shared%20Documents',
      }).path
    ).toBe('/sites/contoso.sharepoint.com:/sites/ops');
  });

  it('addresses the root site by hostname alone', () => {
    expect(resolveSiteByPath({ url: 'https://contoso.sharepoint.com/' }).path).toBe(
      '/sites/contoso.sharepoint.com'
    );
    expect(resolveSiteByPath({ hostname: 'contoso.sharepoint.com' }).path).toBe(
      '/sites/contoso.sharepoint.com'
    );
  });

  it('accepts a URL in path, and a hostname in the old siteId slot', () => {
    expect(resolveSiteByPath({ path: 'https://contoso.sharepoint.com/sites/a' }).path).toBe(
      '/sites/contoso.sharepoint.com:/sites/a'
    );
    expect(resolveSiteByPath({ siteId: 'contoso.sharepoint.com', path: '/sites/a' }).path).toBe(
      '/sites/contoso.sharepoint.com:/sites/a'
    );
  });

  it('rejects a hostname that could smuggle a path', () => {
    expect(resolveSiteByPath({ hostname: 'contoso.sharepoint.com/../x', path: 'a' }).error).toBe(
      'Not a valid SharePoint hostname: contoso.sharepoint.com/../x'
    );
  });
});

// ---- Item 5 --------------------------------------------------------------------

describe('write-body normalization', () => {
  it('wraps a send-mail body that lost its message wrapper, and fixes string recipients', async () => {
    await handlers['send-mail']({
      body: {
        subject: 'Hi',
        body: '<p>Hello</p>',
        toRecipients: ['a@example.com', { address: 'b@example.com', name: 'B' }],
        ccRecipients: 'c@example.com; d@example.com',
        saveToSentItems: true,
      },
    });
    expect(sentBody()).toEqual({
      message: {
        subject: 'Hi',
        body: { contentType: 'HTML', content: '<p>Hello</p>' },
        toRecipients: [
          { emailAddress: { address: 'a@example.com' } },
          { emailAddress: { address: 'b@example.com', name: 'B' } },
        ],
        ccRecipients: [
          { emailAddress: { address: 'c@example.com' } },
          { emailAddress: { address: 'd@example.com' } },
        ],
      },
      saveToSentItems: true,
    });
  });

  it('leaves a correct send-mail body exactly as sent', () => {
    const body = {
      message: {
        subject: 'Hi',
        body: { contentType: 'Text', content: 'x' },
        toRecipients: [{ emailAddress: { address: 'a@example.com' } }],
      },
      saveToSentItems: false,
    };
    expect(normalizeWriteBody('send-mail', body).body).toBe(body);
  });

  it('does not wrap a draft — its body is the message', () => {
    const { body } = normalizeWriteBody('create-draft-email', {
      subject: 'S',
      toRecipients: ['a@example.com'],
    });
    expect(body).toEqual({
      subject: 'S',
      toRecipients: [{ emailAddress: { address: 'a@example.com' } }],
    });
  });

  it('fixes recipients inside a reply message', async () => {
    await handlers['reply-mail-message']({
      messageId: 'm1',
      body: { comment: 'ok', message: { toRecipients: ['a@example.com'] } },
    });
    expect(sentBody().message.toRecipients).toEqual([
      { emailAddress: { address: 'a@example.com' } },
    ]);
    expect(sentBody().comment).toBe('ok');
  });

  it("gives plain-address attendees type 'required' and leaves correct ones alone", () => {
    const correct = { emailAddress: { address: 'x@example.com' }, type: 'optional' };
    const { body } = normalizeWriteBody('create-calendar-event', {
      subject: 'S',
      start: { dateTime: '2026-10-14T15:00:00', timeZone: 'Eastern Standard Time' },
      end: { dateTime: '2026-10-14T15:30:00', timeZone: 'Eastern Standard Time' },
      attendees: ['a@example.com', { address: 'b@example.com' }, correct],
    });
    expect((body as { attendees: unknown[] }).attendees).toEqual([
      { emailAddress: { address: 'a@example.com' }, type: 'required' },
      { emailAddress: { address: 'b@example.com' }, type: 'required' },
      correct,
    ]);
  });

  it('refuses a bare-string start time instead of guessing a time zone', async () => {
    const result = await handlers['create-calendar-event']({
      body: { subject: 'S', start: '2026-10-14T15:00:00', end: '2026-10-14T15:30:00' },
    });
    expect(result.isError).toBe(true);
    const payload = payloadOf(result);
    expect(payload.error).toBe('invalid_event_time');
    expect(payload.message).toContain('timeZone');
    expect(graphRequest).not.toHaveBeenCalled();
  });

  it('ignores tools it does not know', () => {
    const body = { toRecipients: ['a@example.com'] };
    expect(normalizeWriteBody('create-chat', body).body).toBe(body);
  });
});

// ---- Item 6 --------------------------------------------------------------------

describe('Teams chat query options', () => {
  it('list-chats drops $orderby and clamps $top to 50', async () => {
    await handlers['list-chats']({ orderby: 'lastUpdatedDateTime desc', top: 200 });
    expect(sentPath()).not.toContain('orderby');
    expect(sentPath()).toContain('$top=50');
  });

  it('list-chat-messages keeps a supported $orderby and drops any other', async () => {
    await handlers['list-chat-messages']({ chatId: 'c', orderby: 'createdDateTime desc', top: 75 });
    expect(decodeURIComponent(sentPath())).toContain('$orderby=createdDateTime desc');
    expect(sentPath()).toContain('$top=50');

    await handlers['list-chat-messages']({ chatId: 'c', orderby: 'createdDateTime asc' });
    expect(sentPath(1)).not.toContain('orderby');
  });
});

// ---- Item 7 --------------------------------------------------------------------

describe('list-online-meetings', () => {
  it('refuses without a join-URL filter and says where to get one', async () => {
    const result = await handlers['list-online-meetings']({ top: 10 });
    expect(result.isError).toBe(true);
    const payload = payloadOf(result);
    expect(payload.error).toBe('join_url_required');
    expect(payload.message).toContain('onlineMeeting.joinUrl');
    expect(graphRequest).not.toHaveBeenCalled();
  });

  it('passes a JoinWebUrl filter through', async () => {
    await handlers['list-online-meetings']({
      filter: "JoinWebUrl eq 'https://teams.microsoft.com/l/meetup-join/x'",
    });
    expect(graphRequest).toHaveBeenCalledTimes(1);
  });
});

// ---- Item 8 --------------------------------------------------------------------

describe('unsupported query options on mail', () => {
  it('list-mail-rules drops $filter, $search and $orderby', async () => {
    await handlers['list-mail-rules']({
      mailFolderId: 'inbox',
      filter: "displayName eq 'x'",
      orderby: 'sequence',
    });
    expect(sentPath()).toBe('/me/mailFolders/inbox/messageRules');
  });

  it('drops an $orderby that would trip InefficientFilter and sorts the page instead', async () => {
    graphRequest.mockResolvedValueOnce({
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            value: [
              { id: '1', receivedDateTime: '2026-10-01T00:00:00Z' },
              { id: '2', receivedDateTime: '2026-10-03T00:00:00Z' },
            ],
          }),
        },
      ],
    });
    const result = await handlers['list-mail-messages']({
      filter: 'isRead eq false',
      orderby: 'receivedDateTime desc',
    });
    expect(sentPath()).not.toContain('orderby');
    const payload = payloadOf(result);
    expect(payload.value.map((m: { id: string }) => m.id)).toEqual(['2', '1']);
    expect(payload._tsq_clientSorted.field).toBe('receivedDateTime');
  });

  it('keeps $orderby when it leads the filter, as Graph requires', async () => {
    await handlers['list-mail-messages']({
      filter: "receivedDateTime ge 2026-10-01T00:00:00Z and from/emailAddress/address eq 'a@x.com'",
      orderby: 'receivedDateTime desc',
    });
    expect(sentPath()).toContain('orderby');
  });
});

describe('firstFilterProperty', () => {
  it.each([
    ['receivedDateTime ge 2026-01-01', 'receivedDateTime'],
    ['(isRead eq false) and importance eq high', 'isRead'],
    ["contains(subject, 'x')", 'subject'],
    ['not (isRead eq true)', 'isRead'],
    ["from/emailAddress/address eq 'a'", 'from/emailAddress/address'],
    ["toRecipients/any(r: r/emailAddress/address eq 'a')", 'toRecipients'],
  ])('%s → %s', (filter, prop) => {
    expect(firstFilterProperty(filter)).toBe(prop);
  });

  it('returns undefined for an expression it cannot read', () => {
    expect(firstFilterProperty("'x' eq subject")).toBeUndefined();
  });
});

// ---- Item 10 -------------------------------------------------------------------

describe('per-user daily ceilings', () => {
  const env = { ...process.env };
  afterEach(() => {
    process.env = { ...env };
  });

  const usageRecords = (type: string) =>
    usageSpy.mock.calls.map((c) => c[1] as Record<string, unknown>).filter((r) => r.type === type);

  it('is off by default', async () => {
    delete process.env.MS365_MCP_DAILY_LIMIT_DELETE;
    for (let i = 0; i < 5; i++) {
      const r = await handlers['delete-mail-message']({ messageId: `m${i}` });
      expect(r.isError).toBeFalsy();
    }
    expect(usageRecords('m365-limit')).toHaveLength(0);
  });

  it('refuses at the ceiling with daily_limit_exceeded, and logs the 80% and 100% crossings', async () => {
    process.env.MS365_MCP_DAILY_LIMIT_SEND = '5';
    for (let i = 0; i < 5; i++) {
      const r = await handlers['send-chat-message']({
        chatId: 'c',
        body: { body: { content: 'x' } },
      });
      expect(r.isError).toBeFalsy();
    }
    const refused = await handlers['send-mail']({ body: { message: { subject: 's' } } });
    expect(refused.isError).toBe(true);
    expect(payloadOf(refused).error).toBe('daily_limit_exceeded');
    expect(graphRequest).toHaveBeenCalledTimes(5);

    const limits = usageRecords('m365-limit');
    expect(limits.map((r) => r.count)).toEqual([4, 5]);
    expect(limits[0]).toMatchObject({ oid: 'oid-1', limit: 5, kind: 'send' });

    const usage = usageRecords('m365-usage').at(-1);
    expect(usage).toMatchObject({
      tool: 'send-mail',
      outcome: 'error',
      code: 'daily_limit_exceeded',
    });
  });

  it('does not count a failed action against the day', async () => {
    process.env.MS365_MCP_DAILY_LIMIT_DELETE = '1';
    graphRequest.mockResolvedValueOnce(graphError(500, 'Internal Server Error', 'Error'));
    expect((await handlers['delete-mail-message']({ messageId: 'a' })).isError).toBe(true);
    expect((await handlers['delete-mail-message']({ messageId: 'b' })).isError).toBeFalsy();
    const third = await handlers['delete-mail-message']({ messageId: 'c' });
    expect(payloadOf(third).error).toBe('daily_limit_exceeded');
  });

  it('leaves tools outside the list alone', async () => {
    process.env.MS365_MCP_DAILY_LIMIT_SEND = '1';
    process.env.MS365_MCP_DAILY_LIMIT_DELETE = '1';
    for (let i = 0; i < 3; i++) {
      expect(
        (await handlers['create-draft-email']({ body: { subject: 's' } })).isError
      ).toBeFalsy();
    }
  });
});
