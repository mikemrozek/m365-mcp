import { describe, it, expect, vi, beforeEach, afterAll } from 'vitest';
import os from 'node:os';
import path from 'node:path';

/**
 * Covers the failure detail added in tsq.24.
 *
 * The usage record used to say only that a call failed. On 2026-09-18, asked what
 * the week's 29 errors were, the only way to answer was to reproduce each shape by
 * hand: Graph's error text goes to a file inside the container and dies with it.
 *
 * So the record now carries a status and a code — and the second half of these
 * tests exists because that is exactly the change that could leak content. Graph
 * echoes the request in its error message: the terms of a failed search, the
 * recipients of a failed send. Only the status and a symbolic code may be written,
 * which makes "never the message" the property worth asserting, in the same spirit
 * as the SEC-2026-001 guard in log-redaction.test.ts.
 */

const { infoSpy } = vi.hoisted(() => ({ infoSpy: vi.fn() }));

vi.mock('winston', () => {
  class Transport {}
  return {
    default: {
      createLogger: () => ({ info: infoSpy }),
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
  getRequestActor: () => ({
    oid: 'oid-1',
    upn: 'someone@townsquaremedia.com',
    tid: 'tid-1',
  }),
}));

const tmpLogDir = path.join(os.tmpdir(), `m365-usage-log-test-${process.pid}`);
const prevLogDir = process.env.MS365_MCP_USAGE_LOG_DIR;
const prevUsageLog = process.env.USAGE_LOG;

async function loadUsageLog() {
  vi.resetModules();
  process.env.MS365_MCP_USAGE_LOG_DIR = tmpLogDir;
  delete process.env.USAGE_LOG;
  return import('../src/usage-log.js');
}

/** The object handed to winston for the most recent record. */
function lastRecord(): Record<string, unknown> {
  const call = infoSpy.mock.calls.at(-1);
  return call?.[1] as Record<string, unknown>;
}

/** A Graph failure exactly as graph-client formats it, message and all. */
function graphErrorText(code: string, message: string): string {
  return `Microsoft Graph API error: 400 Bad Request - ${JSON.stringify({
    error: { code, message, innerError: { 'request-id': 'abc-123' } },
  })}`;
}

beforeEach(() => {
  infoSpy.mockClear();
});

afterAll(() => {
  if (prevLogDir === undefined) delete process.env.MS365_MCP_USAGE_LOG_DIR;
  else process.env.MS365_MCP_USAGE_LOG_DIR = prevLogDir;
  if (prevUsageLog === undefined) delete process.env.USAGE_LOG;
  else process.env.USAGE_LOG = prevUsageLog;
});

describe('describeFailure', () => {
  it('reads the status and code out of a Graph error', async () => {
    const { describeFailure } = await loadUsageLog();
    const detail = describeFailure(
      new Error(graphErrorText('BadRequest', "Syntax error: character ':' is not valid"))
    );
    expect(detail).toEqual({ status: 400, code: 'BadRequest' });
  });

  it('reads the code of the filter/search conflict', async () => {
    const { describeFailure } = await loadUsageLog();
    const detail = describeFailure(
      new Error(
        `Microsoft Graph API error: 400 Bad Request - ${JSON.stringify({
          error: {
            code: 'SearchWithFilter',
            message: "The query parameter '$filter' is not supported with '$search'.",
          },
        })}`
      )
    );
    expect(detail).toEqual({ status: 400, code: 'SearchWithFilter' });
  });

  // Our own tools refuse with a flat slug rather than Graph's nested shape, and carry
  // no HTTP status because no request was made.
  it('reads our own refusal slug off a tool result', async () => {
    const { describeFailure } = await loadUsageLog();
    const detail = describeFailure({
      isError: true,
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            error: 'invalid_search',
            tool: 'list-mail-messages',
            message: 'The $search parameter has no searchable text.',
          }),
        },
      ],
    });
    expect(detail).toEqual({ code: 'invalid_search' });
  });

  it('says nothing rather than guessing when it cannot tell', async () => {
    const { describeFailure } = await loadUsageLog();
    expect(describeFailure(new Error('socket hang up'))).toEqual({});
    expect(describeFailure(undefined)).toEqual({});
    expect(describeFailure({ content: [] })).toEqual({});
  });

  it('reduces a code to a bare token so prose cannot ride out in it', async () => {
    const { describeFailure } = await loadUsageLog();
    const detail = describeFailure({
      content: [{ type: 'text', text: JSON.stringify({ error: 'not a code "at all", really' }) }],
    });
    expect(detail.code).toBeDefined();
    expect(detail.code).not.toMatch(/[\s"]/);
    expect(detail.code!.length).toBeLessThanOrEqual(64);
  });
});

describe('the usage record', () => {
  it('carries status and code on a failure', async () => {
    const { logToolUsage, describeFailure } = await loadUsageLog();
    logToolUsage(
      'list-mail-messages',
      'error',
      describeFailure(new Error(graphErrorText('BadRequest', 'Syntax error')))
    );

    expect(lastRecord()).toMatchObject({
      type: 'm365-usage',
      tool: 'list-mail-messages',
      outcome: 'error',
      status: 400,
      code: 'BadRequest',
      upn: 'someone@townsquaremedia.com',
    });
  });

  it('leaves both absent on success', async () => {
    const { logToolUsage } = await loadUsageLog();
    logToolUsage('list-mail-messages', 'success');

    const record = lastRecord();
    expect(record.outcome).toBe('success');
    expect(record.status).toBeUndefined();
    expect(record.code).toBeUndefined();
  });

  // The whole point of taking only two fields: Graph's message quotes the request back.
  it('never carries the Graph message, which quotes the request back', async () => {
    const { logToolUsage, describeFailure } = await loadUsageLog();
    const secretTerms = 'subject:Project Falcon salary review';
    const error = new Error(
      graphErrorText(
        'BadRequest',
        `Syntax error: character ':' is not valid at position 7 in '${secretTerms}'.`
      )
    );

    logToolUsage('list-mail-messages', 'error', describeFailure(error));

    const serialized = JSON.stringify(lastRecord());
    expect(serialized).not.toContain('Falcon');
    expect(serialized).not.toContain('salary');
    expect(serialized).not.toContain('Syntax error');
    expect(serialized).not.toContain('request-id');
    // The diagnostic half still survives.
    expect(serialized).toContain('BadRequest');
    expect(serialized).toContain('400');
  });
});

describe('withUsageLog', () => {
  it('records the detail of a returned error result', async () => {
    const { withUsageLog } = await loadUsageLog();
    await withUsageLog('put-file', async () => ({
      isError: true,
      content: [{ type: 'text', text: JSON.stringify({ error: 'upload_failed' }) }],
    }));

    expect(lastRecord()).toMatchObject({
      tool: 'put-file',
      outcome: 'error',
      code: 'upload_failed',
    });
  });

  it('records the detail of a thrown error', async () => {
    const { withUsageLog } = await loadUsageLog();
    await expect(
      withUsageLog('get-file', async () => {
        throw new Error(graphErrorText('ErrorItemNotFound', 'The specified object was not found.'));
      })
    ).rejects.toThrow();

    expect(lastRecord()).toMatchObject({
      tool: 'get-file',
      outcome: 'error',
      status: 400,
      code: 'ErrorItemNotFound',
    });
  });

  it('adds nothing to a successful call', async () => {
    const { withUsageLog } = await loadUsageLog();
    await withUsageLog('get-file', async () => ({ content: [] }));

    const record = lastRecord();
    expect(record.outcome).toBe('success');
    expect(record.status).toBeUndefined();
    expect(record.code).toBeUndefined();
  });
});
