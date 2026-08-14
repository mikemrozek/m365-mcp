import { describe, it, expect, vi, beforeEach } from 'vitest';
import { zipSync, strToU8 } from 'fflate';
import { registerFileTools } from '../src/file-tools.js';
import type GraphClient from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), error: vi.fn(), warn: vi.fn() },
}));

vi.mock('../src/usage-log.js', () => ({
  withUsageLog: (_name: string, fn: () => Promise<unknown>) => fn(),
  logToolUsage: vi.fn(),
  default: { info: vi.fn() },
}));

/** Minimal real DOCX so extraction genuinely succeeds rather than being stubbed. */
function makeDocx(text: string): Buffer {
  const xml =
    `<?xml version="1.0"?><w:document xmlns:w="x"><w:body>` +
    `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`;
  return Buffer.from(
    zipSync({ '[Content_Types].xml': strToU8('<Types/>'), 'word/document.xml': strToU8(xml) })
  );
}

interface Handlers {
  [name: string]: (params: Record<string, unknown>) => Promise<{
    content: { type: 'text'; text: string }[];
    isError?: boolean;
  }>;
}

/** Captures registered tools so they can be invoked directly. */
function harness(client: Partial<GraphClient>) {
  const handlers: Handlers = {};
  const server = {
    tool: (name: string, _d: string, _s: unknown, _a: unknown, handler: Handlers[string]) => {
      handlers[name] = handler;
    },
  } as unknown as Parameters<typeof registerFileTools>[0];

  const registered: string[] = [];
  const skipped: string[] = [];
  registerFileTools(server, client as GraphClient, {
    isToolEnabled: () => true,
    readOnly: false,
    push: (n) => registered.push(n),
    fail: (n, e) => {
      throw new Error(`registration failed for ${n}: ${e.message}`);
    },
    skip: (n) => skipped.push(n),
  });
  return { handlers, registered, skipped };
}

const call = async (h: Handlers, name: string, params: Record<string, unknown>) => {
  const res = await h[name](params);
  return { payload: JSON.parse(res.content[0].text), isError: res.isError };
};

describe('get-file routing', () => {
  let makeRequest: ReturnType<typeof vi.fn>;
  let fetchBinary: ReturnType<typeof vi.fn>;
  let putBinary: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    makeRequest = vi.fn();
    fetchBinary = vi.fn();
    putBinary = vi.fn();
  });

  const client = () =>
    ({ makeRequest, fetchBinary, putBinary }) as unknown as Partial<GraphClient>;

  it('delivers a readable document as text, never as bytes', async () => {
    makeRequest.mockResolvedValueOnce({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'report.docx',
      size: 5000,
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    fetchBinary.mockResolvedValueOnce({ buffer: makeDocx('Quarterly revenue was flat'), contentType: '' });

    const { handlers } = harness(client());
    const { payload } = await call(handlers, 'get-file', { messageId: 'm1', attachmentId: 'a1' });

    expect(payload.delivery).toBe('text');
    expect(payload.text).toContain('Quarterly revenue was flat');
    // The decisive property: no base64 payload anywhere in the response.
    expect(JSON.stringify(payload)).not.toMatch(/contentBytes/);
    expect(putBinary).not.toHaveBeenCalled();
  });

  it('falls back to a URL when the file has no text layer', async () => {
    makeRequest.mockResolvedValueOnce({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'logo.png',
      size: 2048,
      contentType: 'image/png',
    });
    fetchBinary.mockResolvedValue({ buffer: Buffer.from('notanimage'), contentType: 'image/png' });
    putBinary.mockResolvedValueOnce({
      id: 'staged-1',
      '@microsoft.graph.downloadUrl': 'https://tenant.example/staged',
    });

    const { handlers } = harness(client());
    const { payload } = await call(handlers, 'get-file', { messageId: 'm1', attachmentId: 'a1' });

    expect(payload.delivery).toBe('url');
    expect(payload.downloadUrl).toBe('https://tenant.example/staged');
    expect(payload.sha256).toMatch(/^[a-f0-9]{64}$/);
  });

  it('skips extraction entirely for very large files', async () => {
    makeRequest.mockResolvedValueOnce({
      name: 'huge.pdf',
      size: 100 * 1024 * 1024,
      file: { mimeType: 'application/pdf' },
    });
    makeRequest.mockResolvedValueOnce({
      id: 'd1',
      '@microsoft.graph.downloadUrl': 'https://tenant.example/huge',
    });

    const { handlers } = harness(client());
    const { payload } = await call(handlers, 'get-file', { itemId: 'd1' });

    expect(payload.delivery).toBe('url');
    // Never pulled 100MB just to discover it was too big.
    expect(fetchBinary).not.toHaveBeenCalled();
  });

  it('rejects a folder with a useful pointer', async () => {
    makeRequest.mockResolvedValueOnce({ name: 'Reports', folder: { childCount: 3 } });
    const { handlers } = harness(client());
    const { payload, isError } = await call(handlers, 'get-file', { itemId: 'f1' });

    expect(isError).toBe(true);
    expect(payload.error).toMatch(/folder/i);
    expect(payload.error).toMatch(/list-folder-files/);
  });

  it('rejects a non-file attachment type', async () => {
    makeRequest.mockResolvedValueOnce({
      '@odata.type': '#microsoft.graph.itemAttachment',
      name: 'embedded',
    });
    const { handlers } = harness(client());
    const { payload, isError } = await call(handlers, 'get-file', {
      messageId: 'm1',
      attachmentId: 'a1',
    });

    expect(isError).toBe(true);
    expect(payload.error).toMatch(/get-mail-attachment/);
  });

  it('explains what to supply when given nothing identifying', async () => {
    const { handlers } = harness(client());
    const { payload, isError } = await call(handlers, 'get-file', {});
    expect(isError).toBe(true);
    expect(payload.error).toMatch(/messageId|itemId/);
  });
});

describe('attach-file size routing', () => {
  let makeRequest: ReturnType<typeof vi.fn>;
  let fetchBinary: ReturnType<typeof vi.fn>;
  let uploadViaSession: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    makeRequest = vi.fn().mockResolvedValue({});
    fetchBinary = vi.fn();
    uploadViaSession = vi.fn().mockResolvedValue({});
  });

  const client = () =>
    ({ makeRequest, fetchBinary, uploadViaSession }) as unknown as Partial<GraphClient>;

  it('attaches a small file directly', async () => {
    makeRequest.mockResolvedValueOnce({ name: 'note.txt', size: 100 });
    fetchBinary.mockResolvedValueOnce({ buffer: Buffer.alloc(100), contentType: 'text/plain' });

    const { handlers } = harness(client());
    const { payload } = await call(handlers, 'attach-file', {
      draftMessageId: 'draft1',
      itemId: 'i1',
    });

    expect(payload.route).toBe('direct');
    expect(uploadViaSession).not.toHaveBeenCalled();
  });

  it('uses an upload session at the 3 MB boundary — the actual defect', async () => {
    const big = Buffer.alloc(4 * 1024 * 1024);
    makeRequest.mockResolvedValueOnce({ name: 'deck.pdf', size: big.byteLength });
    fetchBinary.mockResolvedValueOnce({ buffer: big, contentType: 'application/pdf' });

    const { handlers } = harness(client());
    const { payload } = await call(handlers, 'attach-file', {
      draftMessageId: 'draft1',
      itemId: 'i1',
    });

    expect(payload.route).toBe('uploadSession');
    expect(uploadViaSession).toHaveBeenCalledOnce();
    const [endpoint, body] = uploadViaSession.mock.calls[0];
    expect(endpoint).toContain('createUploadSession');
    expect(body.AttachmentItem).toMatchObject({ attachmentType: 'file', size: big.byteLength });
    // A direct POST would have been the bug; make sure it did not happen.
    const posted = makeRequest.mock.calls.some((c) => c[1]?.method === 'POST');
    expect(posted).toBe(false);
  });

  it('routes just under the boundary directly and just over it via session', async () => {
    const limit = 3 * 1024 * 1024 - 64 * 1024;

    makeRequest.mockResolvedValueOnce({ name: 'a', size: limit - 1 });
    fetchBinary.mockResolvedValueOnce({ buffer: Buffer.alloc(limit - 1), contentType: 'x' });
    let h = harness(client()).handlers;
    expect((await call(h, 'attach-file', { draftMessageId: 'd', itemId: 'i' })).payload.route).toBe(
      'direct'
    );

    makeRequest.mockResolvedValueOnce({ name: 'b', size: limit + 1 });
    fetchBinary.mockResolvedValueOnce({ buffer: Buffer.alloc(limit + 1), contentType: 'x' });
    h = harness(client()).handlers;
    expect((await call(h, 'attach-file', { draftMessageId: 'd', itemId: 'i' })).payload.route).toBe(
      'uploadSession'
    );
  });

  it('honours a filename override', async () => {
    makeRequest.mockResolvedValueOnce({ name: 'original.pdf', size: 50 });
    fetchBinary.mockResolvedValueOnce({ buffer: Buffer.alloc(50), contentType: 'application/pdf' });

    const { handlers } = harness(client());
    const { payload } = await call(handlers, 'attach-file', {
      draftMessageId: 'd',
      itemId: 'i',
      name: 'Invoice March.pdf',
    });
    expect(payload.name).toBe('Invoice March.pdf');
  });
});

describe('read-only mode', () => {
  it('skips both tools, since each writes', () => {
    const handlers: Handlers = {};
    const server = {
      tool: (name: string, _d: string, _s: unknown, _a: unknown, fn: Handlers[string]) => {
        handlers[name] = fn;
      },
    } as unknown as Parameters<typeof registerFileTools>[0];

    const skipped: string[] = [];
    registerFileTools(server, {} as GraphClient, {
      isToolEnabled: () => true,
      readOnly: true,
      push: () => {},
      fail: () => {},
      skip: (n) => skipped.push(n),
    });

    expect(skipped.sort()).toEqual(['attach-file', 'get-file']);
    expect(Object.keys(handlers)).toHaveLength(0);
  });
});
