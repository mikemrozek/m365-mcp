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

  it('defaults to a LINK, not text — the server does not interpret the file', async () => {
    makeRequest.mockResolvedValueOnce({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'report.docx',
      size: 5000,
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    fetchBinary.mockResolvedValue({ buffer: makeDocx('Quarterly revenue was flat'), contentType: '' });
    putBinary.mockResolvedValueOnce({
      id: 'staged-1',
      '@microsoft.graph.downloadUrl': 'https://tenant.example/staged',
    });

    const { handlers } = harness(client());
    const { payload } = await call(handlers, 'get-file', { messageId: 'm1', attachmentId: 'a1' });

    // Scott's requirement: hand over the file, let the agent decide what to do.
    expect(payload.delivery).toBe('url');
    expect(payload.downloadUrl).toBe('https://tenant.example/staged');
    expect(payload.text).toBeUndefined();
  });

  it('extracts text ONLY when explicitly asked', async () => {
    makeRequest.mockResolvedValueOnce({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'report.docx',
      size: 5000,
      contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    fetchBinary.mockResolvedValueOnce({ buffer: makeDocx('Quarterly revenue was flat'), contentType: '' });

    const { handlers } = harness(client());
    const { payload } = await call(handlers, 'get-file', {
      messageId: 'm1',
      attachmentId: 'a1',
      as: 'text',
    });

    expect(payload.delivery).toBe('text');
    expect(payload.text).toContain('Quarterly revenue was flat');
    expect(JSON.stringify(payload)).not.toMatch(/contentBytes/);
  });

  it('never decodes a binary container as text just because its mime type says xml', async () => {
    // Regression: .docx reports application/vnd.openxmlformats-… which CONTAINS
    // "xml" but is a ZIP. A substring match here returned garbage as UTF-8.
    makeRequest.mockResolvedValueOnce({
      name: 'small.docx',
      size: 4000,
      file: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
    });
    fetchBinary.mockResolvedValue({ buffer: makeDocx('hello'), contentType: '' });
    makeRequest.mockResolvedValueOnce({
      id: 'd1',
      '@microsoft.graph.downloadUrl': 'https://tenant.example/small',
    });

    const { handlers } = harness(client());
    const { payload } = await call(handlers, 'get-file', { itemId: 'd1' });

    expect(payload.delivery).toBe('url');
    expect(payload.content).toBeUndefined();
  });

  it('does not $select the download URL away, and never claims url without one', async () => {
    // Regression, reported from production 2026-08-17: @microsoft.graph.downloadUrl
    // is an OData ANNOTATION, not a selectable property. $select-ing it suppresses
    // it, so the call succeeded and returned delivery:'url' with no url in the
    // payload — JSON.stringify drops undefined keys, so the field vanished silently.
    makeRequest.mockResolvedValueOnce({ name: 'doc.pdf', size: 900000, file: { mimeType: 'application/pdf' } });
    makeRequest.mockResolvedValueOnce({
      id: 'd1',
      '@microsoft.graph.downloadUrl': 'https://tenant.example/real',
    });

    const { handlers } = harness(client());
    const { payload } = await call(handlers, 'get-file', { itemId: 'd1' });

    expect(payload.delivery).toBe('url');
    expect(payload.downloadUrl).toBe('https://tenant.example/real');

    // The item fetch must not carry a $select, or Graph withholds the annotation.
    const itemFetch = makeRequest.mock.calls[1][0] as string;
    expect(itemFetch).not.toMatch(/\$select/);
  });

  it('fails loudly rather than returning a url delivery with no url', async () => {
    makeRequest.mockResolvedValueOnce({ name: 'weird.dat', size: 900000, file: {} });
    makeRequest.mockResolvedValueOnce({ id: 'd1' }); // Graph gave us nothing usable

    const { handlers } = harness(client());
    await expect(call(handlers, 'get-file', { itemId: 'd1' })).rejects.toThrow(/no download URL/i);
  });

  it('returns a genuinely tiny text file inline, since a link would cost more', async () => {
    makeRequest.mockResolvedValueOnce({ name: 'notes.txt', size: 12, file: { mimeType: 'text/plain' } });
    fetchBinary.mockResolvedValueOnce({ buffer: Buffer.from('hello world!'), contentType: 'text/plain' });

    const { handlers } = harness(client());
    const { payload } = await call(handlers, 'get-file', { itemId: 'd1' });

    expect(payload.delivery).toBe('inline');
    expect(payload.content).toBe('hello world!');
  });

  it('gives a link, not an error, when text was asked for but none exists', async () => {
    makeRequest.mockResolvedValueOnce({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'logo.png',
      size: 40000,
      contentType: 'image/png',
    });
    fetchBinary.mockResolvedValue({ buffer: Buffer.from('notanimage'), contentType: 'image/png' });
    putBinary.mockResolvedValueOnce({
      id: 'staged-2',
      '@microsoft.graph.downloadUrl': 'https://tenant.example/png',
    });

    const { handlers } = harness(client());
    const { payload, isError } = await call(handlers, 'get-file', {
      messageId: 'm1',
      attachmentId: 'a1',
      as: 'text',
    });

    expect(isError).toBeUndefined();
    expect(payload.delivery).toBe('url');
    expect(payload.note).toMatch(/no text|none to extract/i);
  });

  it('returns a link for a file with no text layer', async () => {
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

  it('never pulls bytes for a large file it only needs to link to', async () => {
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

describe('url response shape (consistency across branches)', () => {
  let makeRequest: ReturnType<typeof vi.fn>;
  let fetchBinary: ReturnType<typeof vi.fn>;
  let putBinary: ReturnType<typeof vi.fn>;

  beforeEach(() => {
    makeRequest = vi.fn();
    fetchBinary = vi.fn();
    putBinary = vi.fn();
  });
  const client = () => ({ makeRequest, fetchBinary, putBinary }) as unknown as Partial<GraphClient>;

  it('both branches report `bytes`, so a caller can verify either', async () => {
    // Drive branch.
    makeRequest.mockResolvedValueOnce({ name: 'a.pdf', size: 900000, file: { mimeType: 'application/pdf' } });
    makeRequest.mockResolvedValueOnce({ id: 'd1', '@microsoft.graph.downloadUrl': 'https://t/x' });
    let h = harness(client()).handlers;
    const drive = (await call(h, 'get-file', { itemId: 'd1' })).payload;
    expect(drive.bytes).toBe(900000);
    expect(drive.integrity).toMatch(/pass-through/i);
    expect(drive.sha256).toBeUndefined();

    // Staged mail branch.
    makeRequest = vi.fn();
    fetchBinary = vi.fn();
    putBinary = vi.fn();
    makeRequest.mockResolvedValueOnce({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'b.png', size: 5165, contentType: 'image/png',
    });
    makeRequest.mockResolvedValueOnce({ value: [] });          // prune listing
    fetchBinary.mockResolvedValue({ buffer: Buffer.alloc(5000), contentType: 'image/png' });
    putBinary.mockResolvedValueOnce({ id: 's1', '@microsoft.graph.downloadUrl': 'https://t/y' });
    h = harness(client()).handlers;
    const mail = (await call(h, 'get-file', { messageId: 'm', attachmentId: 'a' })).payload;

    expect(mail.bytes).toBe(5000);
    expect(mail.sha256).toMatch(/^[a-f0-9]{64}$/);
    expect(mail.integrity).toMatch(/sha256/i);
  });

  it("surfaces Graph's inflated size ONLY when it disagrees with the real file", async () => {
    makeRequest.mockResolvedValueOnce({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'b.png', size: 5165, contentType: 'image/png',   // MIME-inflated
    });
    makeRequest.mockResolvedValueOnce({ value: [] });
    fetchBinary.mockResolvedValue({ buffer: Buffer.alloc(5000), contentType: 'image/png' });
    putBinary.mockResolvedValueOnce({ id: 's1', '@microsoft.graph.downloadUrl': 'https://t/y' });

    const { handlers } = harness(client());
    const { payload } = await call(handlers, 'get-file', { messageId: 'm', attachmentId: 'a' });

    // 165-byte MIME overhead must be explained, not left to look like corruption.
    expect(payload.graphReportedSize).toBe(5165);
    expect(payload.sizeNote).toMatch(/MIME/i);
  });

  it('stays quiet about size when the two agree', async () => {
    makeRequest.mockResolvedValueOnce({ name: 'a.pdf', size: 900000, file: { mimeType: 'application/pdf' } });
    makeRequest.mockResolvedValueOnce({ id: 'd1', '@microsoft.graph.downloadUrl': 'https://t/x' });
    const { handlers } = harness(client());
    const { payload } = await call(handlers, 'get-file', { itemId: 'd1' });
    expect(payload.graphReportedSize).toBeUndefined();
    expect(payload.sizeNote).toBeUndefined();
  });

  it('prunes staged copies older than 24h before adding another', async () => {
    const old = new Date(Date.now() - 48 * 3600 * 1000).toISOString();
    const fresh = new Date().toISOString();
    makeRequest.mockResolvedValueOnce({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'c.png', size: 100, contentType: 'image/png',
    });
    makeRequest.mockResolvedValueOnce({
      value: [
        { id: 'stale1', name: 'old.pdf', createdDateTime: old },
        { id: 'keep1', name: 'new.pdf', createdDateTime: fresh },
      ],
    });
    makeRequest.mockResolvedValue({});                        // the DELETE
    fetchBinary.mockResolvedValue({ buffer: Buffer.alloc(100), contentType: 'image/png' });
    putBinary.mockResolvedValueOnce({ id: 's2', '@microsoft.graph.downloadUrl': 'https://t/z' });

    const { handlers } = harness(client());
    await call(handlers, 'get-file', { messageId: 'm', attachmentId: 'a' });

    const deletes = makeRequest.mock.calls.filter((c) => c[1]?.method === 'DELETE');
    expect(deletes).toHaveLength(1);
    expect(deletes[0][0]).toContain('stale1');   // the fresh one survives
  });

  it('still returns the file if pruning fails', async () => {
    makeRequest.mockResolvedValueOnce({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'd.png', size: 100, contentType: 'image/png',
    });
    makeRequest.mockRejectedValueOnce(new Error('folder not found'));   // prune blows up
    fetchBinary.mockResolvedValue({ buffer: Buffer.alloc(100), contentType: 'image/png' });
    putBinary.mockResolvedValueOnce({ id: 's3', '@microsoft.graph.downloadUrl': 'https://t/w' });

    const { handlers } = harness(client());
    const { payload } = await call(handlers, 'get-file', { messageId: 'm', attachmentId: 'a' });
    expect(payload.downloadUrl).toBe('https://t/w');   // housekeeping must not break the request
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
    // Size cannot verify integrity (MIME overhead), so a hash must be returned.
    expect(payload.sha256).toMatch(/^[a-f0-9]{64}$/);
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
    expect(payload.sha256).toMatch(/^[a-f0-9]{64}$/);
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
