import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

/**
 * tsq.28 items 1–2: one retry in the shared Graph request path.
 *
 * The property worth asserting beyond "it retries" is the one that must never
 * break: a POST that failed with a 5xx is NOT replayed, because the gateway may
 * have timed out after Exchange already sent the email.
 */

vi.mock('../src/logger.js', () => ({
  default: { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn() },
}));

vi.mock('../src/cloud-config.js', () => ({
  getCloudEndpoints: () => ({
    graphApi: 'https://graph.microsoft.com',
    authority: 'https://login.microsoftonline.com',
  }),
}));

vi.mock('@toon-format/toon', () => ({ encode: (data: unknown) => JSON.stringify(data) }));

const {
  default: GraphClient,
  retryPolicy,
  parseRetryAfterMs,
} = await import('../src/graph-client.js');

const authManager = { getToken: vi.fn().mockResolvedValue('token') };
const secrets = { clientId: 'c', tenantId: 't', cloudType: 'global' as const };

const ok = (body: unknown = { value: [] }) => new Response(JSON.stringify(body), { status: 200 });
const fail = (status: number, headers: Record<string, string> = {}) =>
  new Response(JSON.stringify({ error: { code: `E${status}` } }), { status, headers });

describe('Graph request retry (tsq.28)', () => {
  let fetchSpy: ReturnType<typeof vi.spyOn>;
  let sleepSpy: ReturnType<typeof vi.spyOn>;
  let client: InstanceType<typeof GraphClient>;

  beforeEach(() => {
    fetchSpy = vi.spyOn(globalThis, 'fetch');
    sleepSpy = vi.spyOn(retryPolicy, 'sleep').mockResolvedValue(undefined);
    client = new GraphClient(authManager as never, secrets as never);
  });

  afterEach(() => {
    fetchSpy.mockRestore();
    sleepSpy.mockRestore();
  });

  it('waits Retry-After and retries a 429 once, for any method — throttled means not processed', async () => {
    fetchSpy.mockResolvedValueOnce(fail(429, { 'Retry-After': '3' })).mockResolvedValueOnce(ok());
    const result = await client.makeRequest('/me/sendMail', { method: 'POST', body: '{}' });
    expect(result).toEqual({ value: [] });
    expect(fetchSpy).toHaveBeenCalledTimes(2);
    expect(sleepSpy).toHaveBeenCalledWith(3000);
  });

  it('defaults to 2s when a 429 carries no Retry-After', async () => {
    fetchSpy.mockResolvedValueOnce(fail(429)).mockResolvedValueOnce(ok());
    await client.makeRequest('/me/mailFolders');
    expect(sleepSpy).toHaveBeenCalledWith(2000);
  });

  it('does not retry when Retry-After is beyond the 30s cap', async () => {
    fetchSpy.mockResolvedValueOnce(fail(429, { 'Retry-After': '120' }));
    await expect(client.makeRequest('/me/messages')).rejects.toThrow(/429/);
    expect(fetchSpy).toHaveBeenCalledTimes(1);
    expect(sleepSpy).not.toHaveBeenCalled();
  });

  it('retries only once: a second 429 is returned to the caller', async () => {
    fetchSpy.mockResolvedValueOnce(fail(429)).mockResolvedValueOnce(fail(429));
    await expect(client.makeRequest('/me/messages')).rejects.toThrow(/429/);
    expect(fetchSpy).toHaveBeenCalledTimes(2);
  });

  it.each([500, 502, 503, 504])('retries a GET once after %i', async (status) => {
    fetchSpy.mockResolvedValueOnce(fail(status)).mockResolvedValueOnce(ok());
    await expect(client.makeRequest('/me/mailFolders/inbox/messages')).resolves.toEqual({
      value: [],
    });
    expect(fetchSpy).toHaveBeenCalledTimes(2);
  });

  it('retries a DELETE once after a 503', async () => {
    fetchSpy
      .mockResolvedValueOnce(fail(503))
      .mockResolvedValueOnce(new Response(null, { status: 204 }));
    await expect(client.makeRequest('/me/messages/abc', { method: 'DELETE' })).resolves.toEqual({
      message: 'OK!',
    });
    expect(fetchSpy).toHaveBeenCalledTimes(2);
  });

  it.each(['POST', 'PATCH', 'PUT'])(
    'never replays a %s after a 5xx — the request may already have been processed',
    async (method) => {
      fetchSpy.mockResolvedValueOnce(fail(504));
      await expect(client.makeRequest('/me/sendMail', { method, body: '{}' })).rejects.toThrow(
        /504/
      );
      expect(fetchSpy).toHaveBeenCalledTimes(1);
    }
  );

  it('retries a GET once after a network error, but not a POST', async () => {
    fetchSpy.mockRejectedValueOnce(new TypeError('fetch failed')).mockResolvedValueOnce(ok());
    await expect(client.makeRequest('/me/messages')).resolves.toEqual({ value: [] });
    expect(fetchSpy).toHaveBeenCalledTimes(2);

    fetchSpy.mockReset();
    fetchSpy.mockRejectedValueOnce(new TypeError('fetch failed'));
    await expect(
      client.makeRequest('/me/sendMail', { method: 'POST', body: '{}' })
    ).rejects.toThrow('fetch failed');
    expect(fetchSpy).toHaveBeenCalledTimes(1);
  });

  it('does not retry ordinary client errors', async () => {
    fetchSpy.mockResolvedValueOnce(fail(400));
    await expect(client.makeRequest('/me/messages')).rejects.toThrow(/400/);
    expect(fetchSpy).toHaveBeenCalledTimes(1);
  });

  it('applies to fetchBinary as well', async () => {
    fetchSpy
      .mockResolvedValueOnce(fail(502))
      .mockResolvedValueOnce(
        new Response(new Uint8Array([1, 2, 3]), { headers: { 'content-type': 'image/png' } })
      );
    const { buffer } = await client.fetchBinary('/me/photo/$value');
    expect([...buffer]).toEqual([1, 2, 3]);
    expect(fetchSpy).toHaveBeenCalledTimes(2);
  });
});

describe('parseRetryAfterMs', () => {
  it('reads delta-seconds', () => {
    expect(parseRetryAfterMs('5')).toBe(5000);
    expect(parseRetryAfterMs(' 0 ')).toBe(0);
  });

  it('reads an HTTP-date relative to now', () => {
    const now = Date.parse('2026-10-14T12:00:00Z');
    expect(parseRetryAfterMs('Wed, 14 Oct 2026 12:00:10 GMT', now)).toBe(10_000);
  });

  it('returns undefined for absent or unparseable values', () => {
    expect(parseRetryAfterMs(null)).toBeUndefined();
    expect(parseRetryAfterMs('')).toBeUndefined();
    expect(parseRetryAfterMs('soon')).toBeUndefined();
  });
});
