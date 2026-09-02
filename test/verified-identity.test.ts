import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  clearVerifiedIdentityCache,
  isIdentityVerificationError,
  verifyCallerIdentity,
} from '../src/lib/verified-identity.js';

/** Build a decodable (unsigned) JWT so the tenant/oid claim path is exercised. */
function fakeToken(claims: Record<string, unknown>): string {
  const b64 = (o: unknown) => Buffer.from(JSON.stringify(o)).toString('base64url');
  return `${b64({ typ: 'JWT', alg: 'RS256' })}.${b64(claims)}.sig`;
}

const TENANT = 'a473edd8-ba25-4f04-a0a8-e8ad25c19632';
const OID = '2784a67f-8b11-4bec-9f7d-5914cee129b0';
const UPN = 'mike.mrozek@townsquaremedia.com';

function graphOk(body: unknown) {
  return { ok: true, status: 200, json: async () => body } as unknown as Response;
}
function graphErr(status: number) {
  return { ok: false, status, json: async () => ({}) } as unknown as Response;
}

describe('verifyCallerIdentity', () => {
  const originalTenant = process.env.MS365_MCP_TENANT_ID;

  beforeEach(() => {
    clearVerifiedIdentityCache();
    delete process.env.MS365_MCP_TENANT_ID;
  });

  afterEach(() => {
    vi.restoreAllMocks();
    if (originalTenant === undefined) delete process.env.MS365_MCP_TENANT_ID;
    else process.env.MS365_MCP_TENANT_ID = originalTenant;
  });

  it('takes the identity from Graph, not from the token claims', async () => {
    vi.spyOn(global, 'fetch').mockResolvedValue(graphOk({ id: OID, userPrincipalName: UPN }));

    // Claims deliberately lie about who this is.
    const token = fakeToken({
      oid: 'attacker-supplied-oid',
      preferred_username: 'evil@x',
      tid: TENANT,
    });
    const identity = await verifyCallerIdentity(token);

    expect(identity.oid).toBe(OID);
    expect(identity.upn).toBe(UPN);
    expect(identity.tid).toBe(TENANT);
  });

  it('sends the caller token to Graph as a bearer credential', async () => {
    const spy = vi
      .spyOn(global, 'fetch')
      .mockResolvedValue(graphOk({ id: OID, userPrincipalName: UPN }));

    const token = fakeToken({ tid: TENANT });
    await verifyCallerIdentity(token);

    const [url, init] = spy.mock.calls[0];
    expect(String(url)).toContain('/v1.0/me');
    expect((init?.headers as Record<string, string>).Authorization).toBe(`Bearer ${token}`);
  });

  it('memoises so repeated calls do not re-hit Graph', async () => {
    const spy = vi
      .spyOn(global, 'fetch')
      .mockResolvedValue(graphOk({ id: OID, userPrincipalName: UPN }));

    const token = fakeToken({ tid: TENANT });
    await verifyCallerIdentity(token);
    await verifyCallerIdentity(token);

    expect(spy).toHaveBeenCalledTimes(1);
  });

  it('does not serve one caller identity to a different token', async () => {
    const spy = vi
      .spyOn(global, 'fetch')
      .mockResolvedValueOnce(graphOk({ id: OID, userPrincipalName: UPN }))
      .mockResolvedValueOnce(graphOk({ id: 'other-oid', userPrincipalName: 'other@x' }));

    const a = await verifyCallerIdentity(fakeToken({ tid: TENANT, sub: 'a' }));
    const b = await verifyCallerIdentity(fakeToken({ tid: TENANT, sub: 'b' }));

    expect(a.oid).toBe(OID);
    expect(b.oid).toBe('other-oid');
    expect(spy).toHaveBeenCalledTimes(2);
  });

  it('fails closed when Graph rejects the token', async () => {
    vi.spyOn(global, 'fetch').mockResolvedValue(graphErr(401));

    await expect(verifyCallerIdentity(fakeToken({ tid: TENANT }))).rejects.toSatisfy(
      (e: unknown) => isIdentityVerificationError(e) && e.reason === 'graph_rejected'
    );
  });

  it('distinguishes Graph being down from Graph saying no', async () => {
    vi.spyOn(global, 'fetch').mockResolvedValue(graphErr(503));
    await expect(verifyCallerIdentity(fakeToken({ tid: TENANT }))).rejects.toSatisfy(
      (e: unknown) => isIdentityVerificationError(e) && e.reason === 'graph_unavailable'
    );

    clearVerifiedIdentityCache();
    vi.spyOn(global, 'fetch').mockRejectedValue(new Error('ECONNRESET'));
    await expect(verifyCallerIdentity(fakeToken({ tid: TENANT }))).rejects.toSatisfy(
      (e: unknown) => isIdentityVerificationError(e) && e.reason === 'graph_unavailable'
    );
  });

  it('refuses when Graph returns no usable identity', async () => {
    vi.spyOn(global, 'fetch').mockResolvedValue(graphOk({ id: '', userPrincipalName: '' }));
    await expect(verifyCallerIdentity(fakeToken({ tid: TENANT }))).rejects.toSatisfy(
      (e: unknown) => isIdentityVerificationError(e) && e.reason === 'malformed'
    );
  });

  it('refuses a caller from an unexpected tenant when a tenant is configured', async () => {
    process.env.MS365_MCP_TENANT_ID = TENANT;
    vi.spyOn(global, 'fetch').mockResolvedValue(graphOk({ id: OID, userPrincipalName: UPN }));

    await expect(
      verifyCallerIdentity(fakeToken({ tid: '11111111-2222-3333-4444-555555555555' }))
    ).rejects.toSatisfy(
      (e: unknown) => isIdentityVerificationError(e) && e.reason === 'tenant_mismatch'
    );
  });

  it('accepts the configured tenant', async () => {
    process.env.MS365_MCP_TENANT_ID = TENANT;
    vi.spyOn(global, 'fetch').mockResolvedValue(graphOk({ id: OID, userPrincipalName: UPN }));

    await expect(verifyCallerIdentity(fakeToken({ tid: TENANT }))).resolves.toMatchObject({
      oid: OID,
    });
  });

  it('does not enforce a tenant when configured as common', async () => {
    process.env.MS365_MCP_TENANT_ID = 'common';
    vi.spyOn(global, 'fetch').mockResolvedValue(graphOk({ id: OID, userPrincipalName: UPN }));

    await expect(
      verifyCallerIdentity(fakeToken({ tid: 'anything-at-all' }))
    ).resolves.toMatchObject({ oid: OID });
  });

  it('rejects a missing token without calling Graph', async () => {
    const spy = vi.spyOn(global, 'fetch');
    await expect(verifyCallerIdentity('')).rejects.toSatisfy(
      (e: unknown) => isIdentityVerificationError(e) && e.reason === 'malformed'
    );
    expect(spy).not.toHaveBeenCalled();
  });
});
