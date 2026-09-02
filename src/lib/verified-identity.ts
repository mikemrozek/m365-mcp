import { createHash } from 'node:crypto';
import logger from '../logger.js';
import { getCloudEndpoints, parseCloudType } from '../cloud-config.js';
import { actorFromToken, tenantFromToken } from './microsoft-auth.js';
import type { RequestActor } from '../request-context.js';

/**
 * Authoritative caller identity, for the one case where this server *asserts*
 * who the caller is to a third party rather than merely labelling its own logs.
 *
 * Why this is not JWT signature validation
 * ----------------------------------------
 * The obvious implementation — fetch the Entra JWKS and verify the access
 * token's signature, issuer, audience and tenant — does not work for the tokens
 * we hold. Microsoft Graph access tokens (`aud=https://graph.microsoft.com`,
 * `ver=1.0`) carry a `nonce` in the JWT header, and the signature is computed
 * over a transformed header. Standard JWKS verification therefore fails on a
 * perfectly valid token. Graph tokens are deliberately opaque to everyone but
 * Graph; Microsoft's guidance is that only the resource validates them.
 * Confirmed empirically against a live tenant token on 2026-08-25.
 *
 * So instead of asking "is this token's signature good?", we ask the only party
 * entitled to answer: we call Graph `/me` with the token and take the identity
 * from the response. A forged or altered token cannot produce a `/me` response.
 * This is strictly stronger than reading unverified claims, and it keeps Graph
 * as the authority — which is already true for every other call this server
 * makes.
 *
 * Relationship to `actorFromToken`
 * --------------------------------
 * `actorFromToken` stays exactly as it is: a best-effort claims decode feeding
 * `usage-log`, where Graph is the authority on the request anyway and a bad
 * token simply fails downstream. It must NOT become strict — every one of the
 * server's Graph tools depends on that path, and failing it closed would put a
 * new failure mode in front of all of them to serve one caller.
 *
 * This module is the strict path, used only where identity is asserted outward.
 * It fails closed: no verified identity, no assertion.
 */

/** Identity as reported by Graph itself, not as claimed by the token. */
export interface VerifiedIdentity {
  /** Entra object id (`/me` `id`). */
  oid: string;
  /** User principal name (`/me` `userPrincipalName`). */
  upn: string;
  /** Tenant id. Graph `/me` does not return this, so it comes from the token claims. */
  tid?: string;
}

export class IdentityVerificationError extends Error {
  constructor(
    message: string,
    readonly reason: 'graph_rejected' | 'graph_unavailable' | 'tenant_mismatch' | 'malformed'
  ) {
    super(message);
    this.name = 'IdentityVerificationError';
  }
}

interface CacheEntry {
  identity: VerifiedIdentity;
  expiresAt: number;
}

/**
 * Cache keyed by a SHA-256 of the token — never the token itself, so a heap
 * dump or a stray log line cannot yield a usable credential (SEC-2026-001).
 * Short TTL: this is a freshness/rate tradeoff, not a session. A revoked or
 * expired token stops working within the window.
 */
const CACHE_TTL_MS = 5 * 60 * 1000;
const MAX_CACHE_ENTRIES = 500;
const cache = new Map<string, CacheEntry>();

function cacheKey(token: string): string {
  return createHash('sha256').update(token).digest('hex');
}

function pruneCache(now: number): void {
  for (const [key, entry] of cache) {
    if (entry.expiresAt <= now) cache.delete(key);
  }
  // Bound the map even if every entry is still live (many concurrent users).
  while (cache.size > MAX_CACHE_ENTRIES) {
    const oldest = cache.keys().next();
    if (oldest.done) break;
    cache.delete(oldest.value);
  }
}

/** Clears memoised identities. Test seam. */
export function clearVerifiedIdentityCache(): void {
  cache.clear();
}

/**
 * The tenant this server is willing to assert identities from. `common` (the
 * default) means "no restriction configured" and disables the check; set
 * MS365_MCP_TENANT_ID to a specific tenant to enforce it.
 */
function expectedTenantId(): string | undefined {
  const configured = process.env.MS365_MCP_TENANT_ID?.trim();
  if (!configured || configured.toLowerCase() === 'common') return undefined;
  return configured.toLowerCase();
}

/**
 * Resolve the caller's identity via Graph, or throw.
 *
 * Call this immediately before asserting an identity to a third party. Do not
 * use it to gate ordinary Graph tool calls — those are already gated by Graph.
 */
export async function verifyCallerIdentity(accessToken: string): Promise<VerifiedIdentity> {
  if (!accessToken || typeof accessToken !== 'string') {
    throw new IdentityVerificationError('No access token on the request', 'malformed');
  }

  const now = Date.now();
  const key = cacheKey(accessToken);
  const hit = cache.get(key);
  if (hit && hit.expiresAt > now) {
    return hit.identity;
  }

  const cloudType = parseCloudType(process.env.MS365_MCP_CLOUD_TYPE);
  const { graphApi } = getCloudEndpoints(cloudType);

  let response: Response;
  try {
    response = await fetch(`${graphApi}/v1.0/me?$select=id,userPrincipalName`, {
      headers: { Authorization: `Bearer ${accessToken}`, Accept: 'application/json' },
    });
  } catch (error) {
    // Network failure is not an authorization failure, but it is still a
    // refusal to assert: we do not know who this is.
    throw new IdentityVerificationError(
      `Could not reach Graph to verify caller identity: ${(error as Error).message}`,
      'graph_unavailable'
    );
  }

  if (!response.ok) {
    throw new IdentityVerificationError(
      `Graph rejected the caller's token (HTTP ${response.status})`,
      response.status >= 500 ? 'graph_unavailable' : 'graph_rejected'
    );
  }

  const body = (await response.json()) as { id?: unknown; userPrincipalName?: unknown };
  const oid = typeof body.id === 'string' ? body.id.trim() : '';
  const upn = typeof body.userPrincipalName === 'string' ? body.userPrincipalName.trim() : '';
  if (!oid || !upn) {
    throw new IdentityVerificationError(
      'Graph returned no usable identity for the caller',
      'malformed'
    );
  }

  // Tenant comes from the token claims: /me does not report it. The claims are
  // unverified on their own, but a token that Graph just accepted for this user
  // is not attacker-chosen, so the pairing is sound. Read `tid` directly rather
  // than via actorFromToken, which returns nothing when a token carries no
  // oid/upn — that would misreport a decodable token as an unknown tenant.
  const tid = tenantFromToken(accessToken)?.toLowerCase();

  const expected = expectedTenantId();
  if (expected) {
    if (!tid) {
      // Opaque or claim-less token: Graph accepted it, but we cannot tell which
      // tenant it belongs to. Asserting a tenant we cannot name is exactly what
      // this module exists to prevent, so refuse.
      throw new IdentityVerificationError(
        'Caller tenant could not be determined from the token; refusing to assert identity',
        'tenant_mismatch'
      );
    }
    if (tid !== expected) {
      throw new IdentityVerificationError(
        `Caller is from tenant ${tid}, not the configured tenant`,
        'tenant_mismatch'
      );
    }
  }

  // Tripwire, not a gate. Graph is the authority either way; a disagreement
  // means something upstream is rewriting tokens and is worth seeing in a log.
  const claimedOid = actorFromToken(accessToken)?.oid;
  if (claimedOid && claimedOid.toLowerCase() !== oid.toLowerCase()) {
    logger.warn(
      `Token oid claim disagrees with Graph /me (claim=${claimedOid}, graph=${oid}); using Graph.`
    );
  }

  const identity: VerifiedIdentity = { oid, upn, tid };
  pruneCache(now);
  cache.set(key, { identity, expiresAt: now + CACHE_TTL_MS });
  return identity;
}

/** Narrowing helper for callers that want to branch on the failure reason. */
export function isIdentityVerificationError(e: unknown): e is IdentityVerificationError {
  return e instanceof IdentityVerificationError;
}

/** The unverified, log-only actor. Kept distinct so the two cannot be confused. */
export type { RequestActor };
