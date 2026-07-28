import { AsyncLocalStorage } from 'node:async_hooks';

/** Identity of the caller, derived from the Microsoft access token's JWT claims. */
export interface RequestActor {
  /** Stable Entra object id (`oid`) — the durable per-user key for usage reporting. */
  oid?: string;
  /** Human-readable sign-in name (`preferred_username` / `upn` / `unique_name`). */
  upn?: string;
  /** Tenant id (`tid`). */
  tid?: string;
}

export interface RequestContext {
  accessToken: string;
  /** Present in HTTP/OAuth mode; absent for stdio/CLI (single local user). */
  actor?: RequestActor;
}

export const requestContext = new AsyncLocalStorage<RequestContext>();

export function getRequestTokens(): RequestContext | undefined {
  return requestContext.getStore();
}

/** Convenience accessor for the caller's identity, if any. */
export function getRequestActor(): RequestActor | undefined {
  return requestContext.getStore()?.actor;
}
