import { Client, type AuthenticationProvider } from "@microsoft/microsoft-graph-client";
import { getAccessToken } from "../auth/msal.js";

/**
 * Single shared Graph SDK client.
 *
 * AuthenticationProvider delegates to MSAL on every call. MSAL caches and
 * refreshes internally, so we don't need our own token-TTL bookkeeping here.
 *
 * Auth refresh on 401: the default middleware pipeline does NOT retry a
 * request if it fails with 401 after the access token was rotated out from
 * under us (e.g., the user revoked consent, or a long-running process
 * outlived the refresh-token lifetime and reacquired mid-session). We flag
 * a "force refresh" hint to the auth provider on such retries.
 *
 * The graph() helper returns a plain Client; individual tools that want the
 * retry behavior call via withRetry() below.
 */

let _forceFreshToken = false;

const authProvider: AuthenticationProvider = {
  async getAccessToken() {
    const force = _forceFreshToken;
    _forceFreshToken = false;
    return getAccessToken(undefined, { forceRefresh: force });
  },
};

let _client: Client | null = null;

export function graph(): Client {
  if (_client) return _client;
  // `defaultVersion` is only honored by Client.init, not initWithMiddleware.
  // The SDK already targets v1.0 by default; per-call override is via .version().
  _client = Client.initWithMiddleware({ authProvider });
  return _client;
}

/**
 * Run a Graph call; on a 401 retry exactly once with a forced fresh token.
 * Other 4xx/5xx pass through unchanged. Tools wrap their request in this
 * for robustness against mid-session token rotation (user revokes consent,
 * refresh-token lifetime hit, etc.).
 */
export async function withAuthRetry<T>(fn: () => Promise<T>): Promise<T> {
  try {
    return await fn();
  } catch (err) {
    const status = (err as { statusCode?: number })?.statusCode;
    if (status !== 401) throw err;
    _forceFreshToken = true;
    return await fn();
  }
}
