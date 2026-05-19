import { Client, type AuthenticationProvider } from "@microsoft/microsoft-graph-client";
import { getAccessToken } from "../auth/msal.js";

/**
 * Single shared Graph SDK client.
 *
 * AuthenticationProvider delegates to MSAL on every call. MSAL caches and
 * refreshes internally, so we don't need our own token-TTL bookkeeping here.
 *
 * Auth refresh on 401: the SDK's default middleware does NOT retry a request
 * when the access token rotates out from under us mid-session (user revokes
 * consent, refresh-token lifetime exceeded, etc.). withAuthRetry handles
 * that — see its docblock.
 *
 * The earlier design used a module-level `_forceFreshToken` flag that the
 * auth provider read-and-cleared. That was racy under parallel tool calls:
 * if tool A's retry set the flag, an in-flight tool B could consume it,
 * leaving A's retry without a fresh token. We now force-refresh MSAL's
 * cache *directly* on retry (throwing away the returned token); MSAL is
 * the single source of truth, and the next provider call picks up the new
 * token via the normal silent path. No shared mutable retry state.
 */

const authProvider: AuthenticationProvider = {
  async getAccessToken() {
    return getAccessToken();
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

interface AuthErrorShape {
  statusCode?: number;
  code?: string;
}

function isAuthError(err: unknown): boolean {
  const e = err as AuthErrorShape;
  return e?.statusCode === 401 || e?.code === "InvalidAuthenticationToken";
}

/**
 * In-flight dedup for the force-refresh path.
 *
 * MSAL's cache plugin serializes cache I/O (read/write), but NOT the actual
 * token request to the authority. Two concurrent withAuthRetry calls hitting
 * 401 at the same time would otherwise fire two parallel refresh requests
 * — and because refresh-token rotation invalidates the previous refresh
 * token, the slower of the two can fail with `invalid_grant`, surfacing as
 * a confusing re-auth prompt right after a successful one.
 *
 * One in-flight promise per process; concurrent waiters share it. The
 * promise resolves once MSAL has the new token in cache; the actual access
 * token returned is discarded because the auth provider will fetch it
 * fresh on retry.
 */
let _refreshInFlight: Promise<void> | null = null;
function refreshTokenOnce(): Promise<void> {
  if (_refreshInFlight) return _refreshInFlight;
  _refreshInFlight = (async () => {
    try {
      await getAccessToken(undefined, { forceRefresh: true });
    } finally {
      _refreshInFlight = null;
    }
  })();
  return _refreshInFlight;
}

/**
 * Run a Graph call; on auth failure retry exactly once with a forced fresh
 * MSAL token. Other 4xx/5xx pass through unchanged.
 *
 * We match auth failures on BOTH the HTTP status code AND the Graph error
 * code, because the SDK surfaces auth failures differently depending on
 * whether the response was parsed as a Graph error envelope or a raw HTTP
 * fault. explainGraphError already checks both — withAuthRetry now does too.
 *
 * If the refresh itself fails, that error propagates — the original 401 is
 * shadowed, but the actionable signal (re-auth needed) reaches the caller
 * via explainGraphError on the next tool call.
 */
export async function withAuthRetry<T>(fn: () => Promise<T>): Promise<T> {
  try {
    return await fn();
  } catch (err) {
    if (!isAuthError(err)) throw err;
    await refreshTokenOnce();
    return await fn();
  }
}
