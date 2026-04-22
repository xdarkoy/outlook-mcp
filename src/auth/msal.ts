import {
  PublicClientApplication,
  LogLevel,
  type AuthenticationResult,
  type Configuration,
  type ICachePlugin,
  type TokenCacheContext,
} from "@azure/msal-node";
import { readCache, writeCache } from "./tokenCache.js";

/**
 * Microsoft identity / MSAL wiring.
 *
 * Flow:
 *   1. Try silent token acquisition from the persistent cache.
 *   2. If no cached account exists, fall back to device code flow.
 *      The device-code prompt is written to STDERR — stdout is reserved
 *      for the MCP JSON-RPC stream.
 *
 * Tenancy: `common` means the app accepts both AAD work/school accounts
 * (multi-tenant) AND personal Microsoft accounts. This matches our product
 * positioning: the paying market is Azure AD firms, but consumers come
 * along for free.
 */

const DEFAULT_CLIENT_ID = process.env.OUTLOOK_MCP_CLIENT_ID ?? "";
const AUTHORITY = `https://login.microsoftonline.com/${process.env.OUTLOOK_MCP_TENANT ?? "common"}`;

/**
 * Scopes requested. Note the deliberate absence of `Mail.Send` — this MCP
 * server is physically incapable of sending mail. Only drafts.
 *
 * `offline_access` is required for refresh tokens, otherwise the user would
 * have to re-authenticate whenever the access token expires (~1h).
 */
export const DEFAULT_SCOPES = [
  "offline_access",
  "User.Read",
  "Mail.Read",
  "Mail.ReadWrite",
  "Calendars.ReadWrite",
];

/**
 * Serialize cache read/write with an in-process promise chain.
 *
 * Without this, two parallel `acquireTokenSilent` calls can interleave their
 * before/after hooks — both read the same snapshot, both write back, and the
 * later write silently discards the rotated refresh token from the earlier
 * one. That forces the user to re-authenticate interactively later, for no
 * visible reason. Cross-process safety is explicitly out of scope for MVP
 * (one server = one MSAL instance).
 */
let cacheChain: Promise<void> = Promise.resolve();
function serialize<T>(task: () => Promise<T>): Promise<T> {
  const next = cacheChain.then(task, task);
  cacheChain = next.then(
    () => undefined,
    () => undefined,
  );
  return next;
}

const cachePlugin: ICachePlugin = {
  beforeCacheAccess(ctx: TokenCacheContext) {
    return serialize(async () => {
      const data = await readCache();
      if (data) ctx.tokenCache.deserialize(data);
    });
  },
  afterCacheAccess(ctx: TokenCacheContext) {
    return serialize(async () => {
      if (ctx.cacheHasChanged) {
        await writeCache(ctx.tokenCache.serialize());
      }
    });
  },
};

function buildClient(clientId: string): PublicClientApplication {
  const config: Configuration = {
    auth: {
      clientId,
      authority: AUTHORITY,
    },
    cache: { cachePlugin },
    system: {
      loggerOptions: {
        logLevel: LogLevel.Warning,
        piiLoggingEnabled: false,
        loggerCallback: (_level, message) => {
          // MSAL logs go to stderr so they never contaminate MCP stdout.
          process.stderr.write(`[msal] ${message}\n`);
        },
      },
    },
  };
  return new PublicClientApplication(config);
}

let _client: PublicClientApplication | null = null;
function client(): PublicClientApplication {
  if (_client) return _client;
  if (!DEFAULT_CLIENT_ID) {
    throw new Error(
      "OUTLOOK_MCP_CLIENT_ID environment variable is not set. " +
        "Register an app in Azure AD (see docs/setup-admin.md), then set OUTLOOK_MCP_CLIENT_ID to the Application (client) ID."
    );
  }
  _client = buildClient(DEFAULT_CLIENT_ID);
  return _client;
}

/**
 * Return the first signed-in account from the cache, or null if the user
 * has not signed in yet. Used by accountType detection to distinguish
 * MSA from AAD via the account's tenantId — more reliable than decoding
 * the access token, since MSA Graph access tokens are opaque (not JWTs).
 */
export async function getSignedInAccount() {
  const app = client();
  const accounts = await app.getTokenCache().getAllAccounts();
  return accounts[0] ?? null;
}

async function trySilent(
  scopes: string[],
  forceRefresh: boolean,
): Promise<AuthenticationResult | null> {
  const app = client();
  const accounts = await app.getTokenCache().getAllAccounts();
  if (accounts.length === 0) return null;
  // MVP: one user per install. If multiple accounts exist in the cache
  // (e.g., the user signed in twice with different accounts), we pick the
  // first and warn on stderr so it's visible in logs.
  if (accounts.length > 1) {
    const names = accounts.map((a) => a.username).join(", ");
    process.stderr.write(
      `[outlook-mcp] multiple accounts in cache (${names}); using '${accounts[0]!.username}'. Set OUTLOOK_MCP_CACHE_DIR to a fresh directory to start over.\n`,
    );
  }
  const account = accounts[0]!;
  try {
    return await app.acquireTokenSilent({ account, scopes, forceRefresh });
  } catch {
    return null;
  }
}

/**
 * Error thrown by getAccessToken() when no cached token exists and
 * interactive login is disabled (the default for server-mode).
 *
 * The message is surfaced verbatim to the MCP client, which lets the LLM
 * tell the user what to do. MCP clients like Claude Desktop hide stderr —
 * so if we started a blocking device-code flow on the first tool call, the
 * user would see a spinner that "hangs forever" and typically cancels
 * before the code becomes visible. We refuse to hang. Instead we demand
 * a one-time CLI pre-auth via `npx outlook-mcp-local login`.
 */
export class LoginRequiredError extends Error {
  constructor() {
    super(
      "Not signed in. Run this in a terminal ONCE to authenticate:\n\n" +
        "    npx outlook-mcp-local login\n\n" +
        "The command prints a device-code URL; sign in with your Microsoft " +
        "account and you're done. The refresh token is cached locally; " +
        "subsequent tool calls work without any prompt.",
    );
    this.name = "LoginRequiredError";
  }
}

async function acquireViaDeviceCode(scopes: string[]): Promise<AuthenticationResult> {
  const app = client();
  const result = await app.acquireTokenByDeviceCode({
    scopes,
    deviceCodeCallback: (response) => {
      // Stderr only. Stdout is the MCP JSON-RPC channel. Also only visible
      // when the user is watching a terminal — in server mode we refuse to
      // reach this codepath (see getAccessToken's `allowInteractive` guard).
      process.stderr.write(
        "\n==================== outlook-mcp: Sign-in required ====================\n" +
          response.message +
          "\n========================================================================\n\n",
      );
    },
  });
  if (!result) throw new Error("Device code flow returned no authentication result.");
  return result;
}

/**
 * Public entry: return a valid access token, refreshing or prompting as needed.
 *
 * `forceRefresh` bypasses MSAL's in-memory token cache to fetch a new access
 * token from the authority — used by graph/withRetry after a 401.
 *
 * `allowInteractive` must be explicitly set to true to allow the blocking
 * device-code flow. In MCP server mode it is false, so tool calls fail fast
 * with a LoginRequiredError instead of silently hanging. The `login`
 * subcommand (see dist/cli/login.ts) opts in.
 */
export async function getAccessToken(
  scopes: string[] = DEFAULT_SCOPES,
  opts?: { forceRefresh?: boolean; allowInteractive?: boolean },
): Promise<string> {
  const silent = await trySilent(scopes, !!opts?.forceRefresh);
  if (silent?.accessToken) return silent.accessToken;
  if (!opts?.allowInteractive) {
    throw new LoginRequiredError();
  }
  const interactive = await acquireViaDeviceCode(scopes);
  if (!interactive.accessToken) {
    throw new Error("Authentication succeeded but no access token was returned.");
  }
  return interactive.accessToken;
}
