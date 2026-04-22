import { getAccessToken } from "../auth/msal.js";

/**
 * One-time interactive login via device-code flow. Writes the refresh
 * token to the local cache (~/.outlook-mcp/cache.json) so the MCP server
 * can run unattended afterwards.
 *
 * Subsequent `npx outlook-mcp-local` invocations (stdio MCP mode) reuse
 * this cached token silently. When the refresh token expires (90 days
 * typical), run this command again.
 */
export async function runLogin(): Promise<number> {
  process.stderr.write("[outlook-mcp] starting device-code login...\n");
  try {
    await getAccessToken(undefined, { allowInteractive: true });
    process.stderr.write(
      "\n[outlook-mcp] ✓ signed in. Refresh token cached at ~/.outlook-mcp/cache.json.\n" +
        "[outlook-mcp] You can now run the MCP server without a login prompt.\n",
    );
    return 0;
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    process.stderr.write(`\n[outlook-mcp] login failed: ${msg}\n`);
    return 1;
  }
}
