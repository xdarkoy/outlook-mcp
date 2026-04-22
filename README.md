# outlook-mcp

**The Anti-Copilot.** A local-first [Model Context Protocol](https://modelcontextprotocol.io) server that exposes Outlook / Microsoft 365 as tools for any MCP-capable LLM client — Claude Desktop, Claude Code, VS Code, Cursor, Continue, AnythingLLM, and anything else that speaks MCP.

Your mail, your calendar, your attachments. **Your** model. No Copilot. No cloud middleman between your data and your LLM.

## Why this exists

Microsoft Copilot ships a single, closed AI experience that many privacy-conscious users and companies will not adopt — data-protection concerns, per-seat pricing, vendor lock-in. The "Bring Your Own Model" (BYOM) wave solves that, but it lacks what Copilot has for free: a tight integration with Outlook.

`outlook-mcp` is that integration. It runs locally, speaks Microsoft Graph directly from your machine, and presents a clean set of tools to whatever LLM client you prefer.

### What makes this different

- **Local-first.** The server runs on your machine. Your mail content is fetched directly from Graph to your LLM — no SaaS middleman.
- **`save_attachment` actually saves to your disk.** A pure cloud assistant can never do that. This is the killer feature for invoice/contract/CV workflows.
- **Draft-first, by design.** The OAuth scope `Mail.Send` is deliberately *not* requested. This server physically cannot send email. Every generated reply lands in your Drafts folder for human review. No hallucinated mail to the boss.
- **Works with any MCP client.** You are not locked into one chat UI.
- **Personal AND work accounts.** Auto-detects MSA (hotmail.com, outlook.com) vs. AAD (work/school) and routes each tool to the right backend.

## Status

**v0.1.0 — ready for personal use.** MVP scope complete; all 7 tools implemented and live-verified against a real Outlook.com mailbox.

| Tool                   | Live-tested (MSA) | Code-verified (AAD) |
|------------------------|-------------------|---------------------|
| `list_emails`          | ✅                 | ✅                   |
| `read_email`           | ✅                 | ✅                   |
| `search_emails`        | ✅                 | ⚠️ not live-tested   |
| `save_attachment`      | ✅                 | ✅                   |
| `list_calendar_events` | ✅                 | ✅                   |
| `create_event`         | ✅                 | ✅                   |
| `create_draft`         | ✅                 | ✅                   |

"Code-verified" means the AAD path exists in the source and is exercised by the dual-backend routing, but a live end-to-end test against a work/school tenant is pending. If you hit issues on AAD, please open an issue with the exact error — we fix it fast.

## One-time setup

1. **Register an Azure AD app.** Takes ~3 minutes. See [docs/setup-admin.md](docs/setup-admin.md) for screenshots and the exact clicks. You end up with an **Application (client) ID** (a GUID).

2. **One-time login:**
   ```bash
   OUTLOOK_MCP_CLIENT_ID=<your-client-id> npx outlook-mcp-local login
   ```
   A device-code URL appears in the terminal. Open it, sign in, confirm permissions. The refresh token is cached at `~/.outlook-mcp/cache.json`.

3. **Wire the server into your LLM client** (see below). From now on the server runs silently — no more login prompts until the refresh token expires (typically 90 days).

## Claude Desktop

Edit `claude_desktop_config.json` (Windows: `%APPDATA%\Claude\`, macOS: `~/Library/Application Support/Claude/`):

```jsonc
{
  "mcpServers": {
    "outlook": {
      "command": "npx",
      "args": ["-y", "outlook-mcp-local"],
      "env": {
        "OUTLOOK_MCP_CLIENT_ID": "<your-app-client-id>"
      }
    }
  }
}
```

Restart Claude Desktop. Ask: *"List my 5 most recent emails."*

## Claude Code CLI

```bash
claude mcp add outlook --scope user \
  --env OUTLOOK_MCP_CLIENT_ID=<your-id> \
  -- npx -y outlook-mcp-local
```

## VS Code / Cursor / Continue / AnythingLLM

Same `command` + `args` + `env` shape — check your client's MCP config docs for the exact path.

## Configuration

| Variable                         | Required | Default                     | Purpose                                              |
|----------------------------------|----------|-----------------------------|------------------------------------------------------|
| `OUTLOOK_MCP_CLIENT_ID`          | **yes**  | —                           | Your Azure AD Application (client) ID.               |
| `OUTLOOK_MCP_TENANT`             | no       | `common`                    | `common`, `organizations`, or a tenant ID.           |
| `OUTLOOK_MCP_ALLOWED_DIR`        | no       | `~/Downloads/outlook-mcp/`  | Where `save_attachment` may write files.             |
| `OUTLOOK_MCP_CACHE_DIR`          | no       | `~/.outlook-mcp/`           | Token cache location.                                |
| `OUTLOOK_MCP_MAX_ATTACHMENT_MB`  | no       | `50`                        | Hard cap before `save_attachment` aborts (OOM guard).|

## Subcommands

```
outlook-mcp-local         # Run as MCP stdio server (the default — this is what MCP clients invoke)
outlook-mcp-local login   # Interactive one-time sign-in via device-code flow
outlook-mcp-local help    # Show this help
```

## Security model

- **No send capability.** The OAuth token has no `Mail.Send` scope. Even if the LLM asks, the server cannot send. Draft-first, always.
- **Token cache on disk.** `~/.outlook-mcp/cache.json` is written with mode `0600` on POSIX and user-profile ACL on Windows. Cache directory path is realpath-resolved to defeat symlink escapes.
- **Attachment writes are jailed.** `save_attachment` never writes outside `OUTLOOK_MCP_ALLOWED_DIR`. Path traversal, Windows-reserved names, NTFS alternate data streams, and Unicode RTL/LTR override attacks (e.g. `invoice‮fdp.exe`) are all blocked. O_EXCL atomic writes prevent concurrent-save races; filename collisions append `(2)`, `(3)`, …
- **Mail content never leaves your machine except to Graph and your chosen LLM.** No telemetry. No license check. No phone-home.
- **Deliberate exclusions.** No `Mail.Delete`. No `User.ReadWrite.All`. No app-only permissions. The server always acts as the signed-in user, never as a daemon with broader reach.

## Tool caveats worth knowing

- **`search_emails`** returns results ranked by relevance, not date. On personal (MSA) accounts the `received:this-week` filter is silently ignored by Microsoft's backend — use `list_emails` with `since`/`until` if you need strict date filtering. The personal backend also has no total-count.
- **`list_calendar_events`** caps at `limit` (default 50, max 100) and does not follow `@odata.nextLink`. For busy calendars, narrow the window.
- **`create_event`** automatically sends meeting invitations if `attendees` is set. That is usually the desired behavior but worth knowing — if the LLM invents an attendee, a real invite goes out.
- **`create_draft`** never sends. If you pass `body` on a reply, it REPLACES Graph's quoted-original body — the caller fully controls the outgoing text.

## Roadmap

v0.2 candidates:
- Streaming download for very large attachments via `/$value`
- Custom mailbox subfolder resolution in `list_emails`
- Contacts read/search
- Tasks / To-Do
- OneDrive file search + fetch

Pro-feature candidates:
- Teams chat search
- SharePoint document RAG hooks
- Docker image for enterprise deployments

Issues and PRs welcome.

## Development

```bash
git clone …
cd outlook-mcp
npm install
npm run build
npm test
```

Tests: 40+ unit tests covering path safety (RLO visual attacks, zero-width chars, Windows-reserved names), create_event timezone logic, account-type detection, and an MCP stdio smoke test. Real Microsoft Graph calls are exercised via `scripts/live-call.mjs` for manual acceptance testing.

## License

MIT. See [LICENSE](LICENSE).
