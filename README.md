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

**v0.2.3 — twelve tools, hardened core.** MVP scope complete, plus fixes addressing schema/handler mismatches on non-UTC times, opaque-folder-ID encoding in `list_emails`, InefficientFilter robustness, and a doc/schema mismatch in `list_calendar_events`. `AGENTS.md` documents this repo's trust-boundary guardrails for future agent-driven changes. See [CHANGELOG.md](https://github.com/xdarkoy/outlook-mcp/blob/main/CHANGELOG.md) for the full history.

| Tool                   | Live-tested (MSA) | Code-verified (AAD) |
|------------------------|-------------------|---------------------|
| `list_emails`          | ✅                 | ✅                   |
| `read_email`           | ✅                 | ✅                   |
| `search_emails`        | ✅                 | ⚠️ not live-tested   |
| `save_attachment`      | ✅                 | ✅                   |
| `mark_email_read`      | ⚠️ code-verified   | ⚠️ code-verified     |
| `move_email`           | ⚠️ code-verified   | ⚠️ code-verified     |
| `list_folders`         | ⚠️ code-verified   | ⚠️ code-verified     |
| `list_calendar_events` | ✅                 | ✅                   |
| `create_event`         | ✅                 | ✅                   |
| `update_event`         | ⚠️ code-verified   | ⚠️ code-verified     |
| `delete_event`         | ⚠️ code-verified   | ⚠️ code-verified     |
| `create_draft`         | ✅                 | ✅                   |

"Code-verified" means the path exists in the source and is exercised by the type system, schema tests, and the MCP smoke test, but a live end-to-end Graph call against a real mailbox is pending. If you hit issues, please open an issue with the exact error — we fix it fast.

## One-time setup

1. **Register an Azure AD app.** Takes ~3 minutes. See [docs/setup-admin.md](https://github.com/xdarkoy/outlook-mcp/blob/main/docs/setup-admin.md) for screenshots and the exact clicks. You end up with an **Application (client) ID** (a GUID).

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
- **`Mail.ReadWrite` caveat.** The granted scope covers read/draft/mark/move and *also* deletion. This server exposes mark/move only — `move_email` to `deleteditems` is the supported "delete" path (recoverable from trash). There is no `delete_email` tool.
- **Token cache on disk.** `~/.outlook-mcp/cache.json` is written with mode `0600` on POSIX and user-profile ACL on Windows. Cache directory path is realpath-resolved to defeat symlink escapes.
- **Attachment writes are jailed.** `save_attachment` never writes outside `OUTLOOK_MCP_ALLOWED_DIR`. Path traversal, Windows-reserved names (incl. `COM0`/`LPT0`), NTFS alternate data streams, and Unicode RTL/LTR override attacks (e.g. `invoice‮fdp.exe`) are all blocked. O_EXCL atomic writes prevent concurrent-save races; filename collisions append `(2)`, `(3)`, … Decoded byte count is verified against Graph's reported `size` before writing — malformed base64 cannot produce a silently truncated file.
- **Mail content never leaves your machine except to Graph and your chosen LLM.** No telemetry. No license check. No phone-home.
- **Deliberate exclusions.** No `Mail.Send`. No `User.ReadWrite.All`. No app-only permissions. The server always acts as the signed-in user, never as a daemon with broader reach.

## Tool caveats worth knowing

- **`search_emails`** returns results ranked by relevance, not date. On personal (MSA) accounts the `received:this-week` filter is silently ignored by Microsoft's backend — use `list_emails` with `since`/`until` if you need strict date filtering. The personal backend also has no total-count. List/search responses include a `truncated` boolean so the LLM can tell when to widen `limit` or narrow filters.
- **`list_calendar_events`** caps at `limit` (default 50, max 100) and does not follow `@odata.nextLink`. For busy calendars, narrow the window.
- **`create_event`** automatically sends meeting invitations if `attendees` is set. That is usually the desired behavior but worth knowing — if the LLM invents an attendee, a real invite goes out. When the same address is passed twice with different types, the higher-priority type wins (`required > optional > resource`) rather than last-write.
- **`update_event`** is a sparse merge — only the fields you pass are changed. EXCEPTION: if `attendees` is provided, the entire attendee list is *replaced*. Graph emails update notifications to affected attendees automatically; there is no silent-update mode.
- **`delete_event` is destructive.** For a meeting with attendees, Graph sends a cancellation notice to everyone before removing the event. There is no silent-delete mode at the API level. Make the implication explicit to the user before invoking.
- **`move_email`** issues a NEW Graph message ID — the original id becomes invalid after the move. To "delete" a message in the Outlook sense, move it to `deleteditems` (recoverable from trash).
- **`create_draft`** never sends. If you pass `body` on a reply, it REPLACES Graph's quoted-original body — the caller fully controls the outgoing text.

## Roadmap

Shipped in v0.2:
- `mark_email_read`, `move_email`, `list_folders` — completes the mail-management surface.
- `update_event`, `delete_event` — completes the calendar CRUD surface.
- Custom subfolder resolution via `list_folders` (returned IDs work in `list_emails` / `move_email`).
- Truncation hints on every list/search response.
- Hardened auth retry (in-flight refresh dedup so parallel 401s don't trigger duplicate refresh-token rotations).

v0.3 candidates:
- Live-test the new tools against MSA + AAD tenants.
- Streaming download for very large attachments via `/$value`.
- Mocked-Graph handler tests for the five new tools.
- Contacts read/search.
- Tasks / To-Do.
- OneDrive file search + fetch.

Pro-feature candidates:
- Teams chat search.
- SharePoint document RAG hooks.
- Docker image for enterprise deployments.

Issues and PRs welcome.

## Development

```bash
git clone …
cd outlook-mcp
npm install
npm run build
npm test
```

Tests: 100+ unit tests across eight files — path safety (RLO visual attacks, zero-width chars, Windows-reserved names incl. COM0/LPT0), create_event timezone logic, account-type detection, MSA search-query escaping (incl. KQL backslash hazard), attendee dedup tiebreaker, schema-level refinement (separator rejection on `targetFilename`, range limits, required fields, naive vs. offset datetime acceptance), Graph `$filter` builder (incl. InefficientFilter baseline workaround, OData quote escaping), plus an MCP stdio smoke test that verifies all 12 tools register. Real Microsoft Graph calls are exercised via `scripts/live-call.mjs` for manual acceptance testing.

## License

MIT. See [LICENSE](https://github.com/xdarkoy/outlook-mcp/blob/main/LICENSE).
