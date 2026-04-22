# Azure AD App Registration — Setup Guide

> **For personal users:** if you have a @hotmail.com, @outlook.com, or @live.com account, follow the "Personal / consumer" path. 3 minutes.
>
> **For IT admins:** if you are deploying this for a company (Microsoft Entra ID tenant), follow the "Work / school / tenant" path. You'll review exactly which Graph permissions the app needs — there is nothing hidden.

---

## What you are creating

A **Public client** (a.k.a. native / desktop) application registration. It has:

- No client secret (public clients use device-code flow).
- No admin-consented high-privilege scopes.
- Delegated permissions only — the app always acts as the signed-in user, never as a daemon.

### Required delegated permissions

| Permission            | Why                                                |
|-----------------------|----------------------------------------------------|
| `offline_access`      | Refresh tokens so the user doesn't re-sign-in hourly. |
| `User.Read`           | Identify which account is currently signed in.    |
| `Mail.Read`           | `list_emails`, `read_email`, `search_emails`.     |
| `Mail.ReadWrite`      | `create_draft` (writes drafts only — see below).  |
| `Calendars.ReadWrite` | `list_calendar_events`, `create_event`.           |

### NOT requested (by design)

- **`Mail.Send`** — this server physically cannot send email. Every draft must be reviewed and sent manually in Outlook. This is an intentional trust boundary.
- **`Mail.Delete`** — the server cannot delete messages.
- **No application permissions** — nothing runs without a live user.

---

## Personal / consumer walkthrough

*(Screenshots of each step live in `docs/screenshots/` — placeholders marked below.)*

1. Go to <https://portal.azure.com> and sign in with the Microsoft account you want `outlook-mcp` to access.

2. Search for **"App registrations"** in the top bar. *`[screenshot: portal-search.png]`*

3. Click **New registration**.

4. **Name:** anything — e.g. `outlook-mcp-local`.

5. **Supported account types:** choose **"Accounts in any organizational directory and personal Microsoft accounts"**. This single choice makes the app usable for both consumer and work accounts. *`[screenshot: account-types.png]`*

6. **Redirect URI:** leave blank. Click **Register**.

7. On the Overview page, copy the **Application (client) ID** (a GUID). This is your `OUTLOOK_MCP_CLIENT_ID`. *`[screenshot: overview-client-id.png]`*

8. Left menu → **Authentication**.

9. Scroll down to **Advanced settings**. Find the toggle labeled **"Allow public client flows"** (older portals) or **"Enable the following mobile and desktop flows"** (newer portals) and set it to **Yes**. Save. *`[screenshot: allow-public-client.png]`*

10. Left menu → **API permissions**.

11. Click **Add a permission** → **Microsoft Graph** → **Delegated permissions**, then tick the five listed above: `offline_access`, `User.Read`, `Mail.Read`, `Mail.ReadWrite`, `Calendars.ReadWrite`. **Add permissions**. *`[screenshot: permissions-list.png]`*

Done. Set `OUTLOOK_MCP_CLIENT_ID` in your MCP client config and run:

```bash
OUTLOOK_MCP_CLIENT_ID=<your-id> npx outlook-mcp-local login
```

---

## Work / school / tenant walkthrough

Everything in the consumer path applies, plus:

1. At step 5, you may choose **"Accounts in this organizational directory only (Single tenant)"** if you want to lock the app to your own tenant. Multi-tenant is also fine for related tenants.

2. At step 11, after adding the five delegated permissions, an admin should click **Grant admin consent for \<Tenant\>**. Without it, each user sees a consent dialog on first login — also acceptable, just more friction.

3. **Conditional Access**: if your tenant enforces CA policies requiring a compliant/managed device, users need to run `outlook-mcp` from such a device. Device-code flow respects CA.

4. **Application scope restriction (optional):** use an [ApplicationAccessPolicy](https://learn.microsoft.com/graph/auth-limit-mailbox-access) or [RBAC for Applications](https://learn.microsoft.com/graph/rbac-for-apps) to limit which mailboxes the registration can be used against. Belt-and-braces — the app only ever accesses the signed-in user's mailbox anyway.

### Reviewing exactly what the server does

Everything the server can do is in [`src/tools/`](../src/tools/). One file per tool, each ~100 lines, all calling Microsoft Graph directly. There is no telemetry, no outbound network call that isn't to `graph.microsoft.com` or `login.microsoftonline.com`.

---

## Validating your setup

After login:

```bash
OUTLOOK_MCP_CLIENT_ID=<your-id> npx outlook-mcp-local login
```

Expected output:

```
[outlook-mcp] starting device-code login...

==================== outlook-mcp: Sign-in required ====================
To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code ABCD-1234 to authenticate.
========================================================================

[outlook-mcp] ✓ signed in. Refresh token cached at ~/.outlook-mcp/cache.json.
```

If you see that, you're done — tool calls from your MCP client (Claude Desktop etc.) will work without further prompts.

### Most common setup failures

- **`AADSTS7000218`** — "Allow public client flows" is off. Go back to Authentication → Advanced settings.
- **403 from Graph** — admin consent missing (work tenant) or a delegated permission is missing. Re-check step 11.
- **Wrong tenant** — for personal accounts use the default `common`. For a specific work tenant, set `OUTLOOK_MCP_TENANT` to the tenant GUID or domain.
- **`OUTLOOK_MCP_CLIENT_ID is not set`** — restart your MCP client completely after editing its config (closing the window is not enough; quit from the system tray).

---

## Uninstalling cleanly

1. Delete `~/.outlook-mcp/` to revoke the local refresh token.
2. In Azure Portal → App registrations → your app → **Delete** to remove the registration.
3. Users can additionally revoke the grant at <https://myaccount.microsoft.com/consent>.
