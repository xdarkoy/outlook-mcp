# Agent Maintenance Notes

This repo has product-level trust boundaries that are easy to break with a small code change. Before changing any of these areas, search prior agent sessions when a local history tool is available (for example `ctx search` or the host's built-in session history) and cite the relevant decision in the final summary:

- OAuth scopes in `src/auth/msal.ts`
- outbound mail behavior in `src/tools/createDraft.ts`
- attachment persistence in `src/tools/saveAttachment.ts` and `src/util/paths.ts`
- calendar create/update/delete behavior in `src/tools/createEvent.ts`, `src/tools/updateEvent.ts`, and `src/tools/deleteEvent.ts`

Mail boundary: this server must not request `Mail.Send` and must not expose a tool that sends mail. `create_draft` is the outbound-mail tool; it creates drafts only, and a human sends them from Outlook.

Calendar boundary: calendar tools may create meetings, update meetings, and delete events. These operations can notify attendees through Microsoft Graph. Keep that side effect visible in tool descriptions and results.