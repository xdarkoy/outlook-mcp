# Changelog

All notable changes to this project are documented in this file. The
format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and
this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.1] - 2026-05-19

Bug-fix release. All four issues were external-reviewer findings on v0.2.0; no API surface change beyond the response-shape rename in `update_event`.

### Fixed

- **`create_event` / `update_event` schema vs. handler mismatch (HIGH).** The Zod schema accepted only offset-bearing datetimes (`datetime({ offset: true })`), but the handler explicitly rejected those when `timeZone` was a non-UTC IANA zone — making the documented `naive datetime + Europe/Berlin` path unreachable through the schema. Schema now uses `datetime({ offset: true, local: true })` to accept both forms; the handler's normalization rules (the source of truth for the offset/tz matrix) are unchanged.
- **`list_emails` opaque folder IDs were mis-parsed (HIGH).** The Graph SDK does NOT percent-encode path segments. Opaque folder IDs returned by `list_folders` can contain `/`, `+`, or `=` — those landed in the URL as path delimiters, hitting the wrong resource. `list_emails` now `encodeURIComponent`s the folder segment, matching what `list_folders` already does for `parentFolderId`. Well-known names ("inbox", "sentitems", …) are ASCII-only, so encoding them is a no-op.
- **`list_emails` InefficientFilter rejection (MEDIUM).** The v0.2.0 commit removed the workaround that dropped `$orderby` when a `from` filter was present. That removal was correct for sorting integrity, but on stricter tenants Graph rejects `startsWith(from/…, …)` + `$orderby receivedDateTime desc` as InefficientFilter when the filter doesn't also touch `receivedDateTime`. `buildFilter` now injects a permissive `receivedDateTime ge 1900-01-01T00:00:00Z` baseline when `from` is set without a user-supplied `since`. Caller-supplied `since` already satisfies the requirement and the baseline is skipped.
- **`update_event` notification field was misleading (LOW).** Response previously returned `notificationsSent: <attendees-were-replaced>`. That gave a false negative whenever a subject/time/location change on a meeting with existing attendees was made WITHOUT touching the attendees field — Graph sends notifications for those changes too. Field renamed to `attendeesReplaced` (which is what we actually know) plus a `notificationsNote` explaining Graph's auto-notify behavior. No more silent-update illusion.

### Tests

- 100 cases (up from 85). +4 datetime cases for `CreateEventInput` and `UpdateEventInput` (schema-level naive vs. offset acceptance). +10 cases in a new `test-build-filter.mjs` covering the InefficientFilter baseline matrix, OData quote escaping, clause order stability.

## [0.2.0] - 2026-05-19

### Added — new tools

- **`mark_email_read`** — mark a message as read or unread (PATCH `/me/messages/{id}`).
- **`move_email`** — move a message into another folder (well-known name or folder ID from `list_folders`). Returns the new Graph message ID, since a move re-issues it.
- **`list_folders`** — enumerate mail folders, flat per call. With `parentFolderId` set, returns immediate children for one-level-deep traversal.
- **`update_event`** — partial update of a calendar event. Sparse merge on most fields; attendees field, if present, *replaces* the entire attendee list. Graph emails notifications to affected attendees automatically.
- **`delete_event`** — remove a calendar event. Destructive: for meetings with attendees, Graph sends a cancellation notice before removal. No silent-delete mode at the API level.

### Added — response shape

- `truncated: boolean` on every list / search response (`list_emails`, `search_emails`, `list_folders`, `list_calendar_events`). Lets the LLM detect when to widen `limit` or narrow filters without an extra `$count` round-trip.

### Fixed — concurrency & correctness

- **Auth retry race.** `withAuthRetry` previously used a module-level `_forceFreshToken` flag that two concurrent tool calls could clobber — call A's retry could set the flag, call B could consume it, leaving A's retry without a fresh token. Replaced with direct `getAccessToken({ forceRefresh: true })` on retry, plus an in-flight dedup so two simultaneous 401s share one refresh request (avoids `invalid_grant` from refresh-token rotation when two refreshes fire in parallel).
- **401 detection.** `withAuthRetry` now matches auth failures on both `statusCode === 401` *and* `code === "InvalidAuthenticationToken"`, matching `explainGraphError`. Previously only `statusCode` was checked, so some SDK error shapes skipped the retry entirely.
- **`list_emails` ordering.** Removed the `useServerOrderBy = !args.from` workaround. The old code dropped server `$orderby` whenever a sender filter was set and sorted client-side — which gave the wrong result for paginated requests (the server returned an arbitrary `top` matches, not the newest). Server-side `orderby + filter` now runs unconditionally.
- **`save_attachment` integrity check.** Decoded byte count is now compared against Graph's reported `size`. Node's base64 decoder silently drops invalid characters; the old code would write a truncated file and report success. Mismatches now fail loudly with an actionable hint.
- **`search_emails` MSA query escape.** `escapeMsaSearchQuery` now strips backslashes from user input before KQL-quoting. Raw `\"` in the input would otherwise produce `\\"` on the wire, which KQL parses as "escaped backslash + string-end" — a Graph 400 with a confusing message. Backslashes are practically absent from real email search queries; replacing them with space loses zero expressivity and guarantees a well-formed quoted KQL string.
- **`create_event` / `update_event` attendee dedup tiebreaker.** When the same address appears multiple times with different types, the higher-priority type now wins (`required > optional > resource`) instead of last-write. Prevents silent downgrade of a `required` attendee to `optional` when the LLM merges multiple sources.
- **Server version drift.** Server now reads its advertised MCP version from `package.json` at startup. Previously hardcoded to `0.1.0`.
- **Account-type cache invalidation on relogin.** `runLogin` now resets the cached MSA/AAD classifier so a re-login with a different account type doesn't route `search_emails` to the wrong backend.
- **`COM0`/`LPT0` reserved-name check.** Added to the Windows-reserved-name set in `sanitizeFilename`. Modern Windows blocks these as filenames; the old list omitted them.
- **`targetFilename` schema refinement.** `save_attachment`'s `targetFilename` now rejects path separators at the schema level rather than silently rewriting them to `_`. The doc and the behavior now match.

### Changed — internals

- `normalizeAttendees` extracted from `create_event` and reused by `update_event`. Tiebreaker logic lives in one place.
- `escapeMsaSearchQuery` extracted from `search_emails` and exported for unit testing.
- `withAuthRetry`'s in-flight refresh dedup is module-local; one refresh per process at a time.

### Removed

- `wellKnownName` from `list_folders`'s `$select`. That property only exists on the `/beta` endpoint; selecting it on `/v1.0` risks a 400 on stricter tenants. The tool description now spells out that well-known folder names (e.g. `archive`, `inbox`) can be passed directly to other tools without calling `list_folders` first.

### Documentation

- Trust narrative updated: `Mail.ReadWrite` is broader than "create drafts" — it also covers mark/move/delete. The server exposes mark/move only; `move_email → deleteditems` is the supported "delete" path (recoverable from trash).
- `Mail.Send` remains deliberately absent. The trust anchor is unchanged.

### Tests

- 85 unit tests (up from 47) across seven files: path safety incl. COM0/LPT0, TZ normalization, account classification, MSA search escaping (incl. KQL backslash hazard), attendee tiebreaker (11 cases), schema validation (22 cases), MCP stdio smoke test.

## [0.1.1] - 2026-05-18

- Fix README 404s on npm; add `repository`, `bugs`, `author` metadata.
- Fix `bin` path (drop leading `./`).

## [0.1.0] - 2026-05-18

- Initial release. Seven tools: `list_emails`, `read_email`, `search_emails`, `save_attachment`, `list_calendar_events`, `create_event`, `create_draft`.
