import { z } from "zod";

/**
 * Zod is the single source of truth: each schema defines both the runtime
 * validator and the TypeScript type, and is converted to JSON Schema for MCP.
 * Keep descriptions rich — LLMs read them as tool guidance.
 */

/**
 * Microsoft Graph well-known mail folder names. These can be passed in any
 * `folder` field directly — Graph resolves them server-side. Anything else
 * must be a folder ID from list_folders.
 *
 * `archive` is a tenant-optional folder (always present on personal accounts,
 * created on first use on AAD tenants); we list it because Outlook clients
 * surface it as a primary action.
 */
export const WELLKNOWN_MAIL_FOLDERS = [
  "inbox",
  "drafts",
  "sentitems",
  "deleteditems",
  "junkemail",
  "archive",
  "outbox",
  "conversationhistory",
] as const;

const FOLDER_DESCRIPTION =
  "Mail folder. Either a well-known name (inbox, drafts, sentitems, deleteditems, junkemail, archive, outbox, conversationhistory) or a folder ID returned by list_folders. Unknown names produce a 404 from Graph — use list_folders to discover custom folders.";

export const ListEmailsInput = z.object({
  folder: z
    .string()
    .min(1)
    .optional()
    .describe(
      "Mail folder to list. Defaults to 'inbox'. Either a well-known name (inbox, drafts, sentitems, deleteditems, junkemail, archive, outbox, conversationhistory) or a folder ID returned by list_folders. Unknown names produce a 404 from Graph — use list_folders to discover custom folders.",
    ),
  from: z
    .string()
    .optional()
    .describe(
      "Filter: only emails whose sender address STARTS WITH this value (case-sensitive, Graph OData limitation). For a full domain match use 'user@acme.com' or 'acme.com'. For fuzzy / case-insensitive matching use the search_emails tool instead."
    ),
  since: z
    .string()
    .datetime({ offset: true })
    .optional()
    .describe("Filter: only emails received at or after this ISO-8601 timestamp (e.g. 2026-04-15T00:00:00Z)."),
  until: z
    .string()
    .datetime({ offset: true })
    .optional()
    .describe("Filter: only emails received strictly before this ISO-8601 timestamp."),
  unreadOnly: z
    .boolean()
    .optional()
    .describe("If true, return only unread messages."),
  limit: z
    .number()
    .int()
    .min(1)
    .max(100)
    .optional()
    .describe("Maximum number of emails to return. Default 25, max 100."),
});
export type ListEmailsInput = z.infer<typeof ListEmailsInput>;

export const ReadEmailInput = z.object({
  messageId: z.string().min(1).describe("The Graph message ID (returned by list_emails or search_emails)."),
  includeBody: z
    .boolean()
    .optional()
    .describe("If true, include the full body (text preferred over HTML). Default true."),
});
export type ReadEmailInput = z.infer<typeof ReadEmailInput>;

export const SearchEmailsInput = z.object({
  query: z
    .string()
    .min(1)
    .describe(
      "Full-text search query using Microsoft Graph search syntax. Examples: 'invoice from:acme.com', 'subject:contract received:this week', 'hasAttachment:true Q4 report'."
    ),
  limit: z
    .number()
    .int()
    .min(1)
    .max(250)
    .optional()
    .describe("Maximum number of results. Default 25, max 250."),
});
export type SearchEmailsInput = z.infer<typeof SearchEmailsInput>;

export const SaveAttachmentInput = z.object({
  messageId: z.string().min(1).describe("The Graph message ID that contains the attachment."),
  attachmentId: z
    .string()
    .min(1)
    .describe("The Graph attachment ID (from read_email's 'attachments' list)."),
  targetFilename: z
    .string()
    .min(1)
    // Hard reject on separators rather than silently rewriting them to '_'.
    // sanitizeFilename further inside the writer is defense-in-depth; this
    // refinement gives the LLM a clear schema-level rejection so it doesn't
    // build wrong expectations about what filename will land on disk.
    .refine((s) => !/[\\/]/.test(s), {
      message: "targetFilename must not contain path separators ('/' or '\\\\'). Use only a bare filename.",
    })
    .optional()
    .describe(
      "Optional filename override. If omitted, the original attachment filename is used. Must be a bare filename — no path separators, no parent traversal. The destination directory is fixed by the server's OUTLOOK_MCP_ALLOWED_DIR setting and is not LLM-controllable.",
    ),
});
export type SaveAttachmentInput = z.infer<typeof SaveAttachmentInput>;

export const MarkEmailInput = z.object({
  messageId: z.string().min(1).describe("The Graph message ID to update."),
  read: z
    .boolean()
    .describe("true to mark the message as read, false to mark it as unread."),
});
export type MarkEmailInput = z.infer<typeof MarkEmailInput>;

export const MoveEmailInput = z.object({
  messageId: z.string().min(1).describe("The Graph message ID to move."),
  destinationFolder: z
    .string()
    .min(1)
    .describe(
      "Destination " + FOLDER_DESCRIPTION.charAt(0).toLowerCase() + FOLDER_DESCRIPTION.slice(1),
    ),
});
export type MoveEmailInput = z.infer<typeof MoveEmailInput>;

export const ListFoldersInput = z.object({
  parentFolderId: z
    .string()
    .min(1)
    .optional()
    .describe(
      "If set, list the immediate children of this folder (subfolders). If omitted, list top-level folders in the user's mailbox.",
    ),
  limit: z
    .number()
    .int()
    .min(1)
    .max(200)
    .optional()
    .describe("Maximum number of folders to return. Default 100, max 200."),
});
export type ListFoldersInput = z.infer<typeof ListFoldersInput>;

export const ListCalendarEventsInput = z.object({
  from: z
    .string()
    .datetime({ offset: true })
    .describe("Start of the time window (ISO-8601). E.g. '2026-04-22T00:00:00Z'."),
  to: z
    .string()
    .datetime({ offset: true })
    .describe("End of the time window (ISO-8601, exclusive)."),
  limit: z
    .number()
    .int()
    .min(1)
    .max(100)
    .optional()
    .describe("Max events to return. Default 50."),
});
export type ListCalendarEventsInput = z.infer<typeof ListCalendarEventsInput>;

/**
 * Attendee item: either a bare email string (shortcut for required) or a
 * {address, type} object. `resource` is for meeting rooms / equipment.
 */
const AttendeeInput = z.union([
  z.string().email(),
  z.object({
    address: z.string().email(),
    type: z.enum(["required", "optional", "resource"]).optional(),
  }),
]);
export type AttendeeInput = z.infer<typeof AttendeeInput>;

export const CreateEventInput = z.object({
  subject: z.string().min(1).describe("Event subject / title."),
  // `local: true` permits naive (no offset, no Z) ISO strings — the documented
  // way to pass times in a non-UTC IANA zone. `offset: true` keeps Z and
  // +HH:MM suffixes valid for UTC use. Without `local`, naive forms get
  // rejected by Zod before the handler ever sees them, which makes the
  // "naive + Europe/Berlin" path advertised by timeZone unusable.
  start: z
    .string()
    .datetime({ offset: true, local: true })
    .describe("Start time in ISO-8601. Naive (no offset) is allowed and uses timeZone; offset suffix (Z or +HH:MM) is only allowed when timeZone is UTC (or omitted)."),
  end: z
    .string()
    .datetime({ offset: true, local: true })
    .describe("End time in ISO-8601. Same rules as start."),
  timeZone: z
    .string()
    .optional()
    .describe("IANA time zone (e.g. 'Europe/Berlin'). Defaults to 'UTC' if omitted. For non-UTC zones, pass naive datetimes (no Z/offset suffix)."),
  attendees: z
    .array(AttendeeInput)
    .optional()
    .describe(
      "Attendees to invite. Each entry is either a bare email (treated as 'required') or " +
        "{address, type} where type is 'required' | 'optional' | 'resource'. Use 'resource' for meeting rooms / equipment.",
    ),
  body: z.string().optional().describe("Event body / description (plain text)."),
  location: z
    .string()
    .max(256)
    .optional()
    .describe("Free-text location or meeting room. Graph truncates above 256 chars."),
});
export type CreateEventInput = z.infer<typeof CreateEventInput>;

/**
 * update_event: PATCH /me/events/{id}. All fields besides eventId are optional;
 * only provided fields are sent to Graph.
 *
 * Attendees semantics: if `attendees` is provided, the entire attendee list is
 * REPLACED — Graph does not merge. Graph will email update notifications to
 * affected attendees automatically (added, removed, or whose times changed).
 */
export const UpdateEventInput = z.object({
  eventId: z.string().min(1).describe("The Graph event ID to update."),
  subject: z.string().min(1).optional().describe("New event subject / title."),
  start: z
    .string()
    .datetime({ offset: true, local: true })
    .optional()
    .describe("New start time in ISO-8601. Naive (no offset) is allowed and uses timeZone; offset suffix is only allowed when timeZone is UTC."),
  end: z
    .string()
    .datetime({ offset: true, local: true })
    .optional()
    .describe("New end time in ISO-8601. Same rules as start."),
  timeZone: z
    .string()
    .optional()
    .describe(
      "IANA time zone (e.g. 'Europe/Berlin') for the start/end datetimes. Required if start or end uses a naive (no-offset) ISO string. Defaults to 'UTC'.",
    ),
  attendees: z
    .array(AttendeeInput)
    .optional()
    .describe(
      "REPLACES the entire attendee list. Each entry is either a bare email (treated as 'required') or {address, type}. To keep existing attendees, fetch the event first, then pass the full desired list. Graph will email updates to affected attendees automatically.",
    ),
  body: z.string().optional().describe("New event body / description (plain text)."),
  location: z
    .string()
    .max(256)
    .optional()
    .describe("New free-text location."),
});
export type UpdateEventInput = z.infer<typeof UpdateEventInput>;

/**
 * delete_event: DELETE /me/events/{id}. Graph behavior:
 *   - Solo events: gone.
 *   - Meetings with attendees on a sent invite: Graph sends a cancellation
 *     notice to all attendees, then removes the event. There is no "delete
 *     without notifying" mode at the API level.
 *
 * Callers should make the destructive nature explicit to the user before
 * invocation. The server does not gate the call.
 */
export const DeleteEventInput = z.object({
  eventId: z.string().min(1).describe("The Graph event ID to delete."),
});
export type DeleteEventInput = z.infer<typeof DeleteEventInput>;

/**
 * create_draft: By design, this MCP server CAN NOT send email directly.
 * The OAuth scope Mail.Send is deliberately NOT requested. A draft lands
 * in the user's Drafts folder and must be reviewed + sent manually in Outlook.
 * This is the trust anchor of the product.
 */
export const CreateDraftInput = z.object({
  to: z.array(z.string().email()).min(1).describe("Recipient email addresses (To:)."),
  cc: z.array(z.string().email()).optional().describe("CC recipients."),
  bcc: z.array(z.string().email()).optional().describe("BCC recipients."),
  subject: z.string().describe("Email subject."),
  body: z.string().describe("Email body. Plain text by default."),
  bodyFormat: z
    .enum(["text", "html"])
    .optional()
    .describe("Body format. Default 'text'."),
  replyToMessageId: z
    .string()
    .optional()
    .describe(
      "If set, create the draft as a reply to this message (preserves thread, quotes original)."
    ),
});
export type CreateDraftInput = z.infer<typeof CreateDraftInput>;
