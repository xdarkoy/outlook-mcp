import { z } from "zod";

/**
 * Zod is the single source of truth: each schema defines both the runtime
 * validator and the TypeScript type, and is converted to JSON Schema for MCP.
 * Keep descriptions rich — LLMs read them as tool guidance.
 */

export const ListEmailsInput = z.object({
  folder: z
    .string()
    .optional()
    .describe(
      "Mail folder to list. Common values: 'inbox', 'sentitems', 'drafts', 'deleteditems'. Defaults to 'inbox'."
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
    .optional()
    .describe(
      "Optional filename override. If omitted, the original attachment filename is used. Must NOT contain path separators — only the bare filename. The destination directory is fixed by the server's OUTLOOK_MCP_ALLOWED_DIR setting."
    ),
});
export type SaveAttachmentInput = z.infer<typeof SaveAttachmentInput>;

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
  start: z.string().datetime({ offset: true }).describe("Start time in ISO-8601."),
  end: z.string().datetime({ offset: true }).describe("End time in ISO-8601."),
  timeZone: z
    .string()
    .optional()
    .describe("IANA time zone (e.g. 'Europe/Berlin'). Defaults to 'UTC' if omitted."),
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
