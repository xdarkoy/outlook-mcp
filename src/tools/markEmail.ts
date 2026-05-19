import { MarkEmailInput } from "../schemas.js";
import { graph, withAuthRetry } from "../graph/client.js";
import { explainGraphError } from "../util/errors.js";
import { fail, ok, type ToolDef } from "./types.js";

/**
 * mark_email_read: toggle the read/unread state of a single message.
 *
 * PATCH /me/messages/{id} with { isRead: <bool> }. Graph returns the
 * updated message; we surface only the few fields needed to confirm the
 * write, so the LLM doesn't get a full message dump.
 *
 * Scope check: Mail.ReadWrite covers this — the same scope that gates
 * create_draft. No new permission is needed.
 */

interface GraphMessagePatchResponse {
  id?: string;
  isRead?: boolean;
  subject?: string | null;
}

export const markEmailTool: ToolDef<typeof MarkEmailInput> = {
  name: "mark_email_read",
  description:
    "Mark a message as read or unread. Returns the updated read state. Use list_emails to find candidates; " +
    "the messageId is the same opaque Graph ID returned by list_emails / search_emails / read_email.",
  schema: MarkEmailInput,
  async handler(args) {
    const mid = encodeURIComponent(args.messageId);
    try {
      const updated = (await withAuthRetry(() =>
        graph().api(`/me/messages/${mid}`).patch({ isRead: args.read }),
      )) as GraphMessagePatchResponse;

      return ok({
        updated: true,
        id: updated.id ?? args.messageId,
        subject: updated.subject ?? null,
        read: updated.isRead ?? args.read,
      });
    } catch (err) {
      const e = explainGraphError(err);
      return fail(e.message, e.hint);
    }
  },
};
