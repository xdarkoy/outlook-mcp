import { MoveEmailInput } from "../schemas.js";
import { graph, withAuthRetry } from "../graph/client.js";
import { explainGraphError } from "../util/errors.js";
import { fail, ok, type ToolDef } from "./types.js";

/**
 * move_email: relocate a message into another mail folder.
 *
 * POST /me/messages/{id}/move with { destinationId: "<folder>" }. Graph's
 * `destinationId` accepts EITHER a folder ID (from list_folders) OR a
 * well-known folder name (inbox, archive, deleteditems, …) — same union
 * as everywhere else in this codebase.
 *
 * IMPORTANT: a move produces a NEW Graph message ID. The original ID is
 * invalidated. We return both so the caller can update any cached references.
 *
 * Soft-delete vs hard-delete: moving to "deleteditems" is the standard
 * Outlook "Delete" gesture (recoverable from the trash for ~30 days). We
 * do NOT expose a permanent /delete endpoint — by design, every destructive
 * action through this server is reversible via the Outlook UI.
 */

interface GraphMessageMoveResponse {
  id?: string;
  subject?: string | null;
  parentFolderId?: string;
}

export const moveEmailTool: ToolDef<typeof MoveEmailInput> = {
  name: "move_email",
  description:
    "Move a message into a different mail folder. The destination is either a well-known folder name " +
    "(inbox, drafts, sentitems, deleteditems, junkemail, archive, outbox, conversationhistory) or a folder " +
    "ID from list_folders. Returns the NEW message ID — the original ID becomes invalid after the move. " +
    "To delete a message in the Outlook sense, move it to 'deleteditems' (recoverable from trash).",
  schema: MoveEmailInput,
  async handler(args) {
    const mid = encodeURIComponent(args.messageId);
    try {
      const moved = (await withAuthRetry(() =>
        graph()
          .api(`/me/messages/${mid}/move`)
          .post({ destinationId: args.destinationFolder }),
      )) as GraphMessageMoveResponse;

      return ok({
        moved: true,
        oldId: args.messageId,
        newId: moved.id ?? null,
        subject: moved.subject ?? null,
        destinationFolder: args.destinationFolder,
        // Echo back the resolved parent ID so the caller can tell whether
        // a well-known alias landed where they expected.
        parentFolderId: moved.parentFolderId ?? null,
      });
    } catch (err) {
      const e = explainGraphError(err);
      return fail(e.message, e.hint);
    }
  },
};
