import { ListFoldersInput } from "../schemas.js";
import { graph, withAuthRetry } from "../graph/client.js";
import { explainGraphError } from "../util/errors.js";
import { fail, ok, type ToolDef } from "./types.js";

/**
 * list_folders: enumerate mail folders.
 *
 * Two modes:
 *   - parentFolderId omitted → GET /me/mailFolders (top-level folders)
 *   - parentFolderId set     → GET /me/mailFolders/{id}/childFolders
 *
 * We return a flat list per call; nested traversal is the caller's job
 * (call again with a child's id as parentFolderId). This keeps the tool
 * surface narrow and predictable, and avoids accidentally returning a
 * 10k-folder mailbox in one shot.
 *
 * NOTE on well-known folder identification: Graph v1.0's mailFolder
 * resource does NOT expose a `wellKnownName` property — that field only
 * exists on the /beta endpoint. Selecting it on v1.0 risks a 400 on
 * stricter tenants. To keep this tool reliable across both AAD and MSA,
 * we return only documented v1.0 fields. Callers that need to address a
 * well-known folder can pass its name (e.g. "archive") directly into
 * move_email or list_emails; Graph resolves it server-side.
 */

interface GraphMailFolder {
  id?: string;
  displayName?: string;
  parentFolderId?: string;
  childFolderCount?: number;
  totalItemCount?: number;
  unreadItemCount?: number;
}

export const listFoldersTool: ToolDef<typeof ListFoldersInput> = {
  name: "list_folders",
  description:
    "List mail folders. With no parentFolderId, returns top-level folders. With parentFolderId set, " +
    "returns immediate child folders (one level deep). Use the returned 'id' as a folder reference for " +
    "list_emails or move_email. For Outlook's standard folders you can also pass their well-known names " +
    "directly (inbox, drafts, sentitems, deleteditems, junkemail, archive, outbox, conversationhistory) " +
    "without calling this tool first.",
  schema: ListFoldersInput,
  async handler(args) {
    const limit = args.limit ?? 100;
    try {
      const path = args.parentFolderId
        ? `/me/mailFolders/${encodeURIComponent(args.parentFolderId)}/childFolders`
        : "/me/mailFolders";

      // Fetch one extra row to compute truncated without a separate $count.
      const res = (await withAuthRetry(() =>
        graph()
          .api(path)
          .top(limit + 1)
          .select([
            "id",
            "displayName",
            "parentFolderId",
            "childFolderCount",
            "totalItemCount",
            "unreadItemCount",
          ])
          .get(),
      )) as { value?: GraphMailFolder[] };

      const raw = res.value ?? [];
      const truncated = raw.length > limit;
      const folders = (truncated ? raw.slice(0, limit) : raw).map((f) => ({
        id: f.id ?? null,
        displayName: f.displayName ?? "(unnamed)",
        parentFolderId: f.parentFolderId ?? null,
        childFolderCount: f.childFolderCount ?? 0,
        totalItemCount: f.totalItemCount ?? 0,
        unreadItemCount: f.unreadItemCount ?? 0,
      }));

      return ok({
        parentFolderId: args.parentFolderId ?? null,
        count: folders.length,
        truncated,
        folders,
      });
    } catch (err) {
      const e = explainGraphError(err);
      return fail(e.message, e.hint);
    }
  },
};
