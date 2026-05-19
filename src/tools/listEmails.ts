import { ListEmailsInput } from "../schemas.js";
import { graph, withAuthRetry } from "../graph/client.js";
import { explainGraphError } from "../util/errors.js";
import { fail, ok, type ToolDef } from "./types.js";

/**
 * list_emails: filtered listing of messages in a folder.
 *
 * Uses Graph $filter for sender/date/unread and $orderby for newest-first.
 * Returns a compact JSON summary — just enough for the LLM to pick which
 * message to read_email or save_attachment from.
 */

interface GraphMessageSummary {
  id: string;
  subject?: string | null;
  from?: { emailAddress?: { address?: string; name?: string } } | null;
  receivedDateTime?: string;
  isRead?: boolean;
  hasAttachments?: boolean;
  bodyPreview?: string;
}

function buildFilter(args: {
  from?: string;
  since?: string;
  until?: string;
  unreadOnly?: boolean;
}): string | undefined {
  const parts: string[] = [];
  if (args.from) {
    // Graph's $filter does NOT support tolower() on navigation-property strings,
    // and contains() on from/emailAddress/address is also unsupported — Graph
    // only allows startsWith() here. Match is case-sensitive as a result; for
    // richer matching the LLM should use search_emails (which is KQL-based and
    // case-insensitive by default).
    const esc = args.from.replace(/'/g, "''");
    parts.push(`startsWith(from/emailAddress/address, '${esc}')`);
  }
  if (args.since) parts.push(`receivedDateTime ge ${args.since}`);
  if (args.until) parts.push(`receivedDateTime lt ${args.until}`);
  if (args.unreadOnly) parts.push("isRead eq false");
  return parts.length ? parts.join(" and ") : undefined;
}

export const listEmailsTool: ToolDef<typeof ListEmailsInput> = {
  name: "list_emails",
  description:
    "List emails from a mail folder with optional filters (sender substring, date range, unread-only). " +
    "Returns a compact JSON array of message summaries. Use read_email with a returned id to fetch full content.",
  schema: ListEmailsInput,
  async handler(args) {
    const folder = args.folder ?? "inbox";
    const limit = args.limit ?? 25;
    const filter = buildFilter(args);

    try {
      // Earlier versions dropped server-side $orderby when a sender filter
      // was set, then sorted client-side — but that's WRONG for paginated
      // requests: the server returns an arbitrary `top` matches, and
      // sorting those locally gives you neither the newest matches nor a
      // deterministic page. Graph 2026 supports orderby+filter on
      // receivedDateTime + from together; we now always sort server-side.
      // If a tenant ever rejects this combo, the 400 surfaces via
      // explainGraphError as a clear message rather than silently lying.
      let req = graph()
        // The SDK URL-encodes path segments itself; passing the folder name
        // raw avoids double-encoding. Well-known names (inbox, drafts,
        // sentitems, deleteditems, junkemail, archive, outbox,
        // conversationhistory) resolve directly. For custom folders pass
        // an ID from list_folders.
        .api(`/me/mailFolders/${folder}/messages`)
        // Fetch one extra row so we can set `truncated` without an
        // expensive $count query — if Graph returns more than `limit`,
        // there's at least one more page available.
        .top(limit + 1)
        .orderby("receivedDateTime desc")
        .select([
          "id",
          "subject",
          "from",
          "receivedDateTime",
          "isRead",
          "hasAttachments",
          "bodyPreview",
        ]);
      if (filter) req = req.filter(filter);

      const res = await withAuthRetry(() => req.get());
      const raw: GraphMessageSummary[] = res?.value ?? [];
      const truncated = raw.length > limit;
      const items = truncated ? raw.slice(0, limit) : raw;

      const summary = items.map((m) => ({
        id: m.id,
        subject: m.subject ?? "(no subject)",
        from: m.from?.emailAddress?.address ?? null,
        fromName: m.from?.emailAddress?.name ?? null,
        received: m.receivedDateTime ?? null,
        unread: m.isRead === false,
        hasAttachments: !!m.hasAttachments,
        preview: (m.bodyPreview ?? "").slice(0, 240),
      }));

      return ok({
        folder,
        count: summary.length,
        // truncated: there's at least one more match beyond what we
        // returned. Hint the LLM to either raise `limit` or narrow filters.
        truncated,
        messages: summary,
      });
    } catch (err) {
      const e = explainGraphError(err);
      return fail(e.message, e.hint);
    }
  },
};
