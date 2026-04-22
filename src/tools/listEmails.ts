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
    // Graph's "inconsistent $orderby" rule: combining $filter on one property
    // with $orderby on another triggers a 400 on some endpoints. The fix is
    // to drop $orderby when we have a sender filter and sort client-side.
    const useServerOrderBy = !args.from;

    try {
      let req = graph()
        // The SDK URL-encodes path segments itself; passing the folder name
        // raw avoids double-encoding. Only well-known folder names (inbox,
        // drafts, sentitems, deleteditems, junkemail, archive) are supported
        // here — custom subfolders need an ID lookup, which is a v0.2 item.
        .api(`/me/mailFolders/${folder}/messages`)
        .top(limit)
        .select([
          "id",
          "subject",
          "from",
          "receivedDateTime",
          "isRead",
          "hasAttachments",
          "bodyPreview",
        ]);
      if (useServerOrderBy) req = req.orderby("receivedDateTime desc");
      if (filter) req = req.filter(filter);

      const res = await withAuthRetry(() => req.get());
      const items: GraphMessageSummary[] = res?.value ?? [];

      // Client-side sort when we couldn't ask the server.
      if (!useServerOrderBy) {
        items.sort((a, b) => {
          const ta = a.receivedDateTime ? Date.parse(a.receivedDateTime) : 0;
          const tb = b.receivedDateTime ? Date.parse(b.receivedDateTime) : 0;
          return tb - ta;
        });
      }

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

      return ok({ folder, count: summary.length, messages: summary });
    } catch (err) {
      const e = explainGraphError(err);
      return fail(e.message, e.hint);
    }
  },
};
