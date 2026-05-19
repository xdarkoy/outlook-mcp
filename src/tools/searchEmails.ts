import { SearchEmailsInput } from "../schemas.js";
import { graph, withAuthRetry } from "../graph/client.js";
import { getAccountType } from "../auth/accountType.js";
import { explainGraphError } from "../util/errors.js";
import { fail, ok, type ToolDef } from "./types.js";

/**
 * search_emails: full-text search across the user's mailbox.
 *
 * Two backends, chosen at runtime based on the signed-in account's tenant:
 *
 *   - AAD (work/school) → POST /search/query with entityTypes=["message"].
 *     This is the Microsoft Search / Substrate endpoint. Richer ranking,
 *     supports aggregations, returns hit highlights.
 *
 *   - MSA (personal — hotmail.com, outlook.com, live.com) → GET
 *     /me/messages?$search="…". Microsoft Search is not provisioned for
 *     consumer tenants and POST /search/query returns a hard error. The
 *     $search query option uses a different backend (the classic Outlook
 *     search stack) which IS available for MSA and supports the same
 *     KQL-like syntax (from:, subject:, hasAttachment:, "plain words"...).
 *     Caveats on MSA: cannot combine with $filter, no $orderby, max 1000
 *     results, no total count.
 *
 * Tool surface stays identical — callers pass the same query string and
 * get the same shape of results back. Account detection happens via the
 * MSAL-cached account's `tenantId` (NOT a JWT claim — MSA-issued Graph
 * access tokens are opaque, not JWTs, so decoding `tid` would fail for
 * exactly the case we need to detect). The classifier is in auth/accountType.ts
 * and is cached per process.
 */

interface GraphSearchHit {
  hitId: string;
  resource?: {
    id?: string;
    subject?: string;
    from?: { emailAddress?: { address?: string; name?: string } };
    receivedDateTime?: string;
    hasAttachments?: boolean;
    bodyPreview?: string;
  };
}

interface GraphSearchResponse {
  value?: Array<{
    hitsContainers?: Array<{
      hits?: GraphSearchHit[];
      total?: number | string;
    }>;
  }>;
}

interface GraphMessageSummary {
  id?: string;
  subject?: string;
  from?: { emailAddress?: { address?: string; name?: string } };
  receivedDateTime?: string;
  hasAttachments?: boolean;
  bodyPreview?: string;
}

interface HitSummary {
  id: string | undefined;
  subject: string;
  from: string | null;
  fromName: string | null;
  received: string | null;
  hasAttachments: boolean;
  preview: string;
}

function summarizeFromSearchHit(h: GraphSearchHit): HitSummary {
  return {
    id: h.resource?.id ?? h.hitId,
    subject: h.resource?.subject ?? "(no subject)",
    from: h.resource?.from?.emailAddress?.address ?? null,
    fromName: h.resource?.from?.emailAddress?.name ?? null,
    received: h.resource?.receivedDateTime ?? null,
    hasAttachments: !!h.resource?.hasAttachments,
    preview: (h.resource?.bodyPreview ?? "").slice(0, 240),
  };
}

/**
 * Wrap a user query for Graph's `/me/messages?$search="…"` syntax.
 *
 * Microsoft documents Graph's $search as KQL-flavored. KQL uses backslash
 * as the escape character: `\"` is a literal quote, `\\` is a literal
 * backslash. That makes raw user backslashes a hazard:
 *
 *   Input  →  naive escape  →  Graph parses as
 *   a\"b   →  "a\\"b"        →  "a\\" + b"   (string ends after \\)
 *
 * The pathological case is rare in real email-search queries (users don't
 * type backslashes), but the cost of getting it wrong is a confusing 400
 * from Graph. Pragmatic fix: strip backslashes entirely (replace with
 * space). That loses zero realistic search expressivity and guarantees a
 * well-formed quoted KQL string.
 *
 * Control characters are also stripped because the wire form is URL-
 * encoded into a request line and stray control bytes either get rejected
 * by the SDK or split the request.
 *
 * Exported for unit testing — the actual call site is the MSA branch below.
 */
export function escapeMsaSearchQuery(raw: string): string {
  // eslint-disable-next-line no-control-regex
  const stripped = raw.replace(/[\x00-\x1f\x7f\\]/g, " ");
  const escaped = stripped.replace(/"/g, '\\"');
  return `"${escaped}"`;
}

function summarizeFromMessage(m: GraphMessageSummary): HitSummary {
  return {
    id: m.id,
    subject: m.subject ?? "(no subject)",
    from: m.from?.emailAddress?.address ?? null,
    fromName: m.from?.emailAddress?.name ?? null,
    received: m.receivedDateTime ?? null,
    hasAttachments: !!m.hasAttachments,
    preview: (m.bodyPreview ?? "").slice(0, 240),
  };
}

export const searchEmailsTool: ToolDef<typeof SearchEmailsInput> = {
  name: "search_emails",
  description:
    "Full-text search across the user's ENTIRE mailbox (all folders, all years). Supports KQL-like " +
    "operators: 'from:acme.com', 'subject:contract', 'hasAttachment:true', 'body:Pantheon', plus " +
    "quoted phrases and AND/OR. Works on both personal (hotmail.com / outlook.com) and work/school " +
    "accounts — the tool picks the correct backend automatically. " +
    "IMPORTANT: results are ranked by relevance, NOT sorted by date. Date filters like " +
    "'received:this-week' are honored on AAD (work/school) but are IGNORED by the personal-account " +
    "search backend, which returns matches from any time period. If you need strict date filtering, " +
    "use list_emails (folder-scoped) with `since`/`until` instead. Returns a compact JSON array of " +
    "hits with message IDs usable by read_email.",
  schema: SearchEmailsInput,
  async handler(args) {
    const size = args.limit ?? 25;
    try {
      const accountType = await getAccountType();

      if (accountType === "aad") {
        // Microsoft Search (Substrate) — richer, but AAD only.
        const res = (await withAuthRetry(() =>
          graph()
            .api("/search/query")
            .post({
              requests: [
                {
                  entityTypes: ["message"],
                  query: { queryString: args.query },
                  from: 0,
                  size,
                },
              ],
            }),
        )) as GraphSearchResponse;

        const hits = res.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
        const rawTotal = res.value?.[0]?.hitsContainers?.[0]?.total;
        const coerced = typeof rawTotal === "string" ? Number(rawTotal) : rawTotal;
        const total = Number.isFinite(coerced) ? (coerced as number) : hits.length;

        return ok({
          backend: "microsoft-search",
          query: args.query,
          total,
          returned: hits.length,
          // truncated = there are more results than we returned. Helps the
          // LLM decide whether to ask the user for a narrower query.
          truncated: total > hits.length,
          messages: hits.map(summarizeFromSearchHit),
        });
      }

      // MSA — use /me/messages?$search=. The $search value must be quoted.
      const quoted = escapeMsaSearchQuery(args.query);
      const res = (await withAuthRetry(() =>
        graph()
          .api("/me/messages")
          .top(size)
          .search(quoted)
          .select([
            "id",
            "subject",
            "from",
            "receivedDateTime",
            "hasAttachments",
            "bodyPreview",
          ])
          .get(),
      )) as { value?: GraphMessageSummary[] };

      const items = res.value ?? [];
      return ok({
        backend: "me-messages-search",
        query: args.query,
        // MSA backend does not return a total count.
        total: null,
        returned: items.length,
        // We can't know definitively without a total. Hint the LLM only when
        // we hit the requested cap — that's the case where it likely matters.
        truncated: items.length >= size,
        messages: items.map(summarizeFromMessage),
      });
    } catch (err) {
      const e = explainGraphError(err);
      return fail(e.message, e.hint);
    }
  },
};
