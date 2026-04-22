import { ReadEmailInput } from "../schemas.js";
import { graph, withAuthRetry } from "../graph/client.js";
import { explainGraphError } from "../util/errors.js";
import { fail, ok, type ToolDef } from "./types.js";

/**
 * read_email: fetch one message with body and attachment metadata.
 *
 * We deliberately DO NOT fetch attachment content here — attachments can be
 * many megabytes and dumping them into the tool result would blow the LLM
 * context. The LLM gets IDs/filenames/sizes and can call save_attachment to
 * write the content to disk.
 *
 * Body: we prefer text over HTML. HTML is noisy for LLMs and rarely what the
 * user wants when asking "what does this email say?". The `Prefer:
 * outlook.body-content-type="text"` header makes Graph do the conversion
 * server-side — cheaper and more reliable than HTML-stripping in Node.
 */

interface GraphAttachmentSummary {
  id?: string;
  name?: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
}

interface GraphMessageFull {
  id: string;
  subject?: string | null;
  from?: { emailAddress?: { address?: string; name?: string } } | null;
  toRecipients?: Array<{ emailAddress?: { address?: string; name?: string } }>;
  ccRecipients?: Array<{ emailAddress?: { address?: string; name?: string } }>;
  receivedDateTime?: string;
  sentDateTime?: string;
  isRead?: boolean;
  hasAttachments?: boolean;
  body?: { contentType?: string; content?: string };
  bodyPreview?: string;
  internetMessageId?: string;
  conversationId?: string;
}

function flattenRecipients(
  list?: Array<{ emailAddress?: { address?: string; name?: string } }>,
): Array<{ address: string | null; name: string | null }> {
  return (list ?? []).map((r) => ({
    address: r.emailAddress?.address ?? null,
    name: r.emailAddress?.name ?? null,
  }));
}

export const readEmailTool: ToolDef<typeof ReadEmailInput> = {
  name: "read_email",
  description:
    "Fetch the full content of one email message by id, plus a list of its attachments (metadata only — " +
    "use save_attachment to write an attachment to disk). Returns JSON with body as plain text.",
  schema: ReadEmailInput,
  async handler(args) {
    const includeBody = args.includeBody ?? true;
    // Graph message IDs are opaque, URL-unsafe strings (may contain / + =).
    // The SDK's .api() does NOT encode path segments for us.
    const mid = encodeURIComponent(args.messageId);
    try {
      const msgReq = graph()
        .api(`/me/messages/${mid}`)
        .header("Prefer", 'outlook.body-content-type="text"')
        .select(
          [
            "id",
            "subject",
            "from",
            "toRecipients",
            "ccRecipients",
            "receivedDateTime",
            "sentDateTime",
            "isRead",
            "hasAttachments",
            includeBody ? "body" : "",
            "bodyPreview",
            "internetMessageId",
            "conversationId",
          ].filter(Boolean),
        );

      const msg = (await withAuthRetry(() => msgReq.get())) as GraphMessageFull;

      let attachments: GraphAttachmentSummary[] = [];
      if (msg.hasAttachments) {
        const resp = (await withAuthRetry(() =>
          graph()
            .api(`/me/messages/${mid}/attachments`)
            .select(["id", "name", "contentType", "size", "isInline"])
            .get(),
        )) as { value?: GraphAttachmentSummary[] };
        attachments = resp.value ?? [];
      }

      return ok({
        id: msg.id,
        subject: msg.subject ?? "(no subject)",
        from: msg.from?.emailAddress?.address ?? null,
        fromName: msg.from?.emailAddress?.name ?? null,
        to: flattenRecipients(msg.toRecipients),
        cc: flattenRecipients(msg.ccRecipients),
        received: msg.receivedDateTime ?? null,
        sent: msg.sentDateTime ?? null,
        unread: msg.isRead === false,
        conversationId: msg.conversationId ?? null,
        internetMessageId: msg.internetMessageId ?? null,
        body: includeBody
          ? {
              contentType: msg.body?.contentType ?? "text",
              content: msg.body?.content ?? "",
            }
          : undefined,
        preview: !includeBody ? (msg.bodyPreview ?? "") : undefined,
        attachments: attachments.map((a) => ({
          id: a.id,
          name: a.name ?? "(unnamed)",
          contentType: a.contentType ?? null,
          size: a.size ?? null,
          inline: !!a.isInline,
        })),
      });
    } catch (err) {
      const e = explainGraphError(err);
      return fail(e.message, e.hint);
    }
  },
};
