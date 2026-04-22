import { CreateDraftInput } from "../schemas.js";
import { graph, withAuthRetry } from "../graph/client.js";
import { explainGraphError } from "../util/errors.js";
import { fail, ok, type ToolDef } from "./types.js";

/**
 * create_draft: create (never send!) a draft email.
 *
 * This is the product's trust anchor. Two code paths:
 *
 *   1. Standalone draft — POST /me/messages with toRecipients/body/subject.
 *      Graph creates it in the Drafts folder automatically; `isDraft` is a
 *      server-assigned, read-only property — do NOT set it in the payload
 *      or Graph will 400.
 *
 *   2. Reply draft — POST /me/messages/{id}/createReply with a {message:{…}}
 *      body containing our overrides. Graph merges them into the server-built
 *      reply in a SINGLE round-trip, so there is no orphan-draft risk if our
 *      second call were to fail. If `body.content` is provided here, Graph
 *      uses it instead of the quoted original (the user got to see/choose
 *      what they're sending — that is intentional).
 *
 * In BOTH paths the resulting draft lives in Drafts until a human opens
 * Outlook and clicks Send. The OAuth token does NOT have Mail.Send scope —
 * that is the true "no-send" guarantee, not any flag in the payload. Do NOT
 * add a send path without an explicit product decision reversing this.
 */

interface GraphMessageCreated {
  id: string;
  webLink?: string;
  subject?: string;
}

function toRecipientList(addresses?: string[]) {
  return (addresses ?? []).map((addr) => ({ emailAddress: { address: addr } }));
}

export const createDraftTool: ToolDef<typeof CreateDraftInput> = {
  name: "create_draft",
  description:
    "Create a draft email in the user's Drafts folder — NEVER sends. If replyToMessageId is set, " +
    "the draft is created as a threaded reply; if `body` is provided it replaces Graph's quoted " +
    "original so the caller fully controls the outgoing text. Otherwise a new standalone draft is " +
    "created. The user must review and send manually in Outlook; this server has no ability to send.",
  schema: CreateDraftInput,
  async handler(args) {
    const format = args.bodyFormat ?? "text";
    const body = { contentType: format, content: args.body };

    try {
      let draftId: string;
      let webLink: string | null = null;
      let subject = args.subject;

      if (args.replyToMessageId) {
        // Single atomic round-trip: /createReply accepts a `message` property
        // whose fields override the server-built draft. One request, no
        // orphan on failure.
        const replyTargetId = encodeURIComponent(args.replyToMessageId);

        const messageOverrides: Record<string, unknown> = { body };
        if (args.subject) messageOverrides.subject = args.subject;
        if (args.to !== undefined) messageOverrides.toRecipients = toRecipientList(args.to);
        if (args.cc !== undefined) messageOverrides.ccRecipients = toRecipientList(args.cc);
        if (args.bcc !== undefined) messageOverrides.bccRecipients = toRecipientList(args.bcc);

        const reply = (await withAuthRetry(() =>
          graph()
            .api(`/me/messages/${replyTargetId}/createReply`)
            .post({ message: messageOverrides }),
        )) as GraphMessageCreated;

        draftId = reply.id;
        webLink = reply.webLink ?? null;
        subject = reply.subject ?? subject;
      } else {
        // POST /me/messages creates a draft implicitly; the `isDraft` field
        // is server-assigned and read-only — including it 400s.
        const payload: Record<string, unknown> = {
          subject: args.subject,
          body,
          toRecipients: toRecipientList(args.to),
        };
        if (args.cc && args.cc.length) payload.ccRecipients = toRecipientList(args.cc);
        if (args.bcc && args.bcc.length) payload.bccRecipients = toRecipientList(args.bcc);

        const created = (await withAuthRetry(() =>
          graph().api("/me/messages").post(payload),
        )) as GraphMessageCreated;
        draftId = created.id;
        webLink = created.webLink ?? null;
        subject = created.subject ?? subject;
      }

      return ok({
        drafted: true,
        sent: false,
        id: draftId,
        subject,
        webLink,
        note:
          "Draft saved. This server cannot send — open the draft in Outlook to review and send manually.",
      });
    } catch (err) {
      const e = explainGraphError(err);
      return fail(e.message, e.hint);
    }
  },
};
