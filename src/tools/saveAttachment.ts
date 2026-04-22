import { promises as fs } from "node:fs";
import { SaveAttachmentInput } from "../schemas.js";
import { graph, withAuthRetry } from "../graph/client.js";
import { explainGraphError, ToolError } from "../util/errors.js";
import { resolveWritePath, suffixedPath, allowedDir } from "../util/paths.js";
import { fail, ok, type ToolDef } from "./types.js";

/**
 * save_attachment: write one attachment to the local allowed directory.
 *
 * This is the USP vs. any pure-cloud assistant: invoices, contracts, PDFs
 * end up as real files on disk that the user can drag, open, attach.
 *
 * Trust model:
 *   - Destination root is fixed at server start via OUTLOOK_MCP_ALLOWED_DIR
 *     and realpath()-resolved to defeat symlink escapes.
 *   - The filename from the LLM passes through sanitizeFilename +
 *     resolveWritePath (which enforces the prefix check).
 *   - Writes use O_CREAT|O_EXCL so two parallel saves of the same filename
 *     cannot overwrite each other — a TOCTOU-safe replacement for the old
 *     probe-then-write pattern.
 *   - Size cap via OUTLOOK_MCP_MAX_ATTACHMENT_MB (default 50) prevents an
 *     LLM-triggered OOM on a 100+ MB attachment.
 *   - Only fileAttachments are written. itemAttachments (embedded .eml /
 *     .vcf) and referenceAttachments (SharePoint links) are rejected.
 */

interface FileAttachmentResponse {
  "@odata.type"?: string;
  id?: string;
  name?: string;
  contentType?: string;
  size?: number;
  contentBytes?: string; // base64
}

const FILE_ATTACHMENT_TYPE = "#microsoft.graph.fileAttachment";

function maxAttachmentBytes(): number {
  const raw = process.env.OUTLOOK_MCP_MAX_ATTACHMENT_MB;
  const n = raw ? Number(raw) : 50;
  const mb = Number.isFinite(n) && n > 0 ? n : 50;
  return Math.floor(mb * 1024 * 1024);
}

/**
 * Atomic, non-clobbering write: try O_CREAT|O_EXCL, on EEXIST rename the
 * target with an incremented suffix and retry. Up to 1000 attempts.
 */
async function writeUnique(basePath: string, data: Buffer): Promise<string> {
  let attempt = 0;
  let candidate = basePath;
  while (attempt < 1000) {
    try {
      const handle = await fs.open(candidate, "wx");
      try {
        await handle.writeFile(data);
      } finally {
        await handle.close();
      }
      return candidate;
    } catch (err) {
      const code = (err as NodeJS.ErrnoException).code;
      if (code !== "EEXIST") throw err;
      attempt += 1;
      candidate = suffixedPath(basePath, attempt + 1);
    }
  }
  throw new ToolError("Too many name collisions while saving; giving up.");
}

export const saveAttachmentTool: ToolDef<typeof SaveAttachmentInput> = {
  name: "save_attachment",
  description:
    "Download one attachment from a message and save it to the local filesystem. Writes are confined " +
    "to the server's allowed directory (default ~/Downloads/outlook-mcp/). Existing files are never " +
    "overwritten — the tool appends ' (2)', ' (3)', … to the filename. Returns the absolute path of the saved file.",
  schema: SaveAttachmentInput,
  async handler(args) {
    // Encode opaque Graph IDs: they may contain / + =.
    const mid = encodeURIComponent(args.messageId);
    const aid = encodeURIComponent(args.attachmentId);
    try {
      const att = (await withAuthRetry(() =>
        graph().api(`/me/messages/${mid}/attachments/${aid}`).get(),
      )) as FileAttachmentResponse;

      const kind = att["@odata.type"];
      if (kind !== FILE_ATTACHMENT_TYPE) {
        return fail(
          `Attachment is of type '${kind ?? "unknown"}', which is not a regular file attachment.`,
          "Only fileAttachments can be saved. Item attachments (embedded messages) and reference attachments (cloud links) are not supported in this version.",
        );
      }
      if (!att.contentBytes) {
        return fail(
          "Graph returned no content bytes for this attachment.",
          "This is unusual for a fileAttachment — the attachment may be corrupt or revoked. Large attachments (streaming via /$value) are a v0.2 item.",
        );
      }

      // Size cap BEFORE decoding: stop an oversized base64 string from
      // expanding into a huge Buffer in memory.
      const maxBytes = maxAttachmentBytes();
      const approxBytes = Math.floor((att.contentBytes.length * 3) / 4);
      if (approxBytes > maxBytes) {
        return fail(
          `Attachment size (~${Math.round(approxBytes / 1024 / 1024)} MB) exceeds the configured limit (${Math.round(
            maxBytes / 1024 / 1024,
          )} MB).`,
          "Raise the limit via OUTLOOK_MCP_MAX_ATTACHMENT_MB if this is expected, or save it manually from Outlook for one-off large files.",
        );
      }

      const requestedName = args.targetFilename ?? att.name ?? "attachment.bin";
      const targetPath = await resolveWritePath(requestedName);

      const buf = Buffer.from(att.contentBytes, "base64");
      const finalPath = await writeUnique(targetPath, buf);

      return ok({
        saved: true,
        path: finalPath,
        bytes: buf.byteLength,
        contentType: att.contentType ?? null,
        originalName: att.name ?? null,
        allowedDir: allowedDir(),
      });
    } catch (err) {
      if (err instanceof ToolError) return fail(err.message, err.hint);
      const e = explainGraphError(err);
      return fail(e.message, e.hint);
    }
  },
};
