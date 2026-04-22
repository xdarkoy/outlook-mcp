import { CreateEventInput } from "../schemas.js";
import { graph, withAuthRetry } from "../graph/client.js";
import { explainGraphError, ToolError } from "../util/errors.js";
import { fail, ok, type ToolDef } from "./types.js";

/**
 * create_event: add a calendar event.
 *
 * Behavior:
 *   - If `attendees` is set, Graph automatically sends meeting invitations.
 *     That's by design — the LLM has already shown the user the attendee
 *     list before calling this, so the invite is an intended side effect.
 *
 * Time-zone rules:
 *   Graph's /me/events takes {dateTime, timeZone} where dateTime is a NAIVE
 *   ISO string (no offset) interpreted in the given timeZone. The earlier
 *   implementation blindly stripped any offset — which silently MISBOOKED
 *   the event when caller passed e.g. '2026-05-01T10:00:00+02:00' with
 *   timeZone 'UTC' (booked 10:00 UTC == 12:00 Berlin, two hours off).
 *
 *   New rules:
 *     - Caller ISO has no offset: use the naive string as-is with the
 *       provided timeZone (or UTC default).
 *     - Caller ISO has an offset (Z or ±hh[:mm]) AND timeZone is unset or
 *       'UTC': convert to the UTC instant and send as UTC. Correct.
 *     - Caller ISO has an offset AND timeZone is set to something else:
 *       ambiguous — reject with a clear error rather than guess.
 */

interface GraphEventCreateResponse {
  id: string;
  webLink?: string;
  start?: { dateTime?: string; timeZone?: string };
  end?: { dateTime?: string; timeZone?: string };
}

const OFFSET_RE = /(Z|[+\-]\d{2}:?\d{2})$/;

function hasOffset(iso: string): boolean {
  return OFFSET_RE.test(iso);
}

function toUtcNaive(iso: string): string {
  // Parse the offset-bearing ISO to an instant, then format as naive UTC.
  // Node's Date parser accepts all ISO-8601 variants we care about.
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) {
    throw new ToolError(`Could not parse datetime '${iso}'.`);
  }
  // 'YYYY-MM-DDTHH:mm:ss.sssZ' → strip 'Z' and the trailing '.sss' to get a
  // Graph-friendly naive string. Second-resolution is enough for calendar.
  return d.toISOString().replace(/\.\d+Z$/, "");
}

export function normalizeDateTime(iso: string, tz: string, field: "start" | "end"): string {
  const offsetPresent = hasOffset(iso);
  if (!offsetPresent) return iso;
  if (tz === "UTC") return toUtcNaive(iso);
  throw new ToolError(
    `${field} has an explicit offset ('${iso}') but timeZone is '${tz}'.`,
    "Either drop the offset from the datetime (and keep timeZone), or set timeZone to 'UTC' (and let the offset determine the instant). Do not mix both.",
  );
}

export const createEventTool: ToolDef<typeof CreateEventInput> = {
  name: "create_event",
  description:
    "Create a calendar event. If attendees are provided, Microsoft Graph will automatically send meeting invitations. " +
    "Time zone defaults to UTC if unspecified. Returns the event id and a web link to open it in Outlook.",
  schema: CreateEventInput,
  async handler(args) {
    const tz = args.timeZone ?? "UTC";

    // Reject end <= start before we even talk to Graph — the LLM gets a
    // clear message rather than a cryptic 400.
    if (new Date(args.end).getTime() <= new Date(args.start).getTime()) {
      return fail(
        `Event end (${args.end}) is not after start (${args.start}).`,
        "Swap the values or extend the end time.",
      );
    }

    let startDt: string;
    let endDt: string;
    try {
      startDt = normalizeDateTime(args.start, tz, "start");
      endDt = normalizeDateTime(args.end, tz, "end");
    } catch (err) {
      if (err instanceof ToolError) return fail(err.message, err.hint);
      throw err;
    }
    const payload: Record<string, unknown> = {
      subject: args.subject,
      start: { dateTime: startDt, timeZone: tz },
      end: { dateTime: endDt, timeZone: tz },
    };
    if (args.body) {
      payload.body = { contentType: "text", content: args.body };
    }
    if (args.location) {
      payload.location = { displayName: args.location };
    }
    if (args.attendees && args.attendees.length > 0) {
      // Normalize each entry to {address, type}, then dedupe by address
      // (case-insensitive) — Graph otherwise happily sends two invites.
      const seen = new Map<string, { address: string; type: string }>();
      for (const a of args.attendees) {
        const entry =
          typeof a === "string"
            ? { address: a, type: "required" as const }
            : { address: a.address, type: (a.type ?? "required") as string };
        seen.set(entry.address.toLowerCase(), entry);
      }
      payload.attendees = Array.from(seen.values()).map((e) => ({
        emailAddress: { address: e.address },
        type: e.type,
      }));
    }

    try {
      const created = (await withAuthRetry(() =>
        graph().api("/me/events").post(payload),
      )) as GraphEventCreateResponse;

      return ok({
        created: true,
        id: created.id,
        start: created.start?.dateTime ?? null,
        end: created.end?.dateTime ?? null,
        timeZone: tz,
        webLink: created.webLink ?? null,
        invitesSent: !!(args.attendees && args.attendees.length > 0),
      });
    } catch (err) {
      const e = explainGraphError(err);
      return fail(e.message, e.hint);
    }
  },
};
