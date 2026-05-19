import { UpdateEventInput } from "../schemas.js";
import { graph, withAuthRetry } from "../graph/client.js";
import { explainGraphError, ToolError } from "../util/errors.js";
import { fail, ok, type ToolDef } from "./types.js";
import { normalizeDateTime, normalizeAttendees } from "./createEvent.js";

/**
 * update_event: partial update of an existing calendar event.
 *
 * PATCH /me/events/{id} with only the fields the caller provided. Graph
 * preserves any field absent from the body, so this is a sparse merge — NOT
 * a full replace.
 *
 * Exception: attendees. If `attendees` is present in the input, Graph
 * REPLACES the entire attendee list (no merge semantics at the API level).
 * The schema description spells this out so the LLM doesn't accidentally
 * drop attendees by passing only the new ones.
 *
 * Time handling: we reuse normalizeDateTime from createEvent for the same
 * "naive vs offset vs timeZone" matrix. A start update without an end
 * update (or vice versa) is allowed — Graph will validate that the
 * resulting window is still well-ordered and 400 if not.
 *
 * Cancellation notices: Graph automatically emails attendees when the event
 * subject, time, or location changes on a sent meeting. There is no
 * "silent update" mode at the API. Document for the caller in the tool
 * description.
 */

interface GraphEventUpdateResponse {
  id?: string;
  webLink?: string;
  subject?: string | null;
  start?: { dateTime?: string; timeZone?: string };
  end?: { dateTime?: string; timeZone?: string };
}

export const updateEventTool: ToolDef<typeof UpdateEventInput> = {
  name: "update_event",
  description:
    "Update an existing calendar event (PATCH /me/events). Only fields you pass are changed; everything " +
    "else is preserved. EXCEPTION: if 'attendees' is provided, the entire attendee list is replaced — " +
    "fetch the event first if you only want to add or remove individuals. Graph will email update " +
    "notifications to affected attendees automatically; there is no 'silent' mode.",
  schema: UpdateEventInput,
  async handler(args) {
    const tz = args.timeZone ?? "UTC";

    const payload: Record<string, unknown> = {};

    if (args.subject !== undefined) payload.subject = args.subject;
    if (args.body !== undefined) {
      payload.body = { contentType: "text", content: args.body };
    }
    if (args.location !== undefined) {
      payload.location = { displayName: args.location };
    }

    try {
      if (args.start !== undefined) {
        payload.start = { dateTime: normalizeDateTime(args.start, tz, "start"), timeZone: tz };
      }
      if (args.end !== undefined) {
        payload.end = { dateTime: normalizeDateTime(args.end, tz, "end"), timeZone: tz };
      }
    } catch (err) {
      if (err instanceof ToolError) return fail(err.message, err.hint);
      throw err;
    }

    // If BOTH start and end are being updated, reject end <= start before
    // hitting Graph. If only one is updated, we cannot check without
    // fetching the event first — defer to Graph's validation.
    if (args.start !== undefined && args.end !== undefined) {
      if (new Date(args.end).getTime() <= new Date(args.start).getTime()) {
        return fail(
          `Event end (${args.end}) is not after start (${args.start}).`,
          "Swap the values or extend the end time.",
        );
      }
    }

    if (args.attendees !== undefined) {
      payload.attendees = normalizeAttendees(args.attendees);
    }

    if (Object.keys(payload).length === 0) {
      return fail(
        "Nothing to update — at least one field besides eventId must be provided.",
        "Pass any of subject, start, end, body, location, or attendees.",
      );
    }

    const eid = encodeURIComponent(args.eventId);
    try {
      const updated = (await withAuthRetry(() =>
        graph().api(`/me/events/${eid}`).patch(payload),
      )) as GraphEventUpdateResponse;

      return ok({
        updated: true,
        id: updated.id ?? args.eventId,
        subject: updated.subject ?? null,
        start: updated.start?.dateTime ?? null,
        end: updated.end?.dateTime ?? null,
        timeZone: tz,
        webLink: updated.webLink ?? null,
        notificationsSent: args.attendees !== undefined,
      });
    } catch (err) {
      const e = explainGraphError(err);
      return fail(e.message, e.hint);
    }
  },
};
