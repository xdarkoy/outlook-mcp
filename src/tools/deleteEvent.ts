import { DeleteEventInput } from "../schemas.js";
import { graph, withAuthRetry } from "../graph/client.js";
import { explainGraphError } from "../util/errors.js";
import { fail, ok, type ToolDef } from "./types.js";

/**
 * delete_event: remove a calendar event.
 *
 * DELETE /me/events/{id}. Graph behavior:
 *   - Event with no attendees: gone.
 *   - Meeting with attendees on a sent invite: Graph emails a CANCELLATION
 *     notice to every attendee, then removes the event. There is no
 *     "silent delete" mode at the API. The destructive side effect is
 *     visible (and irreversible from this tool's side).
 *
 * The description below makes the cancellation behavior loud so a calling
 * LLM cannot delete an organized meeting without surfacing the
 * implication. The server itself does not gate; that is the host's job.
 *
 * Returns 204 on success — no response body. We don't echo the event
 * payload back because there's nothing left to echo, and we don't have
 * the subject for confirmation unless we GET the event first. Caller
 * should fetch the event with list_calendar_events before deleting if
 * they want a confirmation summary.
 */

export const deleteEventTool: ToolDef<typeof DeleteEventInput> = {
  name: "delete_event",
  description:
    "Delete a calendar event. DESTRUCTIVE: if the event has attendees, Microsoft Graph automatically " +
    "sends a meeting cancellation to each one — there is no silent-delete mode. Solo events are simply " +
    "removed. The operation is not reversible from this server. Make the implication explicit to the " +
    "user before invoking. Returns the deleted event id on success.",
  schema: DeleteEventInput,
  async handler(args) {
    const eid = encodeURIComponent(args.eventId);
    try {
      await withAuthRetry(() => graph().api(`/me/events/${eid}`).delete());
      return ok({
        deleted: true,
        id: args.eventId,
      });
    } catch (err) {
      const e = explainGraphError(err);
      return fail(e.message, e.hint);
    }
  },
};
