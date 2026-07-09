import { ListCalendarEventsInput } from "../schemas.js";
import { graph, withAuthRetry } from "../graph/client.js";
import { explainGraphError } from "../util/errors.js";
import { fail, ok, type ToolDef } from "./types.js";

/**
 * list_calendar_events: events in a time window.
 *
 * Uses /me/calendarView (not /me/events) because calendarView expands
 * recurring series into individual occurrences — which is almost always what
 * the user actually wants when asking "what's on my calendar next week?".
 *
 * Prefer: outlook.timezone="UTC" normalizes response times to UTC so the LLM
 * doesn't have to reason about the user's tenant timezone.
 */

interface GraphEvent {
  id: string;
  subject?: string | null;
  start?: { dateTime?: string; timeZone?: string };
  end?: { dateTime?: string; timeZone?: string };
  location?: { displayName?: string };
  organizer?: { emailAddress?: { address?: string; name?: string } };
  attendees?: Array<{
    emailAddress?: { address?: string; name?: string };
    status?: { response?: string };
  }>;
  isAllDay?: boolean;
  isCancelled?: boolean;
  onlineMeeting?: { joinUrl?: string } | null;
  bodyPreview?: string;
  webLink?: string;
}

export const listCalendarEventsTool: ToolDef<typeof ListCalendarEventsInput> = {
  name: "list_calendar_events",
  description:
    "List calendar events in a time window. Recurring events are expanded into individual occurrences. " +
    "All returned times are in UTC — convert for display if needed. Use create_event to add a new event. " +
    "`from`/`to` MUST include an explicit offset (e.g. trailing 'Z' or '+02:00') — unlike create_event, " +
    "this tool has no `timeZone` parameter, so a naive (offset-less) datetime would be ambiguous and is " +
    "rejected. Result is capped by `limit` (default 50, max 100); no pagination is followed, so for very " +
    "busy calendars, narrow the window if you hit the cap.",
  schema: ListCalendarEventsInput,
  async handler(args) {
    const limit = args.limit ?? 50;
    // Graph documents startDateTime/endDateTime as UTC — offsets work on
    // most tenants but fail on some. Normalize to the canonical 'Z' form.
    const toUtc = (iso: string) => new Date(iso).toISOString();
    try {
      // Fetch one extra row so we can compute `truncated` without a
      // separate $count call.
      const res = (await withAuthRetry(() => graph()
        .api("/me/calendarView")
        .query({ startDateTime: toUtc(args.from), endDateTime: toUtc(args.to) })
        .header("Prefer", 'outlook.timezone="UTC"')
        .top(limit + 1)
        .orderby("start/dateTime asc")
        .select([
          "id",
          "subject",
          "start",
          "end",
          "location",
          "organizer",
          "attendees",
          "isAllDay",
          "isCancelled",
          "onlineMeeting",
          "bodyPreview",
          "webLink",
        ])
        .get())) as { value?: GraphEvent[] };

      const rawEvents = res.value ?? [];
      const truncated = rawEvents.length > limit;
      const events = (truncated ? rawEvents.slice(0, limit) : rawEvents).map((e) => ({
        id: e.id,
        subject: e.subject ?? "(no subject)",
        start: e.start?.dateTime ?? null,
        end: e.end?.dateTime ?? null,
        allDay: !!e.isAllDay,
        cancelled: !!e.isCancelled,
        location: e.location?.displayName ?? null,
        organizer: e.organizer?.emailAddress?.address ?? null,
        attendees: (e.attendees ?? []).map((a) => ({
          address: a.emailAddress?.address ?? null,
          name: a.emailAddress?.name ?? null,
          response: a.status?.response ?? null,
        })),
        onlineMeetingUrl: e.onlineMeeting?.joinUrl ?? null,
        preview: (e.bodyPreview ?? "").slice(0, 240),
        webLink: e.webLink ?? null,
      }));

      return ok({
        window: { from: args.from, to: args.to, timezone: "UTC" },
        count: events.length,
        // truncated: there are more events in the requested window than we
        // returned. Hint the LLM to narrow the window or raise `limit`.
        truncated,
        events,
      });
    } catch (err) {
      const e = explainGraphError(err);
      return fail(e.message, e.hint);
    }
  },
};
