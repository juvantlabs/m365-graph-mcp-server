/**
 * Tool: m365-graph:list_events
 *
 * List events in a date window. Uses Graph `/me/calendarView`
 * (or `/me/calendars/{id}/calendarView` for a specific calendar)
 * which **expands recurrences** — each occurrence appears as its own
 * event, which is what humans + agents almost always want when asked
 * "what's on my calendar tomorrow?". Compare to `/me/events` which
 * returns recurrence series masters.
 *
 * Required Graph scope: `Calendars.Read` (delegated). Read-only.
 *
 * Input:
 *   start         (ISO date/datetime, required)
 *   end           (ISO date/datetime, required)
 *   calendar_id   (string, optional)
 *   limit         (integer 1..200, default 100)
 *
 * Output: JSON list of events with id, subject, start, end, organizer,
 * attendees, location, body_preview, is_all_day, web_url.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import {
  validateOptionalInteger,
  validateOptionalString,
  validateRequiredISODate,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const definition: ToolDefinition = {
  name: "m365-graph:list_events",
  description:
    "List events in a date window from the user's primary calendar (or a specified calendar_id). Recurrences are expanded — each occurrence is one event in the response. Read-only.",
  inputSchema: {
    type: "object",
    properties: {
      start: {
        type: "string",
        description:
          "Window start, ISO 8601 (e.g. '2026-05-04' or '2026-05-04T00:00:00Z').",
      },
      end: {
        type: "string",
        description:
          "Window end, ISO 8601. Exclusive upper bound is recommended (e.g. midnight of the day after the last day you want).",
      },
      calendar_id: {
        type: "string",
        description: "Optional calendar id from list_calendars. Defaults to /me's primary calendar.",
      },
      limit: {
        type: "integer",
        minimum: 1,
        maximum: 200,
        description: "Maximum events to return (default 100).",
      },
    },
    required: ["start", "end"],
  },
};

interface AttendeeSummary {
  name: string;
  email: string;
  type: string;
  response: string;
}

interface EventSummary {
  id: string;
  subject: string;
  body_preview: string;
  start: { datetime: string; timezone: string };
  end: { datetime: string; timezone: string };
  is_all_day: boolean;
  location: string | null;
  organizer_name: string | null;
  organizer_email: string | null;
  attendees: AttendeeSummary[];
  webLink: string;
}

function summarizeAttendee(a: Record<string, unknown>): AttendeeSummary {
  const email = a.emailAddress as Record<string, unknown> | undefined;
  const status = a.status as Record<string, unknown> | undefined;
  return {
    name: (email?.name as string | undefined) ?? "",
    email: (email?.address as string | undefined) ?? "",
    type: String(a.type ?? "required"),
    response: String(status?.response ?? "none"),
  };
}

export function summarizeEvent(event: Record<string, unknown>): EventSummary {
  const start = event.start as Record<string, unknown> | undefined;
  const end = event.end as Record<string, unknown> | undefined;
  const location = event.location as Record<string, unknown> | undefined;
  const organizer = event.organizer as Record<string, unknown> | undefined;
  const orgEmail = organizer?.emailAddress as Record<string, unknown> | undefined;
  const attendees = Array.isArray(event.attendees) ? event.attendees : [];
  return {
    id: String(event.id ?? ""),
    subject: String(event.subject ?? ""),
    body_preview: String(event.bodyPreview ?? ""),
    start: {
      datetime: String(start?.dateTime ?? ""),
      timezone: String(start?.timeZone ?? ""),
    },
    end: {
      datetime: String(end?.dateTime ?? ""),
      timezone: String(end?.timeZone ?? ""),
    },
    is_all_day: Boolean(event.isAllDay),
    location: (location?.displayName as string | undefined) || null,
    organizer_name: (orgEmail?.name as string | undefined) ?? null,
    organizer_email: (orgEmail?.address as string | undefined) ?? null,
    attendees: attendees.map((a) => summarizeAttendee(a as Record<string, unknown>)),
    webLink: String(event.webLink ?? ""),
  };
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const start = validateRequiredISODate(args.start, "start");
  const end = validateRequiredISODate(args.end, "end");
  const calendarId = validateOptionalString(args.calendar_id, "calendar_id");
  const limit = validateOptionalInteger(args.limit, "limit", {
    min: 1,
    max: 200,
    default: 100,
  });

  const apiBase = calendarId
    ? `/me/calendars/${encodeURIComponent(calendarId)}/calendarView`
    : "/me/calendarView";

  const response = await graph
    .api(apiBase)
    .query({ startDateTime: start, endDateTime: end })
    .top(limit)
    .orderby("start/dateTime")
    .get();

  const items = Array.isArray(response?.value) ? response.value : [];
  const events = items.map((e: Record<string, unknown>) => summarizeEvent(e));

  const result = {
    calendar_id: calendarId ?? null,
    window: { start, end },
    count: events.length,
    events,
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
};

export const listEventsTool: Tool = { category: "read", definition, handler };
