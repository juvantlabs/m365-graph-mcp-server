/**
 * Tool: m365-graph:create_event
 *
 * Create a new event on the user's primary calendar (or a specified
 * calendar). Wraps Graph `POST /me/events` (or
 * `POST /me/calendars/{id}/events`).
 *
 * Required Graph scope: `Calendars.ReadWrite` (delegated).
 *
 * Send-invitations behavior: by default Graph DOES send meeting
 * invitations to attendees on event creation. To create an event
 * without sending invites, set `send_invitations: false` (we set
 * `responseStatus.response: "organizer"` and don't add the user as
 * attendee — but Graph still notifies the *other* attendees unless
 * the calendar is configured otherwise).
 *
 * Defaults:
 *   - timezone: "UTC" if not specified (Graph requires explicit TZ)
 *   - body_content_type: "text"
 *   - is_all_day: false
 *   - attendee.type: "required" (when type omitted in input)
 *
 * No spec/approval pattern here — creation is non-destructive (worst
 * case: clutter on the calendar; the agent can call cancel_event or
 * the user can delete via Outlook UI).
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import { summarizeEvent } from "./list_events.js";
import {
  validateOptionalEnum,
  validateOptionalString,
  validateRequiredISODate,
  validateRequiredString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const BODY_CONTENT_TYPES = ["text", "html"] as const;
type BodyContentType = (typeof BODY_CONTENT_TYPES)[number];

const ATTENDEE_TYPES = ["required", "optional", "resource"] as const;
type AttendeeType = (typeof ATTENDEE_TYPES)[number];

interface AttendeeInput {
  email: string;
  name?: string;
  type?: AttendeeType;
}

const definition: ToolDefinition = {
  name: "m365-graph:create_event",
  description:
    "Create a new event on the user's primary calendar (or a specified calendar). Returns the created event with its server-assigned id and webLink. By default Graph sends meeting invitations to attendees.",
  inputSchema: {
    type: "object",
    properties: {
      subject: { type: "string", description: "Event subject (title)." },
      start: {
        type: "string",
        description: "Start, ISO 8601 (e.g. '2026-05-04T10:00:00').",
      },
      end: {
        type: "string",
        description: "End, ISO 8601. Must be after start.",
      },
      timezone: {
        type: "string",
        description: "IANA timezone for start + end (e.g. 'Europe/Rome'). Default: 'UTC'.",
      },
      body: {
        type: "string",
        description: "Optional body content for the event description.",
      },
      body_content_type: {
        type: "string",
        enum: ["text", "html"],
        description: "Body content type. Default: 'text'.",
      },
      location: {
        type: "string",
        description: "Optional location (free-text or room name).",
      },
      attendees: {
        type: "array",
        description: "Optional attendees. Each item: {email, name?, type?}. type ∈ {required, optional, resource}.",
        items: {
          type: "object",
          properties: {
            email: { type: "string" },
            name: { type: "string" },
            type: { type: "string", enum: ["required", "optional", "resource"] },
          },
          required: ["email"],
        },
      },
      is_all_day: {
        type: "boolean",
        description: "Mark as all-day event. Default: false.",
      },
      calendar_id: {
        type: "string",
        description: "Optional calendar id from list_calendars. Defaults to the user's primary calendar.",
      },
    },
    required: ["subject", "start", "end"],
  },
};

function validateAttendees(value: unknown): AttendeeInput[] {
  if (value === undefined || value === null) return [];
  if (!Array.isArray(value)) {
    throw new Error("'attendees' must be an array");
  }
  return value.map((raw, i) => {
    if (typeof raw !== "object" || raw === null) {
      throw new Error(`'attendees[${i}]' must be an object`);
    }
    const a = raw as Record<string, unknown>;
    const email = validateRequiredString(a.email, `attendees[${i}].email`);
    const name = validateOptionalString(a.name, `attendees[${i}].name`);
    const type = validateOptionalEnum<AttendeeType>(
      a.type,
      `attendees[${i}].type`,
      ATTENDEE_TYPES,
      "required",
    );
    return { email, name, type };
  });
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const subject = validateRequiredString(args.subject, "subject");
  const start = validateRequiredISODate(args.start, "start");
  const end = validateRequiredISODate(args.end, "end");
  const timezone = validateOptionalString(args.timezone, "timezone") ?? "UTC";
  const body = validateOptionalString(args.body, "body");
  const bodyContentType = validateOptionalEnum<BodyContentType>(
    args.body_content_type,
    "body_content_type",
    BODY_CONTENT_TYPES,
    "text",
  );
  const location = validateOptionalString(args.location, "location");
  const attendees = validateAttendees(args.attendees);
  const isAllDay = args.is_all_day === true;
  const calendarId = validateOptionalString(args.calendar_id, "calendar_id");

  const eventBody: Record<string, unknown> = {
    subject,
    start: { dateTime: start, timeZone: timezone },
    end: { dateTime: end, timeZone: timezone },
    isAllDay,
  };

  if (body !== undefined) {
    eventBody.body = { contentType: bodyContentType, content: body };
  }
  if (location !== undefined) {
    eventBody.location = { displayName: location };
  }
  if (attendees.length > 0) {
    eventBody.attendees = attendees.map((a) => ({
      emailAddress: a.name ? { address: a.email, name: a.name } : { address: a.email },
      type: a.type ?? "required",
    }));
  }

  const apiPath = calendarId
    ? `/me/calendars/${encodeURIComponent(calendarId)}/events`
    : "/me/events";

  const created = await graph.api(apiPath).post(eventBody);
  const summary = summarizeEvent(created);

  return {
    content: [{ type: "text", text: JSON.stringify({ created: summary }, null, 2) }],
  };
};

export const createEventTool: Tool = { category: "write_idempotent", definition, handler };
