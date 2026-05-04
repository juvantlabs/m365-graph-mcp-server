/**
 * Tool: m365-graph:update_event
 *
 * Update an existing event. Wraps Graph `PATCH /me/events/{id}`.
 *
 * Required Graph scope: `Calendars.ReadWrite` (delegated).
 *
 * Important: when updating attendees, Graph REPLACES the entire list
 * — it does not merge. To add a single attendee, fetch the event with
 * get_event, append to the existing list, and pass the full updated
 * list. The tool documents this in the `attendees` field description
 * so the agent doesn't accidentally truncate the invite list.
 *
 * Send-update behavior: Graph automatically notifies attendees of
 * changes that affect the meeting (time, location, attendee list).
 * That's the desired UX for a meeting invitation flow; for cosmetic
 * updates the user's email client may not show "updated" indicators.
 *
 * Returns the updated event in summarizeEvent shape.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import { summarizeEvent } from "./list_events.js";
import {
  validateOptionalEnum,
  validateOptionalString,
  validateRequiredString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const BODY_CONTENT_TYPES = ["text", "html"] as const;
type BodyContentType = (typeof BODY_CONTENT_TYPES)[number];

const ATTENDEE_TYPES = ["required", "optional", "resource"] as const;
type AttendeeType = (typeof ATTENDEE_TYPES)[number];

const definition: ToolDefinition = {
  name: "m365-graph:update_event",
  description:
    "Update an existing event (subject, time, body, location, attendees). All fields except event_id are optional — only the provided fields are PATCHed. WARNING: when updating attendees, Graph REPLACES the entire list. Pass the full intended list, not a delta. Returns the updated event.",
  inputSchema: {
    type: "object",
    properties: {
      event_id: { type: "string", description: "Event id from list_events / search_events." },
      subject: { type: "string" },
      start: { type: "string", description: "ISO 8601 datetime; if updating time, also pass `end` and optionally `timezone`." },
      end: { type: "string" },
      timezone: { type: "string", description: "IANA timezone (e.g. 'Europe/Rome'). Required if `start` or `end` is updated." },
      body: { type: "string" },
      body_content_type: { type: "string", enum: ["text", "html"], description: "Default: 'text'." },
      location: { type: "string" },
      attendees: {
        type: "array",
        description: "REPLACES the existing attendee list. Each item: {email, name?, type?}.",
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
      is_all_day: { type: "boolean" },
    },
    required: ["event_id"],
  },
};

function buildPatchBody(args: Record<string, unknown>): Record<string, unknown> {
  const body: Record<string, unknown> = {};

  const subject = validateOptionalString(args.subject, "subject");
  if (subject !== undefined) body.subject = subject;

  const start = validateOptionalString(args.start, "start");
  const end = validateOptionalString(args.end, "end");
  const timezone = validateOptionalString(args.timezone, "timezone");
  if (start !== undefined || end !== undefined) {
    if (timezone === undefined) {
      throw new Error("'timezone' is required when 'start' or 'end' is updated");
    }
    if (start !== undefined) body.start = { dateTime: start, timeZone: timezone };
    if (end !== undefined) body.end = { dateTime: end, timeZone: timezone };
  }

  const bodyText = validateOptionalString(args.body, "body");
  if (bodyText !== undefined) {
    const contentType = validateOptionalEnum<BodyContentType>(
      args.body_content_type,
      "body_content_type",
      BODY_CONTENT_TYPES,
      "text",
    );
    body.body = { contentType, content: bodyText };
  }

  const location = validateOptionalString(args.location, "location");
  if (location !== undefined) body.location = { displayName: location };

  if (args.attendees !== undefined) {
    if (!Array.isArray(args.attendees)) throw new Error("'attendees' must be an array");
    body.attendees = args.attendees.map((raw, i) => {
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
      return {
        emailAddress: name ? { address: email, name } : { address: email },
        type,
      };
    });
  }

  if (args.is_all_day !== undefined) {
    if (typeof args.is_all_day !== "boolean") {
      throw new Error("'is_all_day' must be boolean");
    }
    body.isAllDay = args.is_all_day;
  }

  return body;
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const eventId = validateRequiredString(args.event_id, "event_id");
  const patchBody = buildPatchBody(args);
  if (Object.keys(patchBody).length === 0) {
    throw new Error("update_event requires at least one field to update besides event_id");
  }

  const updated = await graph
    .api(`/me/events/${encodeURIComponent(eventId)}`)
    .patch(patchBody);

  const summary = summarizeEvent(updated);

  return {
    content: [{ type: "text", text: JSON.stringify({ updated: summary }, null, 2) }],
  };
};

export { buildPatchBody };
export const updateEventTool: Tool = { category: "write_idempotent", definition, handler };
