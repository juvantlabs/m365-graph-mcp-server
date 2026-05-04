/**
 * Tool: m365-graph:get_event
 *
 * Fetch a single event's full detail (including the full body,
 * attendees with response statuses, location, recurrence rule for
 * series masters). Wraps Graph `/me/events/{id}`.
 *
 * Required Graph scope: `Calendars.Read` (delegated). Read-only.
 *
 * Input:
 *   event_id  (string, required) — id from list_events / search_events
 *
 * Output: full event summary plus `body` (the long-form HTML/text
 * preview, capped) and `recurrence` (the recurrence rule if any).
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import { summarizeEvent } from "./list_events.js";
import { validateRequiredString } from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const BODY_CHAR_CAP = 8_000;

const definition: ToolDefinition = {
  name: "m365-graph:get_event",
  description:
    "Fetch full details for a single event — body, attendees with response statuses, location, recurrence rule. Body is capped at 8000 chars; use the webLink for full content. Read-only.",
  inputSchema: {
    type: "object",
    properties: {
      event_id: {
        type: "string",
        description: "Event ID from list_events or search_events.",
      },
    },
    required: ["event_id"],
  },
};

interface FullEvent extends ReturnType<typeof summarizeEvent> {
  body_content_type: string;
  body: string;
  body_truncated: boolean;
  recurrence: unknown;
}

export function expandEvent(event: Record<string, unknown>): FullEvent {
  const summary = summarizeEvent(event);
  const body = event.body as Record<string, unknown> | undefined;
  const bodyContent = String(body?.content ?? "");
  const truncated = bodyContent.length > BODY_CHAR_CAP;
  return {
    ...summary,
    body_content_type: String(body?.contentType ?? "text"),
    body: truncated ? bodyContent.slice(0, BODY_CHAR_CAP) : bodyContent,
    body_truncated: truncated,
    recurrence: event.recurrence ?? null,
  };
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const eventId = validateRequiredString(args.event_id, "event_id");

  const event = await graph.api(`/me/events/${encodeURIComponent(eventId)}`).get();
  const result = expandEvent(event);

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
};

export const getEventTool: Tool = { category: "read", definition, handler };
