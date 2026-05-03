/**
 * Tool: m365-graph:list_calendars
 *
 * List the user's calendars (primary calendar + any group / shared
 * calendars they have access to). Wraps Graph `/me/calendars`.
 *
 * Required Graph scope: `Calendars.Read` (delegated). Read-only.
 *
 * Output: JSON list of calendars with id, name, color, owner,
 * canEdit, canShare, isDefaultCalendar.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import { validateOptionalInteger } from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const definition: ToolDefinition = {
  name: "m365-graph:list_calendars",
  description:
    "List the user's calendars: their primary calendar plus any group or shared calendars they have access to. Read-only. Use the returned id with list_events / search_events / get_event to scope to a specific calendar.",
  inputSchema: {
    type: "object",
    properties: {
      limit: {
        type: "integer",
        minimum: 1,
        maximum: 100,
        description: "Maximum number of calendars to return (default 50).",
      },
    },
    required: [],
  },
};

interface CalendarSummary {
  id: string;
  name: string;
  color: string;
  owner_name: string | null;
  owner_email: string | null;
  is_default: boolean;
  can_edit: boolean;
  can_share: boolean;
}

export function summarizeCalendar(cal: Record<string, unknown>): CalendarSummary {
  const owner = cal.owner as Record<string, unknown> | undefined;
  return {
    id: String(cal.id ?? ""),
    name: String(cal.name ?? ""),
    color: String(cal.color ?? "auto"),
    owner_name: (owner?.name as string | undefined) ?? null,
    owner_email: (owner?.address as string | undefined) ?? null,
    is_default: Boolean(cal.isDefaultCalendar),
    can_edit: Boolean(cal.canEdit),
    can_share: Boolean(cal.canShare),
  };
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const limit = validateOptionalInteger(args.limit, "limit", {
    min: 1,
    max: 100,
    default: 50,
  });

  const response = await graph.api("/me/calendars").top(limit).get();
  const items = Array.isArray(response?.value) ? response.value : [];
  const calendars = items.map((c: Record<string, unknown>) => summarizeCalendar(c));

  const result = {
    count: calendars.length,
    calendars,
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
};

export const listCalendarsTool: Tool = { definition, handler };
