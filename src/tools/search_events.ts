/**
 * Tool: m365-graph:search_events
 *
 * Search the user's events by subject. Uses Graph
 * `/me/events?$filter=contains(subject, '…')` — subject-only because
 * `$search` is not supported on the Events resource ("The parameter
 * $search is not currently supported on the Events resource."). Body
 * search would require POST /search/query (a separate Search API),
 * deferred to a future tool.
 *
 * Note: $filter on events does NOT expand recurrences — series
 * masters are returned for recurring events. For "all occurrences in
 * a window" use list_events instead.
 *
 * Required Graph scope: `Calendars.Read` (delegated). Read-only.
 *
 * Input:
 *   query  (string, required)  — substring matched against event subject
 *   limit  (integer 1..50)     — default 20
 *
 * Output: same event shape as list_events.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import { summarizeEvent } from "./list_events.js";
import {
  validateOptionalInteger,
  validateRequiredString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const definition: ToolDefinition = {
  name: "m365-graph:search_events",
  description:
    "Search the user's events by subject (substring match). Read-only. Subject-only because Graph does not support $search on the Events resource. For 'events in a date window' use list_events instead — this tool returns recurrence series masters.",
  inputSchema: {
    type: "object",
    properties: {
      query: {
        type: "string",
        description: "Substring matched against the event subject (case-insensitive on the Graph side).",
      },
      limit: {
        type: "integer",
        minimum: 1,
        maximum: 50,
        description: "Maximum number of results to return (default 20).",
      },
    },
    required: ["query"],
  },
};

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const query = validateRequiredString(args.query, "query");
  const limit = validateOptionalInteger(args.limit, "limit", {
    min: 1,
    max: 50,
    default: 20,
  });

  // OData literal-string escape: ' → ''
  const escapedQuery = query.replace(/'/g, "''");

  const response = await graph
    .api("/me/events")
    .filter(`contains(subject, '${escapedQuery}')`)
    .top(limit)
    .get();

  const items = Array.isArray(response?.value) ? response.value : [];
  const events = items.map((e: Record<string, unknown>) => summarizeEvent(e));

  const result = {
    query,
    count: events.length,
    results: events,
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
};

export const searchEventsTool: Tool = { category: "read", definition, handler };
