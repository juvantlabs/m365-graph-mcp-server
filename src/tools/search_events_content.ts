/**
 * Tool: m365-graph:search_events_content
 *
 * Body + subject content search via the Microsoft Search API
 * (POST /search/query). Distinct from `search_events` which uses
 * `/me/events?$filter=contains(subject, '…')` (subject-only, since
 * Graph rejects $search on the Events resource).
 *
 * Trade-offs vs search_events:
 *   - Hits matched against BODY content too, not just subject.
 *   - Returns Search-API-shaped hits: `{hitId, summary, resource}`.
 *     We map back to summarizeEvent's shape using `resource` (the
 *     event payload) for consistency with the rest of the calendar
 *     surface.
 *   - Returns recurrence series masters (same as search_events) — no
 *     occurrence expansion. Use list_events for windowed occurrences.
 *
 * Required scope: `Calendars.Read` (delegated). The Search API
 * inherits the user's permissions; no extra scope needed for events.
 *
 * Input:
 *   query  (string, required)  — natural-language search term
 *   limit  (integer 1..50)     — default 25
 *   from   (integer 0..)       — pagination offset (default 0)
 *
 * Output: `{ query, count, total, results: [<event summary>] }` where
 * `total` is the server's reported total (may exceed `count` if
 * paginated).
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import { summarizeEvent } from "./list_events.js";
import {
  validateOptionalInteger,
  validateRequiredString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const definition: ToolDefinition = {
  name: "m365-graph:search_events_content",
  description:
    "Search the user's events by subject AND body content. Uses the Microsoft Search API (POST /search/query). Distinct from search_events (subject-only via $filter). Returns recurrence series masters; for occurrences in a window use list_events.",
  inputSchema: {
    type: "object",
    properties: {
      query: { type: "string", description: "Search term — matches subject + body." },
      limit: {
        type: "integer",
        minimum: 1,
        maximum: 50,
        description: "Maximum hits to return (default 25).",
      },
      from: {
        type: "integer",
        minimum: 0,
        description: "Pagination offset (default 0).",
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
    default: 25,
  });
  const from = validateOptionalInteger(args.from, "from", {
    min: 0,
    max: 1_000_000,
    default: 0,
  });

  const body = {
    requests: [
      {
        entityTypes: ["event"],
        query: { queryString: query },
        from,
        size: limit,
      },
    ],
  };

  const response = (await graph.api("/search/query").post(body)) as Record<string, unknown>;

  // Search API response shape:
  // { value: [ { hitsContainers: [ { hits: [ { resource: <event>, ... }, ... ], total } ] } ] }
  const value = Array.isArray(response.value) ? response.value : [];
  const firstResponse = (value[0] ?? {}) as Record<string, unknown>;
  const hitsContainers = Array.isArray(firstResponse.hitsContainers)
    ? firstResponse.hitsContainers
    : [];
  const firstContainer = (hitsContainers[0] ?? {}) as Record<string, unknown>;
  const hits = Array.isArray(firstContainer.hits) ? firstContainer.hits : [];
  const total = Number(firstContainer.total ?? hits.length);

  const events = hits
    .map((h: Record<string, unknown>) => h.resource)
    .filter((r): r is Record<string, unknown> => typeof r === "object" && r !== null)
    .map((r) => summarizeEvent(r));

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            query,
            count: events.length,
            total,
            results: events,
          },
          null,
          2,
        ),
      },
    ],
  };
};

export const searchEventsContentTool: Tool = { category: "read", definition, handler };
