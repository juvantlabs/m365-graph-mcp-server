/**
 * Tool registry — single source of truth for which tools the server
 * exposes. Both the tools/list handler and the tools/call dispatcher
 * read from this module.
 *
 * Each new tool: add an import here + push into ALL_TOOLS. The
 * dispatcher in src/index.ts builds a map from tool name to handler
 * once at startup; no per-call registration overhead.
 */

import type { Tool } from "../types/tool.js";
import { downloadFileTool } from "./download_file.js";
import { getEventTool } from "./get_event.js";
import { listCalendarsTool } from "./list_calendars.js";
import { listDrivesTool } from "./list_drives.js";
import { listEventsTool } from "./list_events.js";
import { listItemsTool } from "./list_items.js";
import { searchEventsTool } from "./search_events.js";
import { searchFilesTool } from "./search_files.js";

export const ALL_TOOLS: ReadonlyArray<Tool> = [
  // Files
  listDrivesTool,
  listItemsTool,
  searchFilesTool,
  downloadFileTool,
  // Calendars
  listCalendarsTool,
  listEventsTool,
  searchEventsTool,
  getEventTool,
];

export function buildHandlerMap(tools: ReadonlyArray<Tool>): Map<string, Tool["handler"]> {
  const m = new Map<string, Tool["handler"]>();
  for (const t of tools) {
    m.set(t.definition.name, t.handler);
  }
  return m;
}
