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
// Files — read
import { downloadFileTool } from "./download_file.js";
import { listDrivesTool } from "./list_drives.js";
import { listItemsTool } from "./list_items.js";
import { searchFilesTool } from "./search_files.js";
// Files — write
import { copyFileTool } from "./copy_file.js";
import { deleteFileTool } from "./delete_file.js";
import { moveFileTool } from "./move_file.js";
import { uploadFileTool } from "./upload_file.js";
// Calendars — read
import { getEventTool } from "./get_event.js";
import { listCalendarsTool } from "./list_calendars.js";
import { listEventsTool } from "./list_events.js";
import { searchEventsTool } from "./search_events.js";
import { searchEventsContentTool } from "./search_events_content.js";
// Calendars — write
import { cancelEventTool } from "./cancel_event.js";
import { createEventTool } from "./create_event.js";
import { declineEventTool } from "./decline_event.js";
import { updateEventTool } from "./update_event.js";

export const ALL_TOOLS: ReadonlyArray<Tool> = [
  // Files — read
  listDrivesTool,
  listItemsTool,
  searchFilesTool,
  downloadFileTool,
  // Files — write
  uploadFileTool,
  copyFileTool,
  moveFileTool,
  deleteFileTool,
  // Calendars — read
  listCalendarsTool,
  listEventsTool,
  searchEventsTool,
  searchEventsContentTool,
  getEventTool,
  // Calendars — write
  createEventTool,
  updateEventTool,
  cancelEventTool,
  declineEventTool,
];

export function buildHandlerMap(tools: ReadonlyArray<Tool>): Map<string, Tool["handler"]> {
  const m = new Map<string, Tool["handler"]>();
  for (const t of tools) {
    m.set(t.definition.name, t.handler);
  }
  return m;
}
