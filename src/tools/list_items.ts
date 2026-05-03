/**
 * Tool: m365-graph:list_items
 *
 * List the children of a folder (root or specified item) in a drive.
 * Wraps Graph `/drives/{id}/items/{item-id}/children` (or
 * `/me/drive/root/children` when no item_id is provided).
 *
 * Required Graph scope: `Files.Read` (delegated). Read-only.
 *
 * Input:
 *   drive_id  (string, optional) — drive to list; default = /me/drive
 *   item_id   (string, optional) — folder ID; default = drive root
 *   limit     (integer, 1..100)  — max items; default 50
 *
 * Output: JSON list of items with name, type (file|folder), size,
 * lastModified, webUrl.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import {
  validateOptionalInteger,
  validateOptionalString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const definition: ToolDefinition = {
  name: "m365-graph:list_items",
  description:
    "List the immediate children (files + folders) of a folder in a drive. Defaults to the root of the user's primary OneDrive; pass drive_id and/or item_id to navigate elsewhere. Read-only.",
  inputSchema: {
    type: "object",
    properties: {
      drive_id: {
        type: "string",
        description: "Optional drive ID. Defaults to the user's primary OneDrive.",
      },
      item_id: {
        type: "string",
        description:
          "Optional folder item ID. Defaults to the drive root. Pass an item ID returned by a previous list_items / search_files call to navigate into subfolders.",
      },
      limit: {
        type: "integer",
        minimum: 1,
        maximum: 100,
        description: "Maximum number of children to return (default 50).",
      },
    },
    required: [],
  },
};

interface DriveItemSummary {
  id: string;
  name: string;
  type: "file" | "folder";
  size: number;
  child_count: number | null;
  lastModified: string;
  webUrl: string;
}

export function summarizeDriveItem(item: Record<string, unknown>): DriveItemSummary {
  const folder = item.folder as Record<string, unknown> | undefined;
  return {
    id: String(item.id ?? ""),
    name: String(item.name ?? ""),
    type: folder !== undefined ? "folder" : "file",
    size: Number(item.size ?? 0),
    child_count: folder !== undefined ? Number(folder.childCount ?? 0) : null,
    lastModified: String(item.lastModifiedDateTime ?? ""),
    webUrl: String(item.webUrl ?? ""),
  };
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const driveId = validateOptionalString(args.drive_id, "drive_id");
  const itemId = validateOptionalString(args.item_id, "item_id");
  const limit = validateOptionalInteger(args.limit, "limit", {
    min: 1,
    max: 100,
    default: 50,
  });

  const driveRoot = driveId ? `/drives/${encodeURIComponent(driveId)}` : "/me/drive";
  const itemPath = itemId
    ? `${driveRoot}/items/${encodeURIComponent(itemId)}/children`
    : `${driveRoot}/root/children`;

  const response = await graph.api(itemPath).top(limit).get();
  const items = Array.isArray(response?.value) ? response.value : [];
  const summaries = items.map((i: Record<string, unknown>) => summarizeDriveItem(i));

  const result = {
    drive_id: driveId ?? null,
    item_id: itemId ?? null,
    count: summaries.length,
    items: summaries,
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
};

export const listItemsTool: Tool = { definition, handler };
