/**
 * Tool: m365-graph:search_files
 *
 * Search for files by name/content within a drive (the user's primary
 * OneDrive by default, or a specified `drive_id`). Wraps the Graph
 * `/drives/{id}/root/search(q='…')` OData function.
 *
 * Required Graph scope: `Files.Read` (delegated). Read-only.
 *
 * Input:
 *   query     (string, required) — search term
 *   drive_id  (string, optional) — drive to search; default = /me/drive
 *   limit     (integer, 1..50)   — max items to return; default 20
 *
 * Output: JSON array of matching items with name, path, size, lastModified, webUrl.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import {
  validateOptionalInteger,
  validateOptionalString,
  validateRequiredString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const definition: ToolDefinition = {
  name: "m365-graph:search_files",
  description:
    "Search for files within a drive by name and content. Defaults to the user's primary OneDrive; pass drive_id to search a specific drive (e.g. a SharePoint document library). Returns up to `limit` matching items. Read-only.",
  inputSchema: {
    type: "object",
    properties: {
      query: {
        type: "string",
        description: "Search term — matches file name and content.",
      },
      drive_id: {
        type: "string",
        description:
          "Optional drive ID to search. Defaults to the user's primary OneDrive (/me/drive).",
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

interface SearchHit {
  id: string;
  name: string;
  path: string;
  size: number;
  is_folder: boolean;
  lastModified: string;
  webUrl: string;
}

export function summarizeSearchHit(item: Record<string, unknown>): SearchHit {
  const parentRef = item.parentReference as Record<string, unknown> | undefined;
  const parentPath = (parentRef?.path as string | undefined) ?? "";
  const name = String(item.name ?? "");
  const path = parentPath ? `${parentPath}/${name}` : name;
  return {
    id: String(item.id ?? ""),
    name,
    path,
    size: Number(item.size ?? 0),
    is_folder: item.folder !== undefined,
    lastModified: String(item.lastModifiedDateTime ?? ""),
    webUrl: String(item.webUrl ?? ""),
  };
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const query = validateRequiredString(args.query, "query");
  const driveId = validateOptionalString(args.drive_id, "drive_id");
  const limit = validateOptionalInteger(args.limit, "limit", {
    min: 1,
    max: 50,
    default: 20,
  });

  const driveRoot = driveId ? `/drives/${encodeURIComponent(driveId)}` : "/me/drive";
  const escapedQuery = query.replace(/'/g, "''");
  const apiPath = `${driveRoot}/root/search(q='${escapedQuery}')`;

  const response = await graph.api(apiPath).top(limit).get();
  const items = Array.isArray(response?.value) ? response.value : [];
  const hits = items.map((i: Record<string, unknown>) => summarizeSearchHit(i));

  const result = {
    query,
    drive_id: driveId ?? null,
    count: hits.length,
    results: hits,
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
};

export const searchFilesTool: Tool = { definition, handler };
