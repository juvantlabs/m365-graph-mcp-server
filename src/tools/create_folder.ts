/**
 * Tool: m365-graph:create_folder
 *
 * Create a folder in a drive (OneDrive primary by default; SharePoint
 * document library via `drive_id`). Wraps Graph's synchronous
 * children-POST:
 *
 *   POST /me/drive/root/children                       (root of primary OneDrive)
 *   POST /me/drive/items/{parent-id}/children          (nested)
 *   POST /drives/{drive-id}/root/children              (non-default drive)
 *   POST /drives/{drive-id}/items/{parent-id}/children (nested)
 *
 * Body:
 *   {
 *     "name": "<name>",
 *     "folder": {},
 *     "@microsoft.graph.conflictBehavior": "fail" | "replace" | "rename"
 *   }
 *
 * Conflict behavior:
 *   - `fail`    (default) — refuse if an item with that name exists
 *   - `replace` — replace the existing item (Graph accepts this; note
 *     that "replacing" an existing folder is a rare intent, so default
 *     stays `fail`)
 *   - `rename`  — let Graph auto-suffix (`Foo` → `Foo 1`)
 *
 * Category: `write_idempotent` — a repeated call with
 * `conflict_behavior: fail` returns an error rather than creating a
 * duplicate; the created folder is reversible via `delete_file`. No
 * two-phase confirmation needed (creation, not destruction).
 *
 * Required Graph scope: `Files.ReadWrite` (delegated) for OneDrive;
 * `Sites.ReadWrite.All` (delegated) for SharePoint document libraries.
 * Both are already requested by src/auth/msal.ts as of v0.4.0.
 *
 * Path/name notes:
 *   - We only address the parent by ID (`parent_item_id`), consistent
 *     with the rest of the file-tool surface. Path-string addressing
 *     is intentionally out of scope.
 *   - Names with leading underscore (`_SDI-ufficiali`) or spaces
 *     (`New Folder`) are legal in OneDrive/SharePoint doc libraries;
 *     we pass through unmodified. Reserved characters (`\ / : * ? " <
 *     > |`) and reserved DOS-style names (CON, PRN, etc.) are rejected
 *     by Graph with a 400 that surfaces back to the agent as a clear
 *     error — no client-side sanitization that would silently mutate
 *     the requested name.
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import {
  validateOptionalEnum,
  validateOptionalString,
  validateRequiredString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const CONFLICT_BEHAVIORS = ["fail", "replace", "rename"] as const;
type ConflictBehavior = (typeof CONFLICT_BEHAVIORS)[number];

const definition: ToolDefinition = {
  name: "m365-graph:create_folder",
  description:
    "Create a folder in a drive (OneDrive primary by default; SharePoint document library via drive_id). Synchronous POST to /children. Default conflict behavior is 'fail' — set to 'rename' to auto-suffix on collision. Returns the created folder's item ID so it can be passed as parent_item_id to a subsequent upload_file / create_folder call.",
  inputSchema: {
    type: "object",
    properties: {
      name: {
        type: "string",
        description:
          "Folder name to create. Passed through unmodified — Graph rejects reserved characters (\\ / : * ? \" < > |) with a 400 the agent can act on.",
      },
      parent_item_id: {
        type: "string",
        description:
          "Optional parent folder ID to create INTO. Defaults to the drive root. Get from list_items.",
      },
      drive_id: {
        type: "string",
        description:
          "Optional drive ID. Defaults to the user's primary OneDrive. Use for SharePoint document libraries.",
      },
      conflict_behavior: {
        type: "string",
        enum: ["fail", "replace", "rename"],
        description:
          "What to do if an item with this name already exists in the parent. Default: 'fail'.",
      },
    },
    required: ["name"],
  },
};

interface CreatedFolderSummary {
  id: string;
  name: string;
  webUrl: string;
  parent_id: string | null;
  child_count: number;
}

export function summarizeCreatedFolder(item: Record<string, unknown>): CreatedFolderSummary {
  const parent = item.parentReference as Record<string, unknown> | undefined;
  const folder = item.folder as Record<string, unknown> | undefined;
  return {
    id: String(item.id ?? ""),
    name: String(item.name ?? ""),
    webUrl: String(item.webUrl ?? ""),
    parent_id: (parent?.id as string | undefined) ?? null,
    child_count: Number(folder?.childCount ?? 0),
  };
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const name = validateRequiredString(args.name, "name");
  const parentItemId = validateOptionalString(args.parent_item_id, "parent_item_id");
  const driveId = validateOptionalString(args.drive_id, "drive_id");
  const conflictBehavior = validateOptionalEnum<ConflictBehavior>(
    args.conflict_behavior,
    "conflict_behavior",
    CONFLICT_BEHAVIORS,
    "fail",
  );

  const drivePath = driveId ? `/drives/${encodeURIComponent(driveId)}` : "/me/drive";
  const parentPath = parentItemId
    ? `${drivePath}/items/${encodeURIComponent(parentItemId)}/children`
    : `${drivePath}/root/children`;

  const body: Record<string, unknown> = {
    name,
    folder: {},
    "@microsoft.graph.conflictBehavior": conflictBehavior,
  };

  console.error(
    `[m365-graph-mcp-server] create_folder: ${drivePath}/${parentItemId ?? "<root>"}/${name} (conflict=${conflictBehavior})`,
  );

  const created = (await graph.api(parentPath).post(body)) as Record<string, unknown>;

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            drive_id: driveId ?? null,
            parent_item_id: parentItemId ?? null,
            created: summarizeCreatedFolder(created),
          },
          null,
          2,
        ),
      },
    ],
  };
};

export const createFolderTool: Tool = { category: "write_idempotent", definition, handler };
