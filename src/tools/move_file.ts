/**
 * Tool: m365-graph:move_file
 *
 * Move a file or folder to a different parent folder. Wraps Graph's
 * `PATCH /me/drive/items/{id}` with a `parentReference` body — atomic,
 * synchronous within the same drive (no async polling required, unlike
 * copy_file).
 *
 * Cross-drive moves are NOT supported via this PATCH (Microsoft
 * recommends "copy + delete" for those). The tool refuses
 * `target_drive_id` for now to make this restriction explicit; agents
 * who need cross-drive should call copy_file then delete_file.
 *
 * Required Graph scope: `Files.ReadWrite` (delegated).
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import {
  validateOptionalString,
  validateRequiredString,
} from "../types/validators.js";
import { summarizeCopiedItem } from "./copy_file.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const definition: ToolDefinition = {
  name: "m365-graph:move_file",
  description:
    "Move a file or folder to a different parent folder within the same drive. Synchronous (atomic PATCH). Cross-drive moves are not supported here — use copy_file + delete_file for those.",
  inputSchema: {
    type: "object",
    properties: {
      item_id: { type: "string", description: "Source item ID." },
      target_parent_id: { type: "string", description: "Destination folder ID (in the same drive)." },
      drive_id: { type: "string", description: "Optional drive ID. Defaults to /me/drive." },
      new_name: { type: "string", description: "Optional rename during the move." },
    },
    required: ["item_id", "target_parent_id"],
  },
};

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const itemId = validateRequiredString(args.item_id, "item_id");
  const targetParentId = validateRequiredString(args.target_parent_id, "target_parent_id");
  const driveId = validateOptionalString(args.drive_id, "drive_id");
  const newName = validateOptionalString(args.new_name, "new_name");

  const apiPath = driveId
    ? `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`
    : `/me/drive/items/${encodeURIComponent(itemId)}`;

  const body: Record<string, unknown> = {
    parentReference: { id: targetParentId },
  };
  if (newName !== undefined) body.name = newName;

  const moved = (await graph.api(apiPath).patch(body)) as Record<string, unknown>;
  const summary = summarizeCopiedItem(moved);

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            moved: summary,
            source: { item_id: itemId, drive_id: driveId ?? null },
          },
          null,
          2,
        ),
      },
    ],
  };
};

export const moveFileTool: Tool = { category: "write_idempotent", definition, handler };
