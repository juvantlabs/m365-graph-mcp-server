/**
 * Tool: m365-graph:delete_file
 *
 * Delete a file or folder. Two-phase spec/approval pattern per
 * handbook ADR 0002 § Delete-class operations:
 *
 *   Call 1 (no confirmation_token):
 *     - Tool fetches item metadata for preview.
 *     - Issues a confirmation_token tied to the spec
 *       {item_id, drive_id} and to this tool name.
 *     - Returns the preview + token + expiry hint.
 *
 *   Call 2 (with confirmation_token):
 *     - Tool verifies the token (exists, not expired, tied to this
 *       tool, spec matches).
 *     - DELETEs the item.
 *     - Consumes the token (single-use).
 *
 * The pattern means a deletion can never happen from a single
 * agent/LLM message — the agent must explicitly review the preview
 * and decide to proceed. Defense against accidental destructive
 * actions in long agent loops.
 *
 * Required Graph scope: `Files.ReadWrite` (delegated).
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import {
  consumeConfirmation,
  issueConfirmation,
} from "../auth/confirmation_tokens.js";
import {
  validateOptionalString,
  validateRequiredString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const TOOL_NAME = "m365-graph:delete_file";

const definition: ToolDefinition = {
  name: TOOL_NAME,
  description:
    "Delete a file or folder. Two-phase: first call returns a preview + confirmation_token; second call (with the token + same args) executes the delete. Token is single-use, expires in 5 minutes, and tied to the exact spec — passing a different item_id with someone else's token fails.",
  inputSchema: {
    type: "object",
    properties: {
      item_id: { type: "string", description: "Item to delete." },
      drive_id: { type: "string", description: "Optional drive ID. Defaults to /me/drive." },
      confirmation_token: {
        type: "string",
        description:
          "Omit on the first call (preview mode). Include the token returned by the preview call to execute the delete.",
      },
    },
    required: ["item_id"],
  },
};

interface DeletePreview {
  item: {
    id: string;
    name: string;
    type: "file" | "folder";
    size: number;
    parent_path: string | null;
    webUrl: string;
  };
  confirmation_token: string;
  expires_at: string;
  expires_in_seconds: number;
  instructions: string;
}

function summarizePreview(item: Record<string, unknown>): DeletePreview["item"] {
  const folder = item.folder !== undefined;
  const parent = item.parentReference as Record<string, unknown> | undefined;
  return {
    id: String(item.id ?? ""),
    name: String(item.name ?? ""),
    type: folder ? "folder" : "file",
    size: Number(item.size ?? 0),
    parent_path: (parent?.path as string | undefined) ?? null,
    webUrl: String(item.webUrl ?? ""),
  };
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const itemId = validateRequiredString(args.item_id, "item_id");
  const driveId = validateOptionalString(args.drive_id, "drive_id");
  const confirmationToken = validateOptionalString(args.confirmation_token, "confirmation_token");

  const drivePath = driveId
    ? `/drives/${encodeURIComponent(driveId)}`
    : "/me/drive";
  const apiPath = `${drivePath}/items/${encodeURIComponent(itemId)}`;

  const spec: Record<string, unknown> = { item_id: itemId };
  if (driveId !== undefined) spec.drive_id = driveId;

  // ---------- Phase 1: preview ----------
  if (!confirmationToken) {
    const item = (await graph.api(apiPath).get()) as Record<string, unknown>;
    const summary = summarizePreview(item);
    const issued = issueConfirmation(TOOL_NAME, spec);
    const preview: DeletePreview = {
      item: summary,
      ...issued,
      instructions:
        `Re-call ${TOOL_NAME} with the SAME args (item_id${driveId ? ", drive_id" : ""}) ` +
        `plus this confirmation_token to execute the delete.`,
    };
    return {
      content: [{ type: "text", text: JSON.stringify({ preview }, null, 2) }],
    };
  }

  // ---------- Phase 2: execute ----------
  const verdict = consumeConfirmation(confirmationToken, TOOL_NAME, spec);
  if (!verdict.ok) {
    throw new Error(
      `Refusing to delete: confirmation_token ${verdict.error}. ` +
        `Re-run without confirmation_token to get a fresh preview + token.`,
    );
  }

  await graph.api(apiPath).delete();

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            deleted: { item_id: itemId, drive_id: driveId ?? null },
            note: "DELETE executed against Graph. The item may be recoverable from the user's recycle bin (default 30-day retention).",
          },
          null,
          2,
        ),
      },
    ],
  };
};

export const deleteFileTool: Tool = { category: "write_irreversible", definition, handler };
