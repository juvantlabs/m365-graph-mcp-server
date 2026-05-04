/**
 * Tool: m365-graph:copy_file
 *
 * Copy a file or folder to a different location. Wraps Graph's async
 * copy operation:
 *
 *   POST /me/drive/items/{id}/copy   → 202 + Location header
 *   GET  <monitor URL>               → poll until completed | failed
 *   GET  <resourceLocation>          → fetch the new item's metadata
 *
 * Polling: exponential backoff (1s → 2s → 4s → … capped at 30s)
 * until completion or `wait_max_seconds` deadline. Per handbook
 * anti-pattern S8, the tool never returns "initiated successfully" —
 * either the copy completes (we return the new item) or we surface
 * the failure / timeout to the agent.
 *
 * Required Graph scope: `Files.ReadWrite` (delegated).
 */

import {
  type Client,
  ResponseType,
} from "@microsoft/microsoft-graph-client";

import {
  validateOptionalInteger,
  validateOptionalString,
  validateRequiredString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const POLL_INITIAL_MS = 1_000;
const POLL_MAX_MS = 30_000;

const definition: ToolDefinition = {
  name: "m365-graph:copy_file",
  description:
    "Copy a file or folder to a different location in a drive (or across drives). Wraps Graph's async copy: the tool polls the monitor URL until the copy completes — never returns 'initiated successfully'. Default wait is 300 seconds; configurable via wait_max_seconds.",
  inputSchema: {
    type: "object",
    properties: {
      item_id: { type: "string", description: "Source item ID (file or folder)." },
      target_parent_id: { type: "string", description: "Destination folder ID." },
      source_drive_id: { type: "string", description: "Optional source drive ID. Defaults to /me/drive." },
      target_drive_id: { type: "string", description: "Optional target drive ID. Defaults to the source drive." },
      new_name: { type: "string", description: "Optional rename during copy." },
      wait_max_seconds: {
        type: "integer",
        minimum: 1,
        maximum: 1800,
        description: "Maximum seconds to poll the monitor URL before timing out (default 300).",
      },
    },
    required: ["item_id", "target_parent_id"],
  },
};

interface CopiedItemSummary {
  id: string;
  name: string;
  size: number;
  webUrl: string;
  parent_id: string | null;
}

export function summarizeCopiedItem(item: Record<string, unknown>): CopiedItemSummary {
  const parent = item.parentReference as Record<string, unknown> | undefined;
  return {
    id: String(item.id ?? ""),
    name: String(item.name ?? ""),
    size: Number(item.size ?? 0),
    webUrl: String(item.webUrl ?? ""),
    parent_id: (parent?.id as string | undefined) ?? null,
  };
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export interface MonitorPollResult {
  status: "completed" | "failed" | "in_progress";
  resource_location?: string;
  percent_complete?: number;
  error?: unknown;
  raw: Record<string, unknown>;
}

export function classifyMonitorStatus(payload: Record<string, unknown>): MonitorPollResult {
  const status = String(payload.status ?? "").toLowerCase();
  const result: MonitorPollResult = {
    status: status === "completed" ? "completed" : status === "failed" ? "failed" : "in_progress",
    raw: payload,
  };
  if (typeof payload.resourceLocation === "string") {
    result.resource_location = payload.resourceLocation;
  }
  if (typeof payload.percentageComplete === "number") {
    result.percent_complete = payload.percentageComplete;
  }
  if (payload.error !== undefined) {
    result.error = payload.error;
  }
  return result;
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const itemId = validateRequiredString(args.item_id, "item_id");
  const targetParentId = validateRequiredString(args.target_parent_id, "target_parent_id");
  const sourceDriveId = validateOptionalString(args.source_drive_id, "source_drive_id");
  const targetDriveId = validateOptionalString(args.target_drive_id, "target_drive_id");
  const newName = validateOptionalString(args.new_name, "new_name");
  const waitMaxSeconds = validateOptionalInteger(args.wait_max_seconds, "wait_max_seconds", {
    min: 1,
    max: 1800,
    default: 300,
  });

  const sourcePath = sourceDriveId
    ? `/drives/${encodeURIComponent(sourceDriveId)}/items/${encodeURIComponent(itemId)}`
    : `/me/drive/items/${encodeURIComponent(itemId)}`;

  // Fetch source metadata up-front so we know the source name (needed
  // for the post-completion fallback if the monitor URL doesn't return
  // resourceLocation).
  const sourceItem = (await graph.api(sourcePath).get()) as Record<string, unknown>;
  const effectiveName = newName ?? String(sourceItem.name ?? "");

  const body: Record<string, unknown> = {
    parentReference: targetDriveId
      ? { driveId: targetDriveId, id: targetParentId }
      : { id: targetParentId },
  };
  if (newName !== undefined) body.name = newName;

  // POST → raw response to read the Location header for the monitor URL
  const raw = (await graph
    .api(`${sourcePath}/copy`)
    .responseType(ResponseType.RAW)
    .post(body)) as Response;

  if (raw.status !== 202) {
    throw new Error(`copy: expected 202 Accepted, got ${raw.status} ${raw.statusText}`);
  }
  const monitorUrl = raw.headers.get("location") ?? raw.headers.get("Location");
  if (!monitorUrl) {
    throw new Error("copy: 202 response missing Location header (monitor URL)");
  }

  const deadline = Date.now() + waitMaxSeconds * 1000;
  let pollInterval = POLL_INITIAL_MS;
  let last: MonitorPollResult | null = null;

  while (Date.now() < deadline) {
    await sleep(pollInterval);
    pollInterval = Math.min(pollInterval * 2, POLL_MAX_MS);

    const payload = (await graph.api(monitorUrl).get()) as Record<string, unknown>;
    last = classifyMonitorStatus(payload);

    if (last.status === "completed") {
      let copied: CopiedItemSummary | null = null;
      if (last.resource_location) {
        // Happy path: monitor URL returned a fetchable resourceLocation.
        const newItem = (await graph
          .api(last.resource_location)
          .get()) as Record<string, unknown>;
        copied = summarizeCopiedItem(newItem);
      } else {
        // Graph quirk: completed responses sometimes omit resourceLocation.
        // Fall back to listing the target parent and matching by name.
        const targetDriveBase = targetDriveId
          ? `/drives/${encodeURIComponent(targetDriveId)}`
          : sourceDriveId
            ? `/drives/${encodeURIComponent(sourceDriveId)}`
            : "/me/drive";
        const childrenPath = `${targetDriveBase}/items/${encodeURIComponent(targetParentId)}/children`;
        const listing = (await graph
          .api(childrenPath)
          .filter(`name eq '${effectiveName.replace(/'/g, "''")}'`)
          .get()) as Record<string, unknown>;
        const matches = Array.isArray(listing?.value) ? listing.value : [];
        if (matches.length > 0) {
          copied = summarizeCopiedItem(matches[0] as Record<string, unknown>);
        }
      }
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                status: "completed",
                copied,
                source: { item_id: itemId, drive_id: sourceDriveId ?? null },
              },
              null,
              2,
            ),
          },
        ],
      };
    }

    if (last.status === "failed") {
      throw new Error(
        `copy operation failed: ${JSON.stringify(last.error ?? "(no detail)")}`,
      );
    }
  }

  throw new Error(
    `copy did not complete within ${waitMaxSeconds}s. Last status: ${JSON.stringify(last?.raw ?? "(no poll yet)")}`,
  );
};

export const copyFileTool: Tool = { definition, handler };
