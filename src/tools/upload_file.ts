/**
 * Tool: m365-graph:upload_file
 *
 * Upload a local file to a drive. Auto-routes between two paths based
 * on file size:
 *   - ≤ 4 MB → single PUT to `/items/{parent}:/{name}:/content`
 *   - > 4 MB → resumable upload session via OneDriveLargeFileUploadTask
 *     (10 MB chunks, automatic resume on transient errors)
 *
 * Required Graph scope: `Files.ReadWrite` (delegated). Writes to the
 * cloud side. Reads from the caller-supplied `local_path` on the
 * server's filesystem.
 *
 * Trust note: the agent supplies `local_path`. The MCP server runs as
 * the user; the user trusts the agent. There's no server-side path
 * sandbox for *uploads* (unlike downloads, which we sandbox to a
 * known dir): an upload tool that refused arbitrary paths would be
 * useless for the canonical "upload this file the user prepared"
 * workflow. The absolute path is logged to stderr so operators can
 * review via the MCP server logs.
 *
 * Size cap: 200 MB (handbook spec § Performance). Refused pre-read.
 *
 * Conflict behavior is parametric:
 *   - `fail`    (default) — refuse if a file with that name exists
 *   - `replace` — overwrite the existing file (creates a new version)
 *   - `rename`  — let Graph append a numeric suffix
 */

import fs from "node:fs/promises";
import { createReadStream } from "node:fs";
import path from "node:path";

import {
  type Client,
  type LargeFileUploadSession,
  type LargeFileUploadTaskOptions,
  OneDriveLargeFileUploadTask,
  StreamUpload,
} from "@microsoft/microsoft-graph-client";

import {
  validateOptionalEnum,
  validateOptionalString,
  validateRequiredString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const SMALL_FILE_CAP = 4 * 1024 * 1024; // 4 MB — single-PUT threshold per Graph docs
export const MAX_FILE_SIZE = 200 * 1024 * 1024; // 200 MB hard cap (handbook spec)
const LARGE_FILE_RANGE = 10 * 1024 * 1024; // 10 MB chunks for resumable

/**
 * Throw if the file size exceeds the configured maximum. Extracted so
 * this defense-in-depth check can be unit-tested independently of the
 * actual fs + Graph integration in the handler.
 */
export function checkSizeCap(size: number): void {
  if (size > MAX_FILE_SIZE) {
    throw new Error(
      `File size (${size} bytes) exceeds the 200 MB cap. Refusing to upload.`,
    );
  }
}

const CONFLICT_BEHAVIORS = ["fail", "replace", "rename"] as const;
type ConflictBehavior = (typeof CONFLICT_BEHAVIORS)[number];

const definition: ToolDefinition = {
  name: "m365-graph:upload_file",
  description:
    "Upload a local file to a drive (OneDrive primary by default; SharePoint via drive_id). Auto-routes between single PUT (≤ 4 MB) and resumable upload (> 4 MB, 10 MB chunks) based on file size. Default conflict behavior is 'fail' — set to 'replace' or 'rename' to overwrite or auto-rename. 200 MB hard cap.",
  inputSchema: {
    type: "object",
    properties: {
      local_path: {
        type: "string",
        description: "Absolute or relative path to the local file to upload. Resolved to absolute before reading.",
      },
      drive_id: {
        type: "string",
        description: "Optional drive ID. Defaults to the user's primary OneDrive.",
      },
      parent_item_id: {
        type: "string",
        description: "Optional folder ID to upload INTO. Defaults to the drive root. Get from list_items.",
      },
      name: {
        type: "string",
        description: "Optional name for the uploaded file. Defaults to the basename of local_path.",
      },
      conflict_behavior: {
        type: "string",
        enum: ["fail", "replace", "rename"],
        description: "What to do if a file with this name already exists. Default: 'fail'.",
      },
    },
    required: ["local_path"],
  },
};

interface UploadResult {
  drive_id: string | null;
  parent_item_id: string | null;
  uploaded: {
    id: string;
    name: string;
    size: number;
    webUrl: string;
    upload_path: "single_put" | "resumable_session";
  };
}

function summarizeUploadedItem(item: Record<string, unknown>): { id: string; name: string; size: number; webUrl: string } {
  return {
    id: String(item.id ?? ""),
    name: String(item.name ?? ""),
    size: Number(item.size ?? 0),
    webUrl: String(item.webUrl ?? ""),
  };
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const localPathRaw = validateRequiredString(args.local_path, "local_path");
  const localPath = path.resolve(localPathRaw);
  const driveId = validateOptionalString(args.drive_id, "drive_id");
  const parentItemId = validateOptionalString(args.parent_item_id, "parent_item_id");
  const customName = validateOptionalString(args.name, "name");
  const conflictBehavior = validateOptionalEnum<ConflictBehavior>(
    args.conflict_behavior,
    "conflict_behavior",
    CONFLICT_BEHAVIORS,
    "fail",
  );

  const stat = await fs.stat(localPath);
  if (!stat.isFile()) {
    throw new Error(`local_path is not a regular file: ${localPath}`);
  }
  checkSizeCap(stat.size);

  const filename = customName ?? path.basename(localPath);
  const drivePath = driveId ? `/drives/${encodeURIComponent(driveId)}` : "/me/drive";
  const parentApi = parentItemId
    ? `${drivePath}/items/${encodeURIComponent(parentItemId)}`
    : `${drivePath}/root`;

  console.error(
    `[m365-graph-mcp-server] upload_file: ${localPath} (${stat.size} bytes) → ${drivePath}/${parentItemId ?? "<root>"}/${filename}`,
  );

  let uploadedItem: Record<string, unknown>;
  let uploadPath: "single_put" | "resumable_session";

  if (stat.size <= SMALL_FILE_CAP) {
    // Small file: single PUT to /items/{parent}:/{name}:/content
    const content = await fs.readFile(localPath);
    uploadedItem = (await graph
      .api(`${parentApi}:/${encodeURIComponent(filename)}:/content`)
      .query({ "@microsoft.graph.conflictBehavior": conflictBehavior })
      .put(content)) as Record<string, unknown>;
    uploadPath = "single_put";
  } else {
    // Large file: resumable upload session, 10 MB chunks.
    const stream = createReadStream(localPath);
    const fileObject = new StreamUpload(stream, filename, stat.size);

    // Graph upload-session URL pattern:
    //   /me/drive/root:/<name>:/createUploadSession                  (root)
    //   /me/drive/items/<parent-id>:/<name>:/createUploadSession     (sub-folder)
    //   /drives/<drive-id>/...                                       (non-default drive)
    const sessionUrl = parentItemId
      ? `${parentApi}:/${encodeURIComponent(filename)}:/createUploadSession`
      : `${parentApi}:/${encodeURIComponent(filename)}:/createUploadSession`;

    const taskOptions: LargeFileUploadTaskOptions = {
      rangeSize: LARGE_FILE_RANGE,
    };

    const uploadSession: LargeFileUploadSession =
      await OneDriveLargeFileUploadTask.createUploadSession(graph, sessionUrl, {
        fileName: filename,
        conflictBehavior,
      });

    const uploadTask = new OneDriveLargeFileUploadTask(
      graph,
      fileObject,
      uploadSession,
      taskOptions,
    );
    const taskResult = await uploadTask.upload();
    uploadedItem = (taskResult.responseBody ?? {}) as Record<string, unknown>;
    uploadPath = "resumable_session";
  }

  const result: UploadResult = {
    drive_id: driveId ?? null,
    parent_item_id: parentItemId ?? null,
    uploaded: {
      ...summarizeUploadedItem(uploadedItem),
      upload_path: uploadPath,
    },
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
};

export const uploadFileTool: Tool = { category: "write_idempotent", definition, handler };
