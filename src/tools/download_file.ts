/**
 * Tool: m365-graph:download_file
 *
 * Download a file from a drive to a per-tenant local sandbox directory,
 * returning the local path. The agent can then read the file via any
 * filesystem-aware MCP (or the agent's built-in file ops) — this tool
 * does not return the file content directly.
 *
 * Why path return, not content:
 *   - Files can be large (handbook spec § Performance: 200 MB cap).
 *   - MCP responses are JSON-typed; binary content via base64 inflates
 *     by 33% and stresses the agent's context.
 *   - Composability: M365-aware fetch + provider-agnostic read are
 *     two concerns that cleanly separate.
 *
 * Sandboxing (handbook spec § Sandboxing + anti-pattern #1):
 *   - All downloads land under <sandbox-root>/<tenant-id>/ where the
 *     filename is constructed as <sha256(item_id)[:16]>-<sanitized name>.
 *     The local filename is server-controlled, NOT derived from
 *     caller-supplied paths. Path injection is structurally impossible.
 *   - sandbox-root is M365_DOWNLOAD_DIR (env), or
 *     $XDG_CACHE_HOME/m365-graph-mcp-server, or ~/.cache/m365-graph-mcp-server.
 *   - 0o700 dir mode + 0o600 file mode (defense-in-depth on shared hosts).
 *   - Resolved path is verified to start with the sandbox root before
 *     writing — an extra check on top of the structurally-safe naming.
 *
 * Size cap: 200 MB (handbook spec § Performance characteristics). Item
 * size is fetched via `/items/{id}` metadata BEFORE downloading; the
 * tool refuses to stream large files.
 *
 * Streaming: the actual download streams from the Graph CDN
 * (@microsoft.graph.downloadUrl is a pre-signed Akamai URL — no Graph
 * auth needed for the bytes themselves). Streamed via fetch + Node
 * streams pipeline (no whole-file buffering — handbook anti-pattern #11).
 *
 * Required Graph scope: `Files.Read` (delegated). Read-only on the
 * cloud side; writes only to the sandbox locally.
 */

import { createWriteStream } from "node:fs";
import { mkdir } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import { Readable } from "node:stream";
import { pipeline } from "node:stream/promises";

import type { Client } from "@microsoft/microsoft-graph-client";

import {
  sanitizeFilename,
  validateOptionalString,
  validateRequiredString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const MAX_FILE_SIZE_BYTES = 200 * 1024 * 1024; // 200 MB per handbook spec
const SANDBOX_DIR_MODE = 0o700;
const SANDBOX_FILE_MODE = 0o600;

const definition: ToolDefinition = {
  name: "m365-graph:download_file",
  description:
    "Download a file from a drive to a local sandbox directory. Returns the local path, not the file content — agents read the file via a filesystem-aware tool. Size is capped at 200 MB. Read-only on the cloud side.",
  inputSchema: {
    type: "object",
    properties: {
      item_id: {
        type: "string",
        description:
          "ID of the file to download. Obtain from list_items or search_files. Folders are rejected (no content).",
      },
      drive_id: {
        type: "string",
        description: "Optional drive ID. Defaults to the user's primary OneDrive.",
      },
    },
    required: ["item_id"],
  },
};

export function getSandboxRoot(tenantId: string): string {
  const override = process.env.M365_DOWNLOAD_DIR;
  if (override) return path.resolve(override, tenantId);

  const xdg = process.env.XDG_CACHE_HOME;
  if (xdg) return path.resolve(xdg, "m365-graph-mcp-server", tenantId);

  return path.resolve(os.homedir(), ".cache", "m365-graph-mcp-server", tenantId);
}

export function deriveSafeLocalPath(sandboxRoot: string, itemId: string, originalName: string): string {
  const itemHash = crypto.createHash("sha256").update(itemId).digest("hex").slice(0, 16);
  const safe = sanitizeFilename(originalName);
  const filename = `${itemHash}-${safe}`;
  const resolved = path.resolve(sandboxRoot, filename);

  // Defense-in-depth even though the filename is server-constructed.
  if (!resolved.startsWith(sandboxRoot + path.sep) && resolved !== sandboxRoot) {
    throw new Error(`Refusing to write outside sandbox: ${resolved}`);
  }
  return resolved;
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const itemId = validateRequiredString(args.item_id, "item_id");
  const driveId = validateOptionalString(args.drive_id, "drive_id");

  const driveRoot = driveId ? `/drives/${encodeURIComponent(driveId)}` : "/me/drive";
  const item = await graph.api(`${driveRoot}/items/${encodeURIComponent(itemId)}`).get();

  if (item.folder !== undefined) {
    throw new Error(`Item ${itemId} is a folder, not a file — refusing to download.`);
  }

  const size = Number(item.size ?? 0);
  if (size > MAX_FILE_SIZE_BYTES) {
    throw new Error(
      `File size (${size} bytes) exceeds the 200 MB cap. Refusing to download.`,
    );
  }

  const downloadUrl = item["@microsoft.graph.downloadUrl"] as string | undefined;
  if (!downloadUrl) {
    throw new Error(
      "No @microsoft.graph.downloadUrl in item metadata — Graph returned an unexpected shape.",
    );
  }

  const tenantId = process.env.M365_TENANT_ID ?? "unknown";
  const sandboxRoot = getSandboxRoot(tenantId);
  await mkdir(sandboxRoot, { recursive: true, mode: SANDBOX_DIR_MODE });

  const localPath = deriveSafeLocalPath(sandboxRoot, itemId, String(item.name ?? "file"));

  const response = await fetch(downloadUrl);
  if (!response.ok) {
    throw new Error(`Download HTTP ${response.status}: ${response.statusText}`);
  }
  if (!response.body) {
    throw new Error("Download response has no body stream.");
  }

  const out = createWriteStream(localPath, { mode: SANDBOX_FILE_MODE });
  // Node 20+ Web→Node stream interop. The `as ReadableStream` cast is
  // because node:stream's Readable.fromWeb expects the Node-bundled
  // ReadableStream, but undici's fetch returns the same shape.
  await pipeline(Readable.fromWeb(response.body as never), out);

  const result = {
    item_id: itemId,
    drive_id: driveId ?? null,
    local_path: localPath,
    size_bytes: size,
    name: String(item.name ?? ""),
    content_type: (item.file as Record<string, unknown> | undefined)?.mimeType ?? null,
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
};

export const downloadFileTool: Tool = { category: "read", definition, handler };
