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
 *   - Same item → same local path; a repeat download overwrites the
 *     previous copy (deterministic, no accumulation).
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
 * Server-side format conversion (`format`, optional):
 *   - `format: "pdf"` asks Graph to convert the file server-side via
 *     `GET {drive}/items/{id}/content?format=pdf`. No local converter
 *     runs — the server never executes anything on the file.
 *   - `format` is a strict enum (`DOWNLOAD_FORMATS`); the query value
 *     is always the server-side constant, never the raw caller string.
 *   - Source type is checked pre-flight against the extensions Graph
 *     documents as PDF-convertible; unsupported types fail fast with a
 *     clear error and no network call beyond the metadata GET.
 *   - Graph answers with a 302 to a pre-authenticated URL for the
 *     converted bytes. The SDK's automatic redirect-following is
 *     disabled for this request: we read `Location` ourselves, require
 *     https, and fetch it WITHOUT the bearer token (same trust model as
 *     `@microsoft.graph.downloadUrl`).
 *   - The converted size is not known in advance, so the 200 MB cap is
 *     also enforced on the stream (Content-Length when present, plus a
 *     byte counter). A capped or failed conversion leaves no partial
 *     file behind.
 *   - Local filename: `<sha256(item_id)[:16]>-<stem>.pdf` — same
 *     sandbox, same naming and overwrite policy as plain downloads.
 *
 * Required Graph scope: `Files.Read` (delegated; `Sites.Read(Write).All`
 * for SharePoint drives). Read-only on the cloud side; writes only to
 * the sandbox locally. Conversion needs no additional scope.
 */

import { createWriteStream } from "node:fs";
import { mkdir, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import crypto from "node:crypto";
import { Readable, Transform, type TransformCallback } from "node:stream";
import { pipeline } from "node:stream/promises";

import {
  RedirectHandlerOptions,
  ResponseType,
  type Client,
} from "@microsoft/microsoft-graph-client";

import {
  sanitizeFilename,
  validateHttpsUrl,
  validateOptionalStrictEnum,
  validateOptionalString,
  validateRequiredString,
} from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const MAX_FILE_SIZE_BYTES = 200 * 1024 * 1024; // 200 MB per handbook spec
const SANDBOX_DIR_MODE = 0o700;
const SANDBOX_FILE_MODE = 0o600;
/** Mirrors sanitizeFilename's length cap so a converted name keeps its extension. */
const MAX_LOCAL_NAME_CHARS = 200;
/** Upper bound on how much of a Graph error body we read to build a message. */
const MAX_ERROR_BODY_CHARS = 4096;

/** Server-side conversion targets accepted by the `format` parameter. */
export const DOWNLOAD_FORMATS = ["pdf"] as const;
export type DownloadFormat = (typeof DOWNLOAD_FORMATS)[number];

/**
 * Source extensions Microsoft Graph documents as convertible with
 * `?format=pdf` ("Download a file in another format").
 */
export const PDF_CONVERTIBLE_EXTENSIONS: ReadonlySet<string> = new Set([
  "doc",
  "docx",
  "dot",
  "dotx",
  "dotm",
  "dsn",
  "dwg",
  "eml",
  "epub",
  "fluidframework",
  "htm",
  "html",
  "loop",
  "markdown",
  "md",
  "msg",
  "odp",
  "ods",
  "odt",
  "pps",
  "ppsx",
  "ppt",
  "pptx",
  "rtf",
  "tif",
  "tiff",
  "xls",
  "xlsm",
  "xlsx",
]);

const CONVERTIBLE_EXTENSIONS: Record<DownloadFormat, ReadonlySet<string>> = {
  pdf: PDF_CONVERTIBLE_EXTENSIONS,
};

const FORMAT_CONTENT_TYPES: Record<DownloadFormat, string> = {
  pdf: "application/pdf",
};

const REDIRECT_STATUSES = new Set([301, 302, 303, 307, 308]);

const definition: ToolDefinition = {
  name: "m365-graph:download_file",
  description:
    "Download a file from a drive to a local sandbox directory. Returns the local path, not the file content — agents read the file via a filesystem-aware tool. Size is capped at 200 MB. Optional `format: \"pdf\"` has Microsoft Graph convert the file to PDF server-side (Word, Excel, PowerPoint, OpenDocument, RTF, HTML, Markdown, email, TIFF, …) and saves the PDF instead of the original. Read-only on the cloud side.",
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
      format: {
        type: "string",
        enum: [...DOWNLOAD_FORMATS],
        description:
          "Optional server-side conversion. 'pdf' converts the file to PDF in Microsoft Graph and saves it as '<name>.pdf' in the sandbox. Omit to download the original bytes. Unsupported source types (e.g. images other than TIFF, archives, existing PDFs) are rejected with an error.",
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

/**
 * Lower-cased extension without the dot, or null when the name has
 * none. A leading dot alone (".profile") is not an extension.
 */
export function fileExtension(name: string): string | null {
  const dot = name.lastIndexOf(".");
  if (dot <= 0 || dot === name.length - 1) return null;
  return name.slice(dot + 1).toLowerCase();
}

/**
 * Throw a caller-actionable error when Graph cannot convert this source
 * type to the requested format.
 */
export function checkConvertible(name: string, format: DownloadFormat): void {
  const ext = fileExtension(name);
  const supported = CONVERTIBLE_EXTENSIONS[format];
  if (ext !== null && supported.has(ext)) return;

  const list = [...supported].join(", ");
  if (ext === format) {
    throw new Error(
      `'${name}' is already a ${format.toUpperCase()} — call download_file without 'format' to fetch it as-is.`,
    );
  }
  const what = ext === null ? "has no file extension" : `has extension '.${ext}'`;
  throw new Error(
    `Cannot convert '${name}' to ${format.toUpperCase()}: the file ${what}, which Microsoft Graph does not convert to ${format}. Supported source types: ${list}.`,
  );
}

/**
 * Graph path for the converted content of an item. The `format` value
 * comes from the enum constant, never from raw caller input.
 */
export function buildConvertedContentPath(
  driveRoot: string,
  itemId: string,
  format: DownloadFormat,
): string {
  return `${driveRoot}/items/${encodeURIComponent(itemId)}/content?format=${encodeURIComponent(format)}`;
}

/**
 * Local name for a converted file: original stem + `.<format>`. The stem
 * is trimmed so the sanitizer's length cap can never cut the extension.
 */
export function convertedFilename(originalName: string, format: DownloadFormat): string {
  const ext = fileExtension(originalName);
  const stem = ext === null ? originalName : originalName.slice(0, -(ext.length + 1));
  const suffix = `.${format}`;
  const trimmed = stem.slice(0, MAX_LOCAL_NAME_CHARS - suffix.length);
  return `${trimmed.length > 0 ? trimmed : "file"}${suffix}`;
}

/**
 * Byte-counting pass-through that errors once `limit` is exceeded.
 * `bytes()` reports how many bytes have passed through.
 */
export function cappedCounter(limit: number): { stream: Transform; bytes: () => number } {
  let seen = 0;
  const stream = new Transform({
    transform(chunk: Buffer, _enc: BufferEncoding, cb: TransformCallback) {
      seen += chunk.length;
      if (seen > limit) {
        cb(new Error("Converted file exceeds the 200 MB cap. Download aborted."));
        return;
      }
      cb(null, chunk);
    },
  });
  return { stream, bytes: () => seen };
}

async function readErrorSummary(raw: Response): Promise<string> {
  let text = "";
  try {
    text = (await raw.text()).slice(0, MAX_ERROR_BODY_CHARS);
  } catch {
    return "";
  }
  try {
    const parsed = JSON.parse(text) as { error?: { code?: unknown; message?: unknown } };
    const code = typeof parsed.error?.code === "string" ? parsed.error.code : "";
    const message = typeof parsed.error?.message === "string" ? parsed.error.message : "";
    return [code, message].filter(Boolean).join(": ");
  } catch {
    return text.trim().slice(0, 300);
  }
}

/**
 * Request the converted content from Graph and return a response whose
 * body is the converted bytes. Handles Graph's 302-to-pre-authenticated
 * URL contract explicitly; tolerates a direct 200 with the body.
 */
async function fetchConverted(
  graph: Client,
  contentPath: string,
  name: string,
  format: DownloadFormat,
): Promise<Response> {
  const raw = (await graph
    .api(contentPath)
    // Do not let the SDK follow the redirect: we validate the target and
    // fetch it ourselves without the bearer token.
    .middlewareOptions([new RedirectHandlerOptions(0, () => false)])
    .responseType(ResponseType.RAW)
    .get()) as Response;

  if (REDIRECT_STATUSES.has(raw.status)) {
    const target = validateHttpsUrl(
      raw.headers.get("location"),
      "PDF conversion: Graph redirect Location",
    );
    await raw.body?.cancel().catch(() => undefined);
    const converted = await fetch(target);
    if (!converted.ok) {
      throw new Error(
        `PDF conversion download HTTP ${converted.status}: ${converted.statusText}`,
      );
    }
    return converted;
  }

  if (raw.status === 200) return raw;

  const detail = await readErrorSummary(raw);
  const suffix = detail ? ` (${detail})` : "";
  if (raw.status === 406 || raw.status === 415 || /notSupported|unsupported/i.test(detail)) {
    throw new Error(
      `Microsoft Graph could not convert '${name}' to ${format.toUpperCase()}: HTTP ${raw.status}${suffix}. The file may be encrypted, corrupt, or of a type Graph cannot convert.`,
    );
  }
  throw new Error(`PDF conversion HTTP ${raw.status}${suffix}`);
}

async function writeCapped(response: Response, localPath: string): Promise<number> {
  const declared = Number(response.headers.get("content-length") ?? NaN);
  if (Number.isFinite(declared) && declared > MAX_FILE_SIZE_BYTES) {
    await response.body?.cancel().catch(() => undefined);
    throw new Error(
      `Converted file size (${declared} bytes) exceeds the 200 MB cap. Refusing to download.`,
    );
  }
  if (!response.body) {
    throw new Error("Download response has no body stream.");
  }

  const counter = cappedCounter(MAX_FILE_SIZE_BYTES);
  const out = createWriteStream(localPath, { mode: SANDBOX_FILE_MODE });
  try {
    await pipeline(Readable.fromWeb(response.body as never), counter.stream, out);
  } catch (err) {
    // No partial artifact left in the sandbox.
    await rm(localPath, { force: true }).catch(() => undefined);
    throw err;
  }
  return counter.bytes();
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const itemId = validateRequiredString(args.item_id, "item_id");
  const driveId = validateOptionalString(args.drive_id, "drive_id");
  const format = validateOptionalStrictEnum(args.format, "format", DOWNLOAD_FORMATS);

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

  const itemName = String(item.name ?? "");
  const sourceContentType =
    (item.file as Record<string, unknown> | undefined)?.mimeType ?? null;

  if (format !== undefined) {
    checkConvertible(itemName, format);

    const tenantId = process.env.M365_TENANT_ID ?? "unknown";
    const sandboxRoot = getSandboxRoot(tenantId);
    await mkdir(sandboxRoot, { recursive: true, mode: SANDBOX_DIR_MODE });
    const localPath = deriveSafeLocalPath(
      sandboxRoot,
      itemId,
      convertedFilename(itemName, format),
    );

    const response = await fetchConverted(
      graph,
      buildConvertedContentPath(driveRoot, itemId, format),
      itemName,
      format,
    );
    const written = await writeCapped(response, localPath);

    const result = {
      item_id: itemId,
      drive_id: driveId ?? null,
      local_path: localPath,
      size_bytes: written,
      name: itemName,
      content_type: FORMAT_CONTENT_TYPES[format],
      format,
      source_size_bytes: size,
      source_content_type: sourceContentType,
    };
    return {
      content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
    };
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
    name: itemName,
    content_type: sourceContentType,
    format: null,
  };

  return {
    content: [{ type: "text", text: JSON.stringify(result, null, 2) }],
  };
};

export const downloadFileTool: Tool = { category: "read", definition, handler };
