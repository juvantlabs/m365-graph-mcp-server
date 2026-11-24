/**
 * What happens to downloaded files on this machine.
 *
 * Retention: `download_file` writes plaintext copies into the sandbox. The server
 * deletes its own files once they are older than M365_DOWNLOAD_TTL_HOURS
 * (default 24): at startup, hourly while it runs, and after each download.
 * It cannot clean up while it isn't running.
 *
 * Only this shape is ever deleted: a regular file named `<16 hex>-<name>`,
 * directly inside a tenant-shaped folder of the sandbox base. A base that
 * points at the wrong place therefore finds nothing to delete. Symlinks
 * (base, tenant folder, file) are never followed, folders never removed.
 *
 * M365_DOWNLOAD_MODE=return: nothing is written; the content comes back in
 * the tool response, up to RETURN_MAX_BYTES.
 *
 * Logs carry counts only: file names are sensitive.
 */
import { lstat, readdir, unlink } from "node:fs/promises";
import path from "node:path";

export const DEFAULT_TTL_HOURS = 24;
export const MIN_TTL_HOURS = 1;
export const MAX_TTL_HOURS = 720;
export const RETURN_MAX_BYTES = 2 * 1024 * 1024;
const SWEEP_INTERVAL_MS = 60 * 60 * 1000;

const OWN_FILE_RE = /^[0-9a-f]{16}-./;
const TENANT_DIR_RE =
  /^(unknown|common|organizations|consumers|[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})$/;

export type DownloadMode = "file" | "return";

/** A value that isn't recognised is refused: guessing could write to disk. */
export function getDownloadMode(
  raw: string | undefined = process.env.M365_DOWNLOAD_MODE,
): DownloadMode {
  if (raw === undefined || raw === "" || raw === "file") return "file";
  if (raw === "return") return "return";
  throw new Error("M365_DOWNLOAD_MODE must be 'file' or 'return'.");
}

/** Hours to keep downloads. Out-of-range values are clamped; anything else is the default. */
export function getTtlHours(
  raw: string | undefined = process.env.M365_DOWNLOAD_TTL_HOURS,
): number {
  if (raw === undefined || !/^\d{1,9}$/.test(raw)) return DEFAULT_TTL_HOURS;
  return Math.min(MAX_TTL_HOURS, Math.max(MIN_TTL_HOURS, Number(raw)));
}

export interface SweepResult {
  deleted: number;
  errors: number;
}

async function realDir(dir: string): Promise<boolean> {
  try {
    return (await lstat(dir)).isDirectory();
  } catch {
    return false;
  }
}

/**
 * Delete expired downloads. `tenantRoot` is `<base>/<tenant>`: every
 * tenant-shaped folder under `<base>`, plus this one, is swept. Never throws.
 */
export async function sweepDownloads(
  tenantRoot: string,
  ttlHours: number = getTtlHours(),
  now: number = Date.now(),
): Promise<SweepResult> {
  const result: SweepResult = { deleted: 0, errors: 0 };
  const base = path.dirname(tenantRoot);
  if (!(await realDir(base))) return result;

  let tenants: string[];
  try {
    tenants = (await readdir(base)).filter(
      (name) => TENANT_DIR_RE.test(name) || name === path.basename(tenantRoot),
    );
  } catch {
    result.errors += 1;
    return result;
  }

  for (const tenant of tenants) {
    const dir = path.join(base, tenant);
    if (!(await realDir(dir))) continue;
    let names: string[];
    try {
      names = (await readdir(dir)).filter((name) => OWN_FILE_RE.test(name));
    } catch {
      result.errors += 1;
      continue;
    }
    for (const name of names) {
      const file = path.join(dir, name);
      try {
        const info = await lstat(file);
        if (!info.isFile() || now - info.mtimeMs <= ttlHours * 3_600_000) continue;
        await unlink(file);
        result.deleted += 1;
      } catch (err) {
        // Another server process may have removed it first.
        if ((err as NodeJS.ErrnoException).code !== "ENOENT") result.errors += 1;
      }
    }
  }
  return result;
}

async function sweepAndReport(tenantRoot: string): Promise<void> {
  const { deleted, errors } = await sweepDownloads(tenantRoot);
  if (deleted > 0 || errors > 0) {
    console.error(`[m365-graph-mcp-server] download retention: ${deleted} deleted, ${errors} errors`);
  }
}

/** Sweep now and hourly. The timer does not keep the process alive. */
export async function startRetention(tenantRoot: string): Promise<NodeJS.Timeout> {
  await sweepAndReport(tenantRoot);
  const timer = setInterval(() => void sweepAndReport(tenantRoot), SWEEP_INTERVAL_MS);
  timer.unref();
  return timer;
}

export interface InlineContent {
  encoding: "text" | "base64";
  content: string;
  size_bytes: number;
}

const TEXT_TYPE_RE = /^(text\/|application\/(json|xml)\b)/i;

/** Collect a stream in memory, refusing more than RETURN_MAX_BYTES. */
export async function readInline(
  chunks: AsyncIterable<Uint8Array | string>,
  contentType: string | null,
): Promise<InlineContent> {
  const parts: Buffer[] = [];
  let size = 0;
  for await (const chunk of chunks) {
    const part = typeof chunk === "string" ? Buffer.from(chunk) : Buffer.from(chunk);
    size += part.length;
    if (size > RETURN_MAX_BYTES) throw new Error(tooLargeToReturn());
    parts.push(part);
  }
  const bytes = Buffer.concat(parts);
  if (contentType !== null && TEXT_TYPE_RE.test(contentType)) {
    try {
      const content = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
      return { encoding: "text", content, size_bytes: size };
    } catch {
      // Not UTF-8: fall through to base64.
    }
  }
  return { encoding: "base64", content: bytes.toString("base64"), size_bytes: size };
}

export function tooLargeToReturn(): string {
  return `Content exceeds the ${RETURN_MAX_BYTES}-byte limit of M365_DOWNLOAD_MODE=return. Nothing was written.`;
}
