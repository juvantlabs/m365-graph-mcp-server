/** Download retention and M365_DOWNLOAD_MODE=return. */
import { lstat, mkdir, mkdtemp, readdir, rm, symlink, utimes, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { Readable } from "node:stream";

import type { Client } from "@microsoft/microsoft-graph-client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import {
  getDownloadMode,
  getTtlHours,
  readInline,
  RETURN_MAX_BYTES,
  startRetention,
  sweepDownloads,
} from "../../src/tools/_downloads.js";
import { downloadFileTool as tool } from "../../src/tools/download_file.js";

const originalEnv = { ...process.env };
const TENANT = "11111111-2222-3333-4444-555555555555";
const OTHER = "99999999-2222-3333-4444-555555555555";
const HEX = "0123456789abcdef";
let base: string;

async function put(file: string, hoursOld: number): Promise<string> {
  await mkdir(path.dirname(file), { recursive: true });
  await writeFile(file, "x");
  const when = new Date(Date.now() - hoursOld * 3_600_000);
  await utimes(file, when, when);
  return file;
}
const exists = (file: string): Promise<boolean> => lstat(file).then(() => true, () => false);

beforeEach(async () => {
  base = await mkdtemp(path.join(os.tmpdir(), "m365-ret-"));
  process.env.M365_DOWNLOAD_DIR = base;
  process.env.M365_TENANT_ID = TENANT;
  delete process.env.M365_DOWNLOAD_MODE;
  delete process.env.M365_DOWNLOAD_TTL_HOURS;
});
afterEach(async () => {
  process.env = { ...originalEnv };
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
  await rm(base, { recursive: true, force: true });
});

describe("retention", () => {
  it("deletes expired downloads of every tenant and keeps recent ones", async () => {
    const old = await put(path.join(base, TENANT, `${HEX}-a.pdf`), 25);
    const other = await put(path.join(base, OTHER, `${HEX}-b.pdf`), 25);
    const recent = await put(path.join(base, TENANT, `${HEX}-c.pdf`), 23);
    expect(await sweepDownloads(path.join(base, TENANT))).toEqual({ deleted: 2, errors: 0 });
    expect([await exists(old), await exists(other), await exists(recent)]).toEqual([false, false, true]);
  });

  it("ignores foreign files, folders, deeper files and folders that aren't a tenant", async () => {
    const kept = [
      await put(path.join(base, TENANT, "notes.txt"), 99),
      await put(path.join(base, TENANT, `${HEX.toUpperCase()}-a.pdf`), 99),
      await put(path.join(base, TENANT, `${HEX}-dir`, `${HEX}-a.pdf`), 99),
      await put(path.join(base, "Documents", `${HEX}-a.pdf`), 99),
      await put(path.join(base, `${HEX}-a.pdf`), 99),
    ];
    expect((await sweepDownloads(path.join(base, TENANT))).deleted).toBe(0);
    for (const file of kept) expect(await exists(file)).toBe(true);
  });

  it("never follows symlinks: file, tenant folder, base", async () => {
    const outside = await mkdtemp(path.join(os.tmpdir(), "m365-out-"));
    try {
      const a = await put(path.join(outside, `${HEX}-a.pdf`), 99);
      const b = await put(path.join(outside, "base", TENANT, `${HEX}-b.pdf`), 99);
      await mkdir(path.join(base, TENANT));
      const link = path.join(base, TENANT, `${HEX}-link`);
      await symlink(a, link);
      await symlink(outside, path.join(base, OTHER));
      await symlink(path.join(outside, "base"), path.join(base, "linkbase"));
      await sweepDownloads(path.join(base, TENANT), 1, Date.now() + 1e9);
      await sweepDownloads(path.join(base, "linkbase", TENANT));
      expect([await exists(a), await exists(b), await exists(link)]).toEqual([true, true, true]);
    } finally {
      await rm(outside, { recursive: true, force: true });
    }
  });

  it("does nothing when the base is missing", async () => {
    expect(await sweepDownloads(path.join(base, "nope", TENANT))).toEqual({ deleted: 0, errors: 0 });
  });

  it("sweeps at startup, logs counts and no names", async () => {
    const log = vi.spyOn(console, "error").mockImplementation(() => undefined);
    const old = await put(path.join(base, TENANT, `${HEX}-secret-name.pdf`), 25);
    clearInterval(await startRetention(path.join(base, TENANT)));
    expect(await exists(old)).toBe(false);
    expect(log.mock.calls.flat().join(" ")).toContain("1 deleted");
    expect(log.mock.calls.flat().join(" ")).not.toContain("secret-name");
  });

  it.each([
    [undefined, 24], ["", 24], ["abc", 24], ["-1", 24], ["1.5", 24], ["0", 1], ["48", 48], ["99999", 720],
  ])("TTL %s is %d hours", (raw, hours) => {
    expect(getTtlHours({ M365_DOWNLOAD_TTL_HOURS: raw })).toBe(hours);
  });
});

describe("download mode", () => {
  it("accepts file and return, refuses anything else", () => {
    expect(getDownloadMode({})).toBe("file");
    expect(getDownloadMode({ M365_DOWNLOAD_MODE: "return" })).toBe("return");
    expect(() => getDownloadMode({ M365_DOWNLOAD_MODE: "Return" })).toThrow(/must be/);
  });

  it("readInline returns text, base64, and enforces the limit while streaming", async () => {
    expect(await readInline(Readable.from([Buffer.from("ciao")]), "text/plain")).toEqual({
      encoding: "text", content: "ciao", size_bytes: 4,
    });
    expect((await readInline(Readable.from([Buffer.from([0xff])]), "text/plain")).encoding).toBe("base64");
    expect((await readInline(Readable.from([Buffer.from("x")]), "application/pdf")).encoding).toBe("base64");
    const half = Buffer.alloc(RETURN_MAX_BYTES / 2 + 1);
    await expect(readInline(Readable.from([half, half]), null)).rejects.toThrow(/Nothing was written/);
  });
});

describe("download_file", () => {
  const fetched = vi.fn();
  beforeEach(() => fetched.mockReset());

  async function call(size: number): Promise<Record<string, unknown>> {
    const item = {
      id: "01F",
      name: "a.txt",
      size,
      file: { mimeType: "text/plain" },
      "@microsoft.graph.downloadUrl": "https://cdn.example/x",
    };
    const client = { api: () => ({ get: async () => item }) } as unknown as Client;
    fetched.mockImplementation(async () => new Response("hello", { status: 200 }));
    vi.stubGlobal("fetch", fetched);
    const res = await tool.handler(client, { item_id: "01F" });
    return JSON.parse((res.content[0] as { text: string }).text) as Record<string, unknown>;
  }

  it("return mode covers PDF conversion", async () => {
    process.env.M365_DOWNLOAD_MODE = "return";
    const item = { id: "01F", name: "a.docx", size: RETURN_MAX_BYTES * 2, file: {} };
    const query = {
      get: async () => item,
      middlewareOptions: () => ({
        responseType: () => ({ get: async () => new Response("%PDF", { status: 200 }) }),
      }),
    };
    const client = { api: () => query } as unknown as Client;
    const res = await tool.handler(client, { item_id: "01F", format: "pdf" });
    const out = JSON.parse((res.content[0] as { text: string }).text);
    expect(out).toMatchObject({ local_path: null, encoding: "base64", size_bytes: 4, format: "pdf" });
    expect(await readdir(base)).toEqual([]);
  });

  it("sweeps expired downloads after a download", async () => {
    const old = await put(path.join(base, TENANT, `${HEX}-old.pdf`), 25);
    const out = await call(5);
    expect(await exists(old)).toBe(false);
    expect(await exists(out.local_path as string)).toBe(true);
  });

  it("return mode writes nothing", async () => {
    process.env.M365_DOWNLOAD_MODE = "return";
    const out = await call(5);
    expect(out).toMatchObject({ local_path: null, encoding: "text", content: "hello", size_bytes: 5 });
    expect(await readdir(base)).toEqual([]);
  });

  it("return mode refuses a large file before fetching it", async () => {
    process.env.M365_DOWNLOAD_MODE = "return";
    await expect(call(RETURN_MAX_BYTES + 1)).rejects.toThrow(/Nothing was written/);
    expect(fetched).not.toHaveBeenCalled();
    expect(await readdir(base)).toEqual([]);
  });
});
