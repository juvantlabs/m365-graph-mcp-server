/**
 * download_file — plain (no conversion) streaming path, end-to-end into
 * a temporary sandbox with a stubbed fetch for the pre-signed CDN URL.
 */
import { mkdtemp, readdir, readFile, rm, stat } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import type { Client } from "@microsoft/microsoft-graph-client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { downloadFileTool } from "../../src/tools/download_file.js";

const originalEnv = { ...process.env };
let sandboxBase: string;

beforeEach(async () => {
  sandboxBase = await mkdtemp(path.join(os.tmpdir(), "m365-dl-stream-"));
  process.env.M365_DOWNLOAD_DIR = sandboxBase;
  process.env.M365_TENANT_ID = "tenant-s";
});

afterEach(async () => {
  process.env = { ...originalEnv };
  vi.unstubAllGlobals();
  await rm(sandboxBase, { recursive: true, force: true });
});

function clientReturning(item: unknown): { client: Client; paths: string[] } {
  const paths: string[] = [];
  const api = vi.fn((p: string) => {
    paths.push(p);
    return { get: vi.fn().mockResolvedValue(item) };
  });
  return { client: { api } as unknown as Client, paths };
}

const fileItem = {
  id: "01FILE",
  name: "report.xlsx",
  size: 5,
  file: { mimeType: "application/vnd.ms-excel" },
  "@microsoft.graph.downloadUrl": "https://cdn.example/presigned?sig=1",
};

describe("downloadFileTool handler — plain streaming download", () => {
  it("streams the CDN bytes into <sandbox>/<tenant>/<hash>-<name> with mode 0600", async () => {
    const { client, paths } = clientReturning(fileItem);
    const fetchMock = vi.fn(async () => new Response(Buffer.from("hello"), { status: 200 }));
    vi.stubGlobal("fetch", fetchMock);

    const res = await downloadFileTool.handler(client, { item_id: "01FILE", drive_id: "b!lib" });
    const out = JSON.parse((res.content[0] as { text: string }).text);

    expect(paths).toEqual(["/drives/b!lib/items/01FILE"]);
    // Pre-signed URL fetched with no init — no bearer token leaves the process.
    expect(fetchMock.mock.calls[0]).toEqual(["https://cdn.example/presigned?sig=1"]);
    const sandbox = path.join(sandboxBase, "tenant-s");
    expect(out.local_path).toMatch(new RegExp(`^${sandbox}/[0-9a-f]{16}-report\\.xlsx$`));
    expect(await readFile(out.local_path, "utf8")).toBe("hello");
    expect((await stat(out.local_path)).mode & 0o777).toBe(0o600);
    expect((await stat(sandbox)).mode & 0o777).toBe(0o700);
    expect(out).toEqual({
      item_id: "01FILE",
      drive_id: "b!lib",
      local_path: out.local_path,
      size_bytes: 5,
      name: "report.xlsx",
      content_type: "application/vnd.ms-excel",
    });
  });

  it("defaults to /me/drive and null drive_id when no drive_id is given", async () => {
    const { client, paths } = clientReturning(fileItem);
    vi.stubGlobal("fetch", vi.fn(async () => new Response("hello", { status: 200 })));
    const res = await downloadFileTool.handler(client, { item_id: "01FILE" });
    expect(paths).toEqual(["/me/drive/items/01FILE"]);
    expect(JSON.parse((res.content[0] as { text: string }).text).drive_id).toBeNull();
  });

  it("surfaces a non-OK CDN response and writes nothing", async () => {
    const { client } = clientReturning(fileItem);
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => new Response("denied", { status: 403, statusText: "Forbidden" })),
    );
    await expect(downloadFileTool.handler(client, { item_id: "01FILE" })).rejects.toThrow(
      "Download HTTP 403: Forbidden",
    );
    expect(await readdir(path.join(sandboxBase, "tenant-s"))).toEqual([]);
  });

  it("rejects a CDN response without a body stream", async () => {
    const { client } = clientReturning(fileItem);
    vi.stubGlobal("fetch", vi.fn(async () => new Response(null, { status: 200 })));
    await expect(downloadFileTool.handler(client, { item_id: "01FILE" })).rejects.toThrow(
      "no body stream",
    );
  });

  it("falls back to null content_type and 'file' name when metadata omits them", async () => {
    const { client } = clientReturning({
      id: "x",
      size: 1,
      "@microsoft.graph.downloadUrl": "https://cdn.example/x",
    });
    vi.stubGlobal("fetch", vi.fn(async () => new Response("1", { status: 200 })));
    const res = await downloadFileTool.handler(client, { item_id: "x" });
    const out = JSON.parse((res.content[0] as { text: string }).text);
    expect(out.content_type).toBeNull();
    expect(out.local_path).toMatch(/-file$/);
  });
});
