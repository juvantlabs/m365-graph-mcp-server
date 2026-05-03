import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import type { Client } from "@microsoft/microsoft-graph-client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { MAX_FILE_SIZE, checkSizeCap, uploadFileTool } from "../../src/tools/upload_file.js";

let tmp: string;

beforeEach(async () => {
  tmp = await mkdtemp(path.join(tmpdir(), "upload-test-"));
});

afterEach(async () => {
  await rm(tmp, { recursive: true, force: true });
});

async function writeTempFile(name: string, content: string | Buffer): Promise<string> {
  const p = path.join(tmp, name);
  await writeFile(p, content);
  return p;
}

function captureSmallPut(returnValue: unknown): {
  apiCalls: string[];
  queries: unknown[];
  putBodies: unknown[];
  client: Client;
} {
  const apiCalls: string[] = [];
  const queries: unknown[] = [];
  const putBodies: unknown[] = [];
  const put = vi.fn().mockImplementation((body: unknown) => {
    putBodies.push(body);
    return Promise.resolve(returnValue);
  });
  const query = vi.fn().mockImplementation((q: unknown) => {
    queries.push(q);
    return { put };
  });
  const api = vi.fn().mockImplementation((path: string) => {
    apiCalls.push(path);
    return { query };
  });
  return { apiCalls, queries, putBodies, client: { api } as unknown as Client };
}

describe("checkSizeCap", () => {
  it("passes for files at the boundary", () => {
    expect(() => checkSizeCap(MAX_FILE_SIZE)).not.toThrow();
  });

  it("throws for files exceeding the 200 MB cap", () => {
    expect(() => checkSizeCap(MAX_FILE_SIZE + 1)).toThrow("exceeds the 200 MB cap");
  });

  it("throws for absurdly large sizes", () => {
    expect(() => checkSizeCap(10 * 1024 * 1024 * 1024)).toThrow();
  });
});



describe("uploadFileTool handler — small file path", () => {
  it("requires local_path", async () => {
    const { client } = captureSmallPut({ id: "x", name: "x", size: 0, webUrl: "" });
    await expect(uploadFileTool.handler(client, {})).rejects.toThrow(
      "'local_path' must be a non-empty string",
    );
  });

  it("rejects local_path pointing at a directory", async () => {
    const { client } = captureSmallPut({});
    await expect(
      uploadFileTool.handler(client, { local_path: tmp }),
    ).rejects.toThrow("not a regular file");
  });

  it("rejects local_path pointing at a non-existent file", async () => {
    const { client } = captureSmallPut({});
    await expect(
      uploadFileTool.handler(client, { local_path: path.join(tmp, "missing.txt") }),
    ).rejects.toThrow();
  });

  it("uploads a small file via single PUT to /me/drive/root:/<name>:/content", async () => {
    const localPath = await writeTempFile("hello.txt", "hello world");
    const remoteResp = {
      id: "uploaded-id",
      name: "hello.txt",
      size: 11,
      webUrl: "https://example/uploaded",
    };
    const { apiCalls, queries, client } = captureSmallPut(remoteResp);

    const resp = await uploadFileTool.handler(client, { local_path: localPath });
    expect(resp.isError).toBeFalsy();
    expect(apiCalls[0]).toBe("/me/drive/root:/hello.txt:/content");
    expect(queries[0]).toEqual({ "@microsoft.graph.conflictBehavior": "fail" });

    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.uploaded.id).toBe("uploaded-id");
    expect(parsed.uploaded.upload_path).toBe("single_put");
  });

  it("uses parent_item_id when provided", async () => {
    const localPath = await writeTempFile("h.txt", "hi");
    const { apiCalls, client } = captureSmallPut({});
    await uploadFileTool.handler(client, {
      local_path: localPath,
      parent_item_id: "folder-1",
    });
    expect(apiCalls[0]).toBe("/me/drive/items/folder-1:/h.txt:/content");
  });

  it("uses drive_id when provided", async () => {
    const localPath = await writeTempFile("h.txt", "hi");
    const { apiCalls, client } = captureSmallPut({});
    await uploadFileTool.handler(client, {
      local_path: localPath,
      drive_id: "drv-1",
    });
    expect(apiCalls[0]).toBe("/drives/drv-1/root:/h.txt:/content");
  });

  it("respects custom name", async () => {
    const localPath = await writeTempFile("original.txt", "x");
    const { apiCalls, client } = captureSmallPut({});
    await uploadFileTool.handler(client, { local_path: localPath, name: "renamed.txt" });
    expect(apiCalls[0]).toBe("/me/drive/root:/renamed.txt:/content");
  });

  it("validates conflict_behavior enum", async () => {
    const localPath = await writeTempFile("x.txt", "x");
    const { client } = captureSmallPut({});
    await expect(
      uploadFileTool.handler(client, { local_path: localPath, conflict_behavior: "merge" }),
    ).rejects.toThrow("must be one of");
  });

  it("threads conflict_behavior into the Graph query param", async () => {
    const localPath = await writeTempFile("x.txt", "x");
    const { queries, client } = captureSmallPut({});
    await uploadFileTool.handler(client, {
      local_path: localPath,
      conflict_behavior: "rename",
    });
    expect(queries[0]).toEqual({ "@microsoft.graph.conflictBehavior": "rename" });
  });
});
