/**
 * upload_file — resumable upload-session path (> 4 MB). The Graph SDK's
 * OneDriveLargeFileUploadTask is replaced with a recording fake so the
 * test checks the session URL, options and result mapping without I/O.
 */
import { mkdtemp, open, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import type { Client } from "@microsoft/microsoft-graph-client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const { createUploadSession, uploadMock, taskCtor } = vi.hoisted(() => ({
  createUploadSession: vi.fn(),
  uploadMock: vi.fn(),
  taskCtor: vi.fn(),
}));

vi.mock("@microsoft/microsoft-graph-client", async (importOriginal) => {
  const actual = await importOriginal<typeof import("@microsoft/microsoft-graph-client")>();
  class FakeLargeFileUploadTask {
    static createUploadSession = createUploadSession;
    constructor(...args: unknown[]) {
      taskCtor(...args);
    }
    upload(): unknown {
      return uploadMock();
    }
  }
  return { ...actual, OneDriveLargeFileUploadTask: FakeLargeFileUploadTask };
});

import { uploadFileTool } from "../../src/tools/upload_file.js";

const SIZE = 5 * 1024 * 1024; // > 4 MB single-PUT threshold
let tmp: string;
let bigFile: string;

beforeEach(async () => {
  tmp = await mkdtemp(path.join(tmpdir(), "upload-large-"));
  bigFile = path.join(tmp, "big.bin");
  const fh = await open(bigFile, "w");
  await fh.truncate(SIZE); // sparse file, no real 5 MB write
  await fh.close();
  createUploadSession.mockReset().mockResolvedValue({ url: "https://upload.example/s", expiry: new Date() });
  uploadMock.mockReset().mockResolvedValue({
    responseBody: { id: "new-id", name: "big.bin", size: SIZE, webUrl: "https://example/new" },
  });
  taskCtor.mockReset();
  vi.spyOn(console, "error").mockImplementation(() => undefined);
});

afterEach(async () => {
  vi.restoreAllMocks();
  await rm(tmp, { recursive: true, force: true });
});

const client = {} as Client;

describe("uploadFileTool handler — resumable session path", () => {
  it("creates a session at the drive root and maps the uploaded item", async () => {
    const res = await uploadFileTool.handler(client, { local_path: bigFile });
    const out = JSON.parse((res.content[0] as { text: string }).text);

    expect(createUploadSession).toHaveBeenCalledWith(
      client,
      "/me/drive/root:/big.bin:/createUploadSession",
      { fileName: "big.bin", conflictBehavior: "fail" },
    );
    const [, fileObject, session, options] = taskCtor.mock.calls[0];
    expect((fileObject as { size: number }).size).toBe(SIZE);
    expect(session).toEqual(expect.objectContaining({ url: "https://upload.example/s" }));
    expect(options).toEqual({ rangeSize: 10 * 1024 * 1024 });
    expect(out).toEqual({
      drive_id: null,
      parent_item_id: null,
      uploaded: {
        id: "new-id",
        name: "big.bin",
        size: SIZE,
        webUrl: "https://example/new",
        upload_path: "resumable_session",
      },
    });
  });

  it("targets drive + parent item, custom name and conflict behaviour", async () => {
    await uploadFileTool.handler(client, {
      local_path: bigFile,
      drive_id: "b!lib",
      parent_item_id: "P1",
      name: "renamed file.bin",
      conflict_behavior: "rename",
    });
    expect(createUploadSession).toHaveBeenCalledWith(
      client,
      "/drives/b!lib/items/P1:/renamed%20file.bin:/createUploadSession",
      { fileName: "renamed file.bin", conflictBehavior: "rename" },
    );
  });

  it("returns empty fields when the task result has no response body", async () => {
    uploadMock.mockResolvedValue({});
    const res = await uploadFileTool.handler(client, { local_path: bigFile });
    const out = JSON.parse((res.content[0] as { text: string }).text);
    expect(out.uploaded).toEqual({
      id: "",
      name: "",
      size: 0,
      webUrl: "",
      upload_path: "resumable_session",
    });
  });
});
