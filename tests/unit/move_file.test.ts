import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { moveFileTool } from "../../src/tools/move_file.js";

function captureRequest(returnValue: unknown): {
  apiCalls: string[];
  bodies: unknown[];
  client: Client;
} {
  const apiCalls: string[] = [];
  const bodies: unknown[] = [];
  const patch = vi.fn().mockImplementation((b: unknown) => {
    bodies.push(b);
    return Promise.resolve(returnValue);
  });
  const api = vi.fn().mockImplementation((path: string) => {
    apiCalls.push(path);
    return { patch };
  });
  return { apiCalls, bodies, client: { api } as unknown as Client };
}

describe("moveFileTool handler", () => {
  it("requires item_id + target_parent_id", async () => {
    const { client } = captureRequest({ id: "x", name: "x" });
    await expect(moveFileTool.handler(client, {})).rejects.toThrow(
      "'item_id' must be a non-empty string",
    );
    await expect(moveFileTool.handler(client, { item_id: "x" })).rejects.toThrow(
      "'target_parent_id' must be a non-empty string",
    );
  });

  it("PATCHes /me/drive/items/{id} with parentReference body", async () => {
    const { apiCalls, bodies, client } = captureRequest({ id: "x", name: "report.pdf" });
    await moveFileTool.handler(client, { item_id: "abc", target_parent_id: "folder-1" });
    expect(apiCalls[0]).toBe("/me/drive/items/abc");
    expect(bodies[0]).toEqual({ parentReference: { id: "folder-1" } });
  });

  it("includes new_name when provided", async () => {
    const { bodies, client } = captureRequest({ id: "x", name: "renamed.pdf" });
    await moveFileTool.handler(client, {
      item_id: "abc",
      target_parent_id: "folder-1",
      new_name: "renamed.pdf",
    });
    expect(bodies[0]).toEqual({ parentReference: { id: "folder-1" }, name: "renamed.pdf" });
  });

  it("uses /drives/{id} path when drive_id provided", async () => {
    const { apiCalls, client } = captureRequest({ id: "x", name: "x" });
    await moveFileTool.handler(client, {
      item_id: "abc",
      target_parent_id: "folder-1",
      drive_id: "drv-1",
    });
    expect(apiCalls[0]).toBe("/drives/drv-1/items/abc");
  });

  it("returns moved item summary in response", async () => {
    const { client } = captureRequest({
      id: "moved",
      name: "report.pdf",
      size: 1024,
      webUrl: "https://example/moved",
      parentReference: { id: "folder-1" },
    });
    const resp = await moveFileTool.handler(client, {
      item_id: "abc",
      target_parent_id: "folder-1",
    });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.moved.id).toBe("moved");
    expect(parsed.moved.parent_id).toBe("folder-1");
  });
});
