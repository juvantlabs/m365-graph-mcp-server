import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import {
  createFolderTool,
  summarizeCreatedFolder,
} from "../../src/tools/create_folder.js";

function captureRequest(returnValue: unknown): {
  apiCalls: string[];
  bodies: unknown[];
  client: Client;
} {
  const apiCalls: string[] = [];
  const bodies: unknown[] = [];
  const post = vi.fn().mockImplementation((b: unknown) => {
    bodies.push(b);
    return Promise.resolve(returnValue);
  });
  const api = vi.fn().mockImplementation((path: string) => {
    apiCalls.push(path);
    return { post };
  });
  return { apiCalls, bodies, client: { api } as unknown as Client };
}

describe("summarizeCreatedFolder", () => {
  it("extracts canonical fields including child_count", () => {
    const item = {
      id: "fld-1",
      name: "Aruba",
      webUrl: "https://example/fld-1",
      parentReference: { id: "root-abc" },
      folder: { childCount: 3 },
    };
    expect(summarizeCreatedFolder(item)).toEqual({
      id: "fld-1",
      name: "Aruba",
      webUrl: "https://example/fld-1",
      parent_id: "root-abc",
      child_count: 3,
    });
  });

  it("defaults child_count to 0 when folder facet omitted", () => {
    expect(
      summarizeCreatedFolder({ id: "x", name: "x", webUrl: "" }).child_count,
    ).toBe(0);
  });

  it("returns null parent_id when parentReference missing", () => {
    expect(
      summarizeCreatedFolder({ id: "x", name: "x", webUrl: "" }).parent_id,
    ).toBeNull();
  });
});

describe("createFolderTool handler", () => {
  it("requires name", async () => {
    const { client } = captureRequest({ id: "x", name: "x", folder: {} });
    await expect(createFolderTool.handler(client, {})).rejects.toThrow(
      "'name' must be a non-empty string",
    );
  });

  it("rejects invalid conflict_behavior", async () => {
    const { client } = captureRequest({ id: "x", name: "x", folder: {} });
    await expect(
      createFolderTool.handler(client, { name: "Aruba", conflict_behavior: "nope" }),
    ).rejects.toThrow("'conflict_behavior' must be one of");
  });

  it("rejects conflict_behavior='replace' — deliberately not supported", async () => {
    // 'replace' would displace an existing folder (with its content) and
    // push this tool from write_idempotent toward write_irreversible
    // semantics. Excluded by design; compose delete_file + create_folder
    // if replacement is truly needed.
    const { client } = captureRequest({ id: "x", name: "x", folder: {} });
    await expect(
      createFolderTool.handler(client, { name: "Aruba", conflict_behavior: "replace" }),
    ).rejects.toThrow("'conflict_behavior' must be one of");
  });

  it("exposes only 'fail' and 'rename' in the JSON schema enum", () => {
    const props = createFolderTool.definition.inputSchema.properties as Record<
      string,
      { enum?: unknown[] }
    >;
    expect(props.conflict_behavior?.enum).toEqual(["fail", "rename"]);
  });

  it("POSTs to /me/drive/root/children by default", async () => {
    const { apiCalls, bodies, client } = captureRequest({
      id: "fld-1",
      name: "Aruba",
      webUrl: "https://example/fld-1",
      parentReference: { id: "root" },
      folder: { childCount: 0 },
    });
    await createFolderTool.handler(client, { name: "Aruba" });
    expect(apiCalls[0]).toBe("/me/drive/root/children");
    expect(bodies[0]).toEqual({
      name: "Aruba",
      folder: {},
      "@microsoft.graph.conflictBehavior": "fail",
    });
  });

  it("uses parent_item_id path when provided", async () => {
    const { apiCalls, client } = captureRequest({
      id: "fld-2",
      name: "2026",
      webUrl: "https://example/fld-2",
      parentReference: { id: "parent-1" },
      folder: {},
    });
    await createFolderTool.handler(client, {
      name: "2026",
      parent_item_id: "parent-1",
    });
    expect(apiCalls[0]).toBe("/me/drive/items/parent-1/children");
  });

  it("uses /drives/{id} path when drive_id provided", async () => {
    const { apiCalls, client } = captureRequest({
      id: "fld-3",
      name: "Namecheap",
      webUrl: "https://example/fld-3",
      parentReference: { id: "root" },
      folder: {},
    });
    await createFolderTool.handler(client, {
      name: "Namecheap",
      drive_id: "drv-fin",
    });
    expect(apiCalls[0]).toBe("/drives/drv-fin/root/children");
  });

  it("combines drive_id + parent_item_id correctly", async () => {
    const { apiCalls, client } = captureRequest({
      id: "fld-4",
      name: "_SDI-ufficiali",
      webUrl: "https://example/fld-4",
      parentReference: { id: "parent-x" },
      folder: {},
    });
    await createFolderTool.handler(client, {
      name: "_SDI-ufficiali",
      drive_id: "drv-fin",
      parent_item_id: "parent-x",
    });
    expect(apiCalls[0]).toBe("/drives/drv-fin/items/parent-x/children");
  });

  it("passes through leading-underscore names unmodified", async () => {
    const { bodies, client } = captureRequest({
      id: "fld-5",
      name: "_Personal-pre_startup",
      webUrl: "https://example/fld-5",
      parentReference: { id: "root" },
      folder: {},
    });
    await createFolderTool.handler(client, { name: "_Personal-pre_startup" });
    expect((bodies[0] as { name: string }).name).toBe("_Personal-pre_startup");
  });

  it("respects conflict_behavior=rename", async () => {
    const { bodies, client } = captureRequest({
      id: "fld-6",
      name: "Aruba 1",
      webUrl: "https://example/fld-6",
      parentReference: { id: "root" },
      folder: {},
    });
    await createFolderTool.handler(client, {
      name: "Aruba",
      conflict_behavior: "rename",
    });
    expect((bodies[0] as Record<string, unknown>)["@microsoft.graph.conflictBehavior"]).toBe(
      "rename",
    );
  });

  it("URL-encodes drive_id and parent_item_id", async () => {
    const { apiCalls, client } = captureRequest({
      id: "x",
      name: "x",
      webUrl: "",
      parentReference: { id: "p" },
      folder: {},
    });
    await createFolderTool.handler(client, {
      name: "New",
      drive_id: "drv/with slash",
      parent_item_id: "parent id",
    });
    expect(apiCalls[0]).toBe(
      "/drives/drv%2Fwith%20slash/items/parent%20id/children",
    );
  });

  it("returns the created folder summary in the response", async () => {
    const { client } = captureRequest({
      id: "fld-created",
      name: "Aruba",
      webUrl: "https://example/fld-created",
      parentReference: { id: "root" },
      folder: { childCount: 0 },
    });
    const resp = await createFolderTool.handler(client, {
      name: "Aruba",
      drive_id: "drv-fin",
    });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.drive_id).toBe("drv-fin");
    expect(parsed.parent_item_id).toBeNull();
    expect(parsed.created.id).toBe("fld-created");
    expect(parsed.created.name).toBe("Aruba");
    expect(parsed.created.parent_id).toBe("root");
    expect(parsed.created.child_count).toBe(0);
  });

  it("is categorized write_idempotent (reversible via delete_file)", () => {
    expect(createFolderTool.category).toBe("write_idempotent");
  });
});
