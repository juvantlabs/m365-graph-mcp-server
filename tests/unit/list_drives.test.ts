import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { listDrivesTool, summarizeDrive } from "../../src/tools/list_drives.js";

describe("summarizeDrive", () => {
  it("extracts core fields from a Graph drive object", () => {
    const drive = {
      id: "drv1",
      driveType: "business",
      name: "OneDrive",
      webUrl: "https://example.com/drv1",
      owner: { user: { displayName: "Alice" } },
    };
    expect(summarizeDrive(drive)).toEqual({
      id: "drv1",
      driveType: "business",
      name: "OneDrive",
      webUrl: "https://example.com/drv1",
      owner: "Alice",
    });
  });

  it("falls back to group displayName when no user owner", () => {
    const drive = {
      id: "drv2",
      driveType: "documentLibrary",
      name: "Team Drive",
      webUrl: "https://example.com/drv2",
      owner: { group: { displayName: "Engineering" } },
    };
    expect(summarizeDrive(drive).owner).toBe("Engineering");
  });

  it("returns null owner when neither user nor group is present", () => {
    const drive = {
      id: "drv3",
      driveType: "personal",
      name: "OneDrive",
      webUrl: "",
    };
    expect(summarizeDrive(drive).owner).toBeNull();
  });

  it("coerces missing fields to empty strings", () => {
    const result = summarizeDrive({});
    expect(result.id).toBe("");
    expect(result.name).toBe("");
    expect(result.webUrl).toBe("");
  });
});

describe("listDrivesTool handler", () => {
  function mockClient(primary: unknown, all: unknown): Client {
    const get = vi.fn(async (path: string) => {
      if (path === "/me/drive") return primary;
      if (path === "/me/drives") return all;
      throw new Error(`unexpected ${path}`);
    });
    const apiMock = vi.fn().mockImplementation((_path: string) => ({ get: () => get(_path) }));
    return { api: apiMock } as unknown as Client;
  }

  it("returns primary + accessible drives", async () => {
    const primary = {
      id: "drv1",
      driveType: "business",
      name: "OneDrive",
      webUrl: "https://example.com/drv1",
      owner: { user: { displayName: "Alice" } },
    };
    const all = { value: [primary] };
    const client = mockClient(primary, all);

    const resp = await listDrivesTool.handler(client, {});
    expect(resp.isError).toBeFalsy();

    const text = (resp.content[0] as { type: string; text: string }).text;
    const parsed = JSON.parse(text);
    expect(parsed.primary.id).toBe("drv1");
    expect(parsed.primary.owner).toBe("Alice");
    expect(parsed.accessible).toHaveLength(1);
  });

  it("handles empty /me/drives gracefully", async () => {
    const primary = { id: "drv1", driveType: "business", name: "OneDrive", webUrl: "" };
    const client = mockClient(primary, { value: [] });

    const resp = await listDrivesTool.handler(client, {});
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.accessible).toHaveLength(0);
  });
});
