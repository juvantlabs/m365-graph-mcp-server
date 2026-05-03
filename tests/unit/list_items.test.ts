import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { listItemsTool, summarizeDriveItem } from "../../src/tools/list_items.js";

describe("summarizeDriveItem", () => {
  it("classifies an item as folder when it has a folder facet", () => {
    const item = {
      id: "i1",
      name: "Documents",
      folder: { childCount: 5 },
      size: 0,
      lastModifiedDateTime: "2026-05-04T10:00:00Z",
      webUrl: "https://example.com/i1",
    };
    const s = summarizeDriveItem(item);
    expect(s.type).toBe("folder");
    expect(s.child_count).toBe(5);
  });

  it("classifies an item as file when no folder facet", () => {
    const item = {
      id: "i2",
      name: "report.pdf",
      size: 1024,
      lastModifiedDateTime: "2026-05-04T10:00:00Z",
      webUrl: "https://example.com/i2",
    };
    const s = summarizeDriveItem(item);
    expect(s.type).toBe("file");
    expect(s.child_count).toBeNull();
    expect(s.size).toBe(1024);
  });

  it("treats folder.childCount=undefined as 0", () => {
    const item = { id: "i3", name: "Empty", folder: {} };
    expect(summarizeDriveItem(item).child_count).toBe(0);
  });
});

function mockClient(response: unknown): Client {
  const top = vi.fn().mockReturnValue({ get: vi.fn().mockResolvedValue(response) });
  const api = vi.fn().mockReturnValue({ top });
  return { api } as unknown as Client;
}

describe("listItemsTool handler", () => {
  it("calls /me/drive/root/children when no item_id given", async () => {
    const apiCalls: string[] = [];
    const top = vi.fn().mockReturnValue({ get: vi.fn().mockResolvedValue({ value: [] }) });
    const api = vi.fn().mockImplementation((path: string) => {
      apiCalls.push(path);
      return { top };
    });
    const client = { api } as unknown as Client;

    await listItemsTool.handler(client, {});
    expect(apiCalls).toEqual(["/me/drive/root/children"]);
  });

  it("calls /me/drive/items/<id>/children when item_id given", async () => {
    const apiCalls: string[] = [];
    const top = vi.fn().mockReturnValue({ get: vi.fn().mockResolvedValue({ value: [] }) });
    const api = vi.fn().mockImplementation((path: string) => {
      apiCalls.push(path);
      return { top };
    });
    const client = { api } as unknown as Client;

    await listItemsTool.handler(client, { item_id: "abc-123" });
    expect(apiCalls[0]).toBe("/me/drive/items/abc-123/children");
  });

  it("URL-encodes drive_id and item_id (defense vs OData injection)", async () => {
    const apiCalls: string[] = [];
    const top = vi.fn().mockReturnValue({ get: vi.fn().mockResolvedValue({ value: [] }) });
    const api = vi.fn().mockImplementation((path: string) => {
      apiCalls.push(path);
      return { top };
    });
    const client = { api } as unknown as Client;

    await listItemsTool.handler(client, { drive_id: "with space", item_id: "id/with/slash" });
    expect(apiCalls[0]).toBe("/drives/with%20space/items/id%2Fwith%2Fslash/children");
  });

  it("validates limit out of range", async () => {
    const client = mockClient({ value: [] });
    await expect(listItemsTool.handler(client, { limit: 999 })).rejects.toThrow(
      "'limit' must be an integer between 1 and 100",
    );
  });

  it("returns count + items in response shape", async () => {
    const items = [
      { id: "f1", name: "report.pdf", size: 1024 },
      { id: "f2", name: "folder", folder: { childCount: 3 } },
    ];
    const client = mockClient({ value: items });

    const resp = await listItemsTool.handler(client, {});
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.count).toBe(2);
    expect(parsed.items[0].type).toBe("file");
    expect(parsed.items[1].type).toBe("folder");
    expect(parsed.items[1].child_count).toBe(3);
  });
});
