import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { SITE_ID_RE, listDrivesTool, summarizeDrive } from "../../src/tools/list_drives.js";

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

  it("says that SharePoint libraries need list_sites + site_id", async () => {
    const primary = { id: "drv1", driveType: "business", name: "OneDrive", webUrl: "" };
    const resp = await listDrivesTool.handler(mockClient(primary, { value: [] }), {});
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.note).toMatch(/list_sites/);
  });
});

describe("listDrivesTool with site_id (SharePoint document libraries)", () => {
  const SITE_ID =
    "contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222";

  function siteClient(value: unknown): { client: Client; paths: string[] } {
    const paths: string[] = [];
    const api = vi.fn().mockImplementation((path: string) => {
      paths.push(path);
      return { get: async () => value };
    });
    return { client: { api } as unknown as Client, paths };
  }

  it("declares site_id in the input schema (scopable by the harness)", () => {
    expect(listDrivesTool.definition.inputSchema.properties).toHaveProperty("site_id");
    expect(listDrivesTool.definition.description).toMatch(/list_sites/);
  });

  it("lists the site's libraries from /sites/{site-id}/drives only", async () => {
    const lib = {
      id: "b!AbC-123_xyz",
      driveType: "documentLibrary",
      name: "invoices",
      webUrl: "https://contoso.sharepoint.com/sites/finance/invoices",
    };
    const { client, paths } = siteClient({ value: [lib] });
    const resp = await listDrivesTool.handler(client, { site_id: SITE_ID });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);

    expect(paths).toEqual([`/sites/${SITE_ID}/drives`]);
    expect(parsed).toMatchObject({ site_id: SITE_ID, count: 1 });
    expect(parsed.drives[0]).toMatchObject({ id: "b!AbC-123_xyz", driveType: "documentLibrary", name: "invoices" });
  });

  it("handles a site with no libraries", async () => {
    const { client } = siteClient({});
    const resp = await listDrivesTool.handler(client, { site_id: SITE_ID });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed).toMatchObject({ count: 0, drives: [] });
  });

  it.each([
    "contoso.sharepoint.com",
    "contoso.sharepoint.com:/sites/finance",
    "../me/drive",
    `${SITE_ID}/../../me`,
    `${SITE_ID}?x=1`,
    "contoso.sharepoint.com,not-a-guid,22222222-2222-2222-2222-222222222222",
  ])("rejects malformed site_id %j before any Graph call", async (bad) => {
    const { client, paths } = siteClient({ value: [] });
    await expect(listDrivesTool.handler(client, { site_id: bad })).rejects.toThrow(/site_id/);
    expect(paths).toEqual([]);
  });

  it("SITE_ID_RE accepts the shape Graph returns", () => {
    expect(SITE_ID_RE.test(SITE_ID)).toBe(true);
  });
});
