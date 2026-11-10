import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { listSitesTool, summarizeSite } from "../../src/tools/list_sites.js";

const SITE = {
  id: "contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,22222222-2222-2222-2222-222222222222",
  name: "finance",
  displayName: "Finance",
  webUrl: "https://contoso.sharepoint.com/sites/finance",
};

function mockClient(value: unknown): {
  client: Client;
  api: ReturnType<typeof vi.fn>;
  query: ReturnType<typeof vi.fn>;
  select: ReturnType<typeof vi.fn>;
  top: ReturnType<typeof vi.fn>;
} {
  const chain = {
    query: vi.fn(),
    select: vi.fn(),
    top: vi.fn(),
    get: vi.fn(async () => value),
  };
  chain.query.mockReturnValue(chain);
  chain.select.mockReturnValue(chain);
  chain.top.mockReturnValue(chain);
  const api = vi.fn().mockReturnValue(chain);
  return { client: { api } as unknown as Client, api, query: chain.query, select: chain.select, top: chain.top };
}

function parse(resp: { content: Array<unknown> }): Record<string, unknown> {
  return JSON.parse((resp.content[0] as { text: string }).text);
}

describe("summarizeSite", () => {
  it("maps Graph site fields", () => {
    expect(summarizeSite(SITE)).toEqual({
      id: SITE.id,
      name: "finance",
      display_name: "Finance",
      web_url: SITE.webUrl,
    });
  });

  it("coerces missing fields to empty strings", () => {
    expect(summarizeSite({})).toEqual({ id: "", name: "", display_name: "", web_url: "" });
  });
});

describe("listSitesTool", () => {
  it("is a read tool named m365-graph:list_sites", () => {
    expect(listSitesTool.category).toBe("read");
    expect(listSitesTool.definition.name).toBe("m365-graph:list_sites");
    expect(listSitesTool.definition.description).toMatch(/list_drives/);
  });

  it("searches all sites by default ('*') via /sites?search=, limit 50", async () => {
    const m = mockClient({ value: [SITE] });
    const out = parse(await listSitesTool.handler(m.client, {}));
    expect(m.api).toHaveBeenCalledWith("/sites");
    expect(m.query).toHaveBeenCalledWith({ search: "*" });
    expect(m.select).toHaveBeenCalledWith("id,name,displayName,webUrl");
    expect(m.top).toHaveBeenCalledWith(50);
    expect(out).toMatchObject({ query: "*", count: 1 });
    expect((out.sites as unknown[])[0]).toMatchObject({ id: SITE.id, display_name: "Finance" });
  });

  it("passes the keyword as a query parameter, never in the path", async () => {
    const m = mockClient({ value: [] });
    await listSitesTool.handler(m.client, { query: "  fin/../ance  ", limit: 5 });
    expect(m.api).toHaveBeenCalledWith("/sites");
    expect(m.query).toHaveBeenCalledWith({ search: "fin/../ance" });
    expect(m.top).toHaveBeenCalledWith(5);
  });

  it("caps the result at limit even if Graph returns more", async () => {
    const m = mockClient({ value: [SITE, SITE, SITE] });
    const out = parse(await listSitesTool.handler(m.client, { limit: 2 }));
    expect(out.count).toBe(2);
  });

  it("handles a missing value array", async () => {
    const out = parse(await listSitesTool.handler(mockClient({}).client, {}));
    expect(out).toMatchObject({ count: 0, sites: [] });
  });

  it.each([
    [{ query: "" }, /non-empty/],
    [{ query: "   " }, /1–200/],
    [{ query: "x".repeat(201) }, /1–200/],
    [{ query: 42 }, /non-empty string/],
    [{ limit: 0 }, /between 1 and 100/],
    [{ limit: 101 }, /between 1 and 100/],
  ])("rejects invalid input %j", async (args, msg) => {
    await expect(listSitesTool.handler(mockClient({ value: [] }).client, args)).rejects.toThrow(msg);
  });
});
