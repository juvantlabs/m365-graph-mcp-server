import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { listSitesTool, matchesQuery, summarizeSite } from "../../src/tools/list_sites.js";

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

/** Client whose successive .get() calls return the given pages in order. */
function sequenceClient(pages: unknown[]): { client: Client; api: ReturnType<typeof vi.fn>; query: ReturnType<typeof vi.fn> } {
  const queue = [...pages];
  const chain = {
    query: vi.fn(),
    select: vi.fn(),
    top: vi.fn(),
    get: vi.fn(async () => queue.shift() ?? { value: [] }),
  };
  chain.query.mockReturnValue(chain);
  chain.select.mockReturnValue(chain);
  chain.top.mockReturnValue(chain);
  const api = vi.fn().mockReturnValue(chain);
  return { client: { api } as unknown as Client, api, query: chain.query };
}

const HYPHEN_SITE = {
  id: "contoso.sharepoint.com,33333333-3333-3333-3333-333333333333,44444444-4444-4444-4444-444444444444",
  name: "team-finance",
  displayName: "Team Finance",
  webUrl: "https://contoso.sharepoint.com/sites/team-finance",
};
const OTHER_SITE = { ...SITE, name: "legal", displayName: "Legal", webUrl: "https://contoso.sharepoint.com/sites/legal" };

describe("matchesQuery", () => {
  const s = summarizeSite(HYPHEN_SITE);
  it.each(["team-finance", "TEAM-FINANCE", "Team Fin", "sites/team-"])("matches %j", (q) => {
    expect(matchesQuery(s, q)).toBe(true);
  });
  it("does not match an unrelated query", () => {
    expect(matchesQuery(s, "legal")).toBe(false);
  });
});

describe("listSitesTool name-filter fallback", () => {
  it("falls back to '*' + substring filter when the keyword search finds nothing (hyphenated name)", async () => {
    const m = sequenceClient([{ value: [] }, { value: [OTHER_SITE, HYPHEN_SITE] }]);
    const out = parse(await listSitesTool.handler(m.client, { query: "team-finance" }));
    expect(m.query.mock.calls).toEqual([[{ search: "team-finance" }], [{ search: "*" }]]);
    expect(out).toMatchObject({ query: "team-finance", match: "name_filter", count: 1 });
    expect((out.sites as Array<{ name: string }>)[0].name).toBe("team-finance");
  });

  it("does not fall back when the keyword search returns results", async () => {
    const m = sequenceClient([{ value: [SITE] }]);
    const out = parse(await listSitesTool.handler(m.client, { query: "finance" }));
    expect(m.query).toHaveBeenCalledTimes(1);
    expect(out).toMatchObject({ match: "search", count: 1 });
  });

  it("does not fall back for '*'", async () => {
    const m = sequenceClient([{ value: [] }]);
    const out = parse(await listSitesTool.handler(m.client, {}));
    expect(m.query).toHaveBeenCalledTimes(1);
    expect(out).toMatchObject({ match: "search", count: 0 });
  });

  it("returns an empty name_filter result when nothing matches", async () => {
    const m = sequenceClient([{ value: [] }, { value: [OTHER_SITE] }]);
    const out = parse(await listSitesTool.handler(m.client, { query: "nope" }));
    expect(out).toMatchObject({ match: "name_filter", count: 0, sites: [] });
  });

  it("applies limit to the filtered result", async () => {
    const many = Array.from({ length: 5 }, (_, i) => ({ ...HYPHEN_SITE, id: `${HYPHEN_SITE.id}${i}` }));
    const m = sequenceClient([{ value: [] }, { value: many }]);
    const out = parse(await listSitesTool.handler(m.client, { query: "team-finance", limit: 2 }));
    expect(out.count).toBe(2);
  });

  it("follows Graph @odata.nextLink pages during the fallback", async () => {
    const next = "https://graph.microsoft.com/v1.0/sites?search=*&$skiptoken=abc";
    const m = sequenceClient([
      { value: [] },
      { value: [OTHER_SITE], "@odata.nextLink": next },
      { value: [HYPHEN_SITE] },
    ]);
    const out = parse(await listSitesTool.handler(m.client, { query: "team-finance" }));
    expect(m.api).toHaveBeenLastCalledWith(next);
    expect(out.count).toBe(1);
  });

  it("never follows a nextLink that is not https://graph.microsoft.com/", async () => {
    const m = sequenceClient([
      { value: [] },
      { value: [OTHER_SITE], "@odata.nextLink": "https://evil.example/sites?x" },
    ]);
    const out = parse(await listSitesTool.handler(m.client, { query: "team-finance" }));
    expect(m.api).not.toHaveBeenCalledWith("https://evil.example/sites?x");
    expect(out.count).toBe(0);
  });

  it("stops after 5 fallback pages", async () => {
    const next = "https://graph.microsoft.com/v1.0/sites?search=*&$skiptoken=x";
    const page = { value: [OTHER_SITE], "@odata.nextLink": next };
    const m = sequenceClient([{ value: [] }, page, page, page, page, page, page, page]);
    await listSitesTool.handler(m.client, { query: "team-finance" });
    // 1 keyword search + 1 first fallback page + 4 nextLink pages
    expect(m.api).toHaveBeenCalledTimes(6);
  });
});
