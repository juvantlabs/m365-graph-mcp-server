import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { searchEventsTool } from "../../src/tools/search_events.js";
import { wireClient } from "../fixtures/graph_wire.js";

function captureFilter(returnValue: unknown): { filters: string[]; client: Client } {
  const filters: string[] = [];
  const get = vi.fn().mockResolvedValue(returnValue);
  const top = vi.fn().mockReturnValue({ get });
  const filter = vi.fn().mockImplementation((f: string) => {
    filters.push(f);
    return { top };
  });
  const api = vi.fn().mockReturnValue({ filter });
  return { filters, client: { api } as unknown as Client };
}

describe("searchEventsTool handler", () => {
  it("requires a query", async () => {
    const { client } = captureFilter({ value: [] });
    await expect(searchEventsTool.handler(client, {})).rejects.toThrow(
      "'query' must be a non-empty string",
    );
  });

  it("uses contains(subject, '…') as the filter", async () => {
    const { filters, client } = captureFilter({ value: [] });
    await searchEventsTool.handler(client, { query: "review" });
    expect(filters[0]).toBe("contains(subject, 'review')");
  });

  it("escapes single quotes via ' → '' (OData literal-string)", async () => {
    const { filters, client } = captureFilter({ value: [] });
    await searchEventsTool.handler(client, { query: "O'Brien" });
    expect(filters[0]).toBe("contains(subject, 'O''Brien')");
  });

  it("validates limit upper bound", async () => {
    const { client } = captureFilter({ value: [] });
    await expect(searchEventsTool.handler(client, { query: "x", limit: 999 })).rejects.toThrow();
  });

  it("returns count + results in response shape", async () => {
    const { client } = captureFilter({
      value: [
        { id: "e1", subject: "Project review", start: {}, end: {} },
        { id: "e2", subject: "Quarterly review", start: {}, end: {} },
      ],
    });
    const resp = await searchEventsTool.handler(client, { query: "review" });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.query).toBe("review");
    expect(parsed.count).toBe(2);
  });
});

describe("searchEventsTool query-string encoding (real SDK serialization)", () => {
  it.each([
    ["plain", "weekly review", "weekly review"],
    ["plus", "C++ review", "C++ review"],
    ["ampersand", "R&D sync", "R&D sync"],
    ["hash", "Q3 #1", "Q3 #1"],
    ["percent", "50% done", "50% done"],
    ["single quote", "O'Brien", "O''Brien"],
    ["non-ASCII", "caffè", "caffè"],
    ["datetime with offset", "2026-11-04T00:00:00+01:00", "2026-11-04T00:00:00+01:00"],
  ])("delivers a %s subject term to Graph unchanged", async (_label, query, literal) => {
    const { client, sent } = wireClient();
    await searchEventsTool.handler(client, { query });
    expect(sent[0].searchParams.get("$filter")).toBe(`contains(subject, '${literal}')`);
  });

  it("cannot add or truncate query parameters through the search term", async () => {
    const { client, sent } = wireClient();
    await searchEventsTool.handler(client, { query: "x')&$top=999&y=#z" });
    expect([...sent[0].searchParams.keys()]).toEqual(["$filter", "$top"]);
    expect(sent[0].searchParams.get("$top")).toBe("20");
    expect(sent[0].hash).toBe("");
  });
});
