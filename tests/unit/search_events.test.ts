import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { searchEventsTool } from "../../src/tools/search_events.js";

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
