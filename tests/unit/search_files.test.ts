import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { searchFilesTool, summarizeSearchHit } from "../../src/tools/search_files.js";

describe("summarizeSearchHit", () => {
  it("joins parentReference.path and name into a virtual path", () => {
    const item = {
      id: "i1",
      name: "report.pdf",
      size: 100,
      parentReference: { path: "/drive/root:/Documents" },
      lastModifiedDateTime: "",
      webUrl: "",
    };
    const s = summarizeSearchHit(item);
    expect(s.path).toBe("/drive/root:/Documents/report.pdf");
  });

  it("handles missing parentReference (falls back to name)", () => {
    expect(summarizeSearchHit({ id: "x", name: "file.txt" }).path).toBe("file.txt");
  });

  it("classifies as folder when folder facet present", () => {
    expect(summarizeSearchHit({ id: "x", name: "Docs", folder: {} }).is_folder).toBe(true);
  });

  it("classifies as file when no folder facet", () => {
    expect(summarizeSearchHit({ id: "x", name: "f.txt" }).is_folder).toBe(false);
  });
});

function captureApiPath(returnValue: unknown): { calls: string[]; client: Client } {
  const calls: string[] = [];
  const top = vi.fn().mockReturnValue({ get: vi.fn().mockResolvedValue(returnValue) });
  const api = vi.fn().mockImplementation((path: string) => {
    calls.push(path);
    return { top };
  });
  return { calls, client: { api } as unknown as Client };
}

describe("searchFilesTool handler", () => {
  it("requires a query string", async () => {
    const { client } = captureApiPath({ value: [] });
    await expect(searchFilesTool.handler(client, {})).rejects.toThrow(
      "'query' must be a non-empty string",
    );
  });

  it("uses /me/drive when no drive_id given", async () => {
    const { calls, client } = captureApiPath({ value: [] });
    await searchFilesTool.handler(client, { query: "report" });
    expect(calls[0]).toBe("/me/drive/root/search(q='report')");
  });

  it("scopes to /drives/{id} when drive_id provided", async () => {
    const { calls, client } = captureApiPath({ value: [] });
    await searchFilesTool.handler(client, { query: "report", drive_id: "drv1" });
    expect(calls[0]).toBe("/drives/drv1/root/search(q='report')");
  });

  it("escapes single quotes in the query (' → '')", async () => {
    const { calls, client } = captureApiPath({ value: [] });
    await searchFilesTool.handler(client, { query: "O'Brien" });
    // OData literal-string escaping: ' → ''
    expect(calls[0]).toBe("/me/drive/root/search(q='O''Brien')");
  });

  it("validates limit upper bound", async () => {
    const { client } = captureApiPath({ value: [] });
    await expect(
      searchFilesTool.handler(client, { query: "x", limit: 999 }),
    ).rejects.toThrow();
  });
});

describe("searchFilesTool handler — result mapping", () => {
  it("maps every Graph hit through summarizeSearchHit and reports the count", async () => {
    const { client } = captureApiPath({
      value: [
        {
          id: "a",
          name: "report.docx",
          size: 10,
          parentReference: { path: "/drive/root:/Legal" },
          lastModifiedDateTime: "2026-01-01T00:00:00Z",
          webUrl: "https://example/a",
        },
        { id: "b", name: "Archive", folder: { childCount: 2 } },
      ],
    });
    const res = await searchFilesTool.handler(client, { query: "report" });
    const out = JSON.parse((res.content[0] as { text: string }).text);
    expect(out.count).toBe(2);
    expect(out.results[0]).toMatchObject({
      id: "a",
      path: "/drive/root:/Legal/report.docx",
      size: 10,
      is_folder: false,
    });
    expect(out.results[1]).toMatchObject({ id: "b", path: "Archive", is_folder: true });
  });

  it("treats a missing value array as zero results", async () => {
    const { client } = captureApiPath({});
    const res = await searchFilesTool.handler(client, { query: "x" });
    expect(JSON.parse((res.content[0] as { text: string }).text).count).toBe(0);
  });
});
