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
