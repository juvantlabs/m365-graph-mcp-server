import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { searchEventsContentTool } from "../../src/tools/search_events_content.js";

function captureRequest(returnValue: unknown): {
  apiCalls: string[];
  postBodies: unknown[];
  client: Client;
} {
  const apiCalls: string[] = [];
  const postBodies: unknown[] = [];
  const post = vi.fn().mockImplementation((b: unknown) => {
    postBodies.push(b);
    return Promise.resolve(returnValue);
  });
  const api = vi.fn().mockImplementation((path: string) => {
    apiCalls.push(path);
    return { post };
  });
  return { apiCalls, postBodies, client: { api } as unknown as Client };
}

describe("searchEventsContentTool handler", () => {
  it("requires query", async () => {
    const { client } = captureRequest({});
    await expect(searchEventsContentTool.handler(client, {})).rejects.toThrow(
      "'query' must be a non-empty string",
    );
  });

  it("POSTs to /search/query with the right entityTypes + queryString", async () => {
    const { apiCalls, postBodies, client } = captureRequest({ value: [] });
    await searchEventsContentTool.handler(client, { query: "review" });
    expect(apiCalls[0]).toBe("/search/query");
    const body = postBodies[0] as { requests: Array<{ entityTypes: string[]; query: { queryString: string } }> };
    expect(body.requests[0].entityTypes).toEqual(["event"]);
    expect(body.requests[0].query.queryString).toBe("review");
  });

  it("threads limit + from into the request", async () => {
    const { postBodies, client } = captureRequest({ value: [] });
    await searchEventsContentTool.handler(client, { query: "x", limit: 10, from: 25 });
    const body = postBodies[0] as { requests: Array<{ size: number; from: number }> };
    expect(body.requests[0].size).toBe(10);
    expect(body.requests[0].from).toBe(25);
  });

  it("validates limit + from ranges", async () => {
    const { client } = captureRequest({});
    await expect(
      searchEventsContentTool.handler(client, { query: "x", limit: 999 }),
    ).rejects.toThrow("must be an integer between 1 and 50");
    await expect(
      searchEventsContentTool.handler(client, { query: "x", from: -1 }),
    ).rejects.toThrow();
  });

  it("maps Search API hits → summarizeEvent shape", async () => {
    const apiResponse = {
      value: [
        {
          hitsContainers: [
            {
              total: 2,
              hits: [
                {
                  hitId: "hit-1",
                  resource: {
                    id: "evt-1",
                    subject: "Project review",
                    start: { dateTime: "2026-05-04T10:00:00", timeZone: "UTC" },
                    end: { dateTime: "2026-05-04T11:00:00", timeZone: "UTC" },
                  },
                },
                {
                  hitId: "hit-2",
                  resource: {
                    id: "evt-2",
                    subject: "Quarterly review",
                    start: {},
                    end: {},
                  },
                },
              ],
            },
          ],
        },
      ],
    };
    const { client } = captureRequest(apiResponse);

    const resp = await searchEventsContentTool.handler(client, { query: "review" });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.count).toBe(2);
    expect(parsed.total).toBe(2);
    expect(parsed.results[0].subject).toBe("Project review");
    expect(parsed.results[1].id).toBe("evt-2");
  });

  it("handles empty response gracefully", async () => {
    const { client } = captureRequest({ value: [] });
    const resp = await searchEventsContentTool.handler(client, { query: "x" });
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.count).toBe(0);
    expect(parsed.results).toEqual([]);
  });
});
