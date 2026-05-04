import type { Client } from "@microsoft/microsoft-graph-client";
import { afterEach, describe, expect, it, vi } from "vitest";

import {
  classifyMonitorStatus,
  copyFileTool,
  summarizeCopiedItem,
} from "../../src/tools/copy_file.js";

describe("classifyMonitorStatus", () => {
  it("maps 'completed' to completed status", () => {
    const r = classifyMonitorStatus({
      status: "completed",
      resourceLocation: "/items/abc",
      percentageComplete: 100,
    });
    expect(r.status).toBe("completed");
    expect(r.resource_location).toBe("/items/abc");
    expect(r.percent_complete).toBe(100);
  });

  it("maps 'failed' to failed status with error payload", () => {
    const r = classifyMonitorStatus({
      status: "failed",
      error: { code: "rejected", message: "permission denied" },
    });
    expect(r.status).toBe("failed");
    expect(r.error).toEqual({ code: "rejected", message: "permission denied" });
  });

  it("maps everything else to in_progress", () => {
    expect(classifyMonitorStatus({ status: "inProgress", percentageComplete: 50 }).status).toBe(
      "in_progress",
    );
    expect(classifyMonitorStatus({ status: "notStarted" }).status).toBe("in_progress");
    expect(classifyMonitorStatus({}).status).toBe("in_progress");
  });

  it("is case-insensitive on status comparison", () => {
    expect(classifyMonitorStatus({ status: "Completed" }).status).toBe("completed");
    expect(classifyMonitorStatus({ status: "FAILED" }).status).toBe("failed");
  });
});

describe("summarizeCopiedItem", () => {
  it("extracts canonical fields", () => {
    const item = {
      id: "i1",
      name: "report.pdf",
      size: 2048,
      webUrl: "https://example/i1",
      parentReference: { id: "parent-1" },
    };
    expect(summarizeCopiedItem(item)).toEqual({
      id: "i1",
      name: "report.pdf",
      size: 2048,
      webUrl: "https://example/i1",
      parent_id: "parent-1",
    });
  });

  it("returns null parent_id when no parentReference", () => {
    expect(summarizeCopiedItem({ id: "x", name: "x", size: 0, webUrl: "" }).parent_id).toBeNull();
  });
});

// ----------------------------------------------------------------------
// Handler tests with fake timers + mocked Graph client.
//
// The handler chains: GET source → POST /copy (raw response) → poll
// monitor URL → GET resourceLocation (or fallback list). We mock all
// four call sites and use fake timers so the polling sleep doesn't
// stretch the test.
// ----------------------------------------------------------------------

interface ApiSpec {
  source?: Record<string, unknown>;
  copyResponse?: { status: number; locationHeader: string | null };
  monitorPolls?: Array<Record<string, unknown>>;
  resourceLocationItem?: Record<string, unknown>;
  childrenListing?: Record<string, unknown>;
}

function makeClient(spec: ApiSpec): { client: Client; calls: string[] } {
  const calls: string[] = [];
  let pollIndex = 0;

  const api = vi.fn().mockImplementation((path: string) => {
    calls.push(path);

    // POST /items/{id}/copy
    if (path.endsWith("/copy")) {
      return {
        responseType: vi.fn().mockReturnValue({
          post: vi.fn().mockResolvedValue({
            status: spec.copyResponse?.status ?? 202,
            statusText: "Accepted",
            headers: {
              get: (k: string) =>
                k.toLowerCase() === "location" ? spec.copyResponse?.locationHeader ?? null : null,
            },
          } as unknown as Response),
        }),
      };
    }

    // monitor URL → poll responses (consume in order)
    if (path.startsWith("https://") && spec.monitorPolls) {
      const payload = spec.monitorPolls[Math.min(pollIndex, spec.monitorPolls.length - 1)];
      pollIndex++;
      return { get: vi.fn().mockResolvedValue(payload) };
    }

    // resourceLocation fetch
    if (path === "/me/drive/items/new-id" && spec.resourceLocationItem) {
      return { get: vi.fn().mockResolvedValue(spec.resourceLocationItem) };
    }

    // children listing (fallback path) — has .filter().get()
    if (path.endsWith("/children") && spec.childrenListing) {
      return {
        filter: vi.fn().mockReturnValue({
          get: vi.fn().mockResolvedValue(spec.childrenListing),
        }),
      };
    }

    // Source metadata (last fallback — any /items/{id} that isn't a /copy)
    return { get: vi.fn().mockResolvedValue(spec.source ?? { id: "src", name: "src.txt" }) };
  });

  return { client: { api } as unknown as Client, calls };
}

afterEach(() => {
  vi.useRealTimers();
});

describe("copyFileTool handler — async polling pattern", () => {
  it("happy path: 202 → poll completed → fetch resourceLocation", async () => {
    vi.useFakeTimers({ toFake: ["setTimeout"] });

    const { client } = makeClient({
      source: { id: "src", name: "report.pdf" },
      copyResponse: { status: 202, locationHeader: "https://graph.microsoft.com/monitor/abc" },
      monitorPolls: [
        { status: "completed", resourceLocation: "/me/drive/items/new-id" },
      ],
      resourceLocationItem: {
        id: "new-id",
        name: "report.pdf",
        size: 2048,
        webUrl: "https://example/new-id",
        parentReference: { id: "target-parent" },
      },
    });

    const promise = copyFileTool.handler(client, {
      item_id: "src",
      target_parent_id: "target-parent",
      wait_max_seconds: 60,
    });

    await vi.advanceTimersByTimeAsync(2_000);
    const resp = await promise;

    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.status).toBe("completed");
    expect(parsed.copied.id).toBe("new-id");
    expect(parsed.copied.parent_id).toBe("target-parent");
  });

  it("failed status throws with the error payload", async () => {
    vi.useFakeTimers({ toFake: ["setTimeout"] });

    const { client } = makeClient({
      source: { id: "src", name: "x" },
      copyResponse: { status: 202, locationHeader: "https://graph.microsoft.com/monitor/x" },
      monitorPolls: [
        { status: "failed", error: { code: "rejected", message: "permission denied" } },
      ],
    });

    // Catch upfront so the eventual rejection isn't an unhandled promise
    // between advanceTimers and the await — fake timers + async error
    // ordering is otherwise fragile.
    const settled = copyFileTool
      .handler(client, {
        item_id: "src",
        target_parent_id: "tgt",
        wait_max_seconds: 60,
      })
      .then(
        () => ({ ok: true as const }),
        (err: Error) => ({ ok: false as const, err }),
      );

    await vi.advanceTimersByTimeAsync(2_000);
    const result = await settled;
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.err.message).toMatch(/copy operation failed/);
    }
  });

  it("falls back to list-by-name when monitor omits resourceLocation", async () => {
    vi.useFakeTimers({ toFake: ["setTimeout"] });

    const { client } = makeClient({
      source: { id: "src", name: "report.pdf" },
      copyResponse: { status: 202, locationHeader: "https://graph.microsoft.com/monitor/x" },
      monitorPolls: [{ status: "completed" }], // no resourceLocation
      childrenListing: {
        value: [
          {
            id: "found-by-name",
            name: "renamed.pdf",
            size: 100,
            webUrl: "https://example/found",
            parentReference: { id: "target-parent" },
          },
        ],
      },
    });

    const promise = copyFileTool.handler(client, {
      item_id: "src",
      target_parent_id: "target-parent",
      new_name: "renamed.pdf",
      wait_max_seconds: 60,
    });

    await vi.advanceTimersByTimeAsync(2_000);
    const resp = await promise;
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.status).toBe("completed");
    expect(parsed.copied.id).toBe("found-by-name");
  });

  it("throws on non-202 from initial POST", async () => {
    const { client } = makeClient({
      source: { id: "src", name: "x" },
      copyResponse: { status: 403, locationHeader: null },
    });

    await expect(
      copyFileTool.handler(client, {
        item_id: "src",
        target_parent_id: "tgt",
      }),
    ).rejects.toThrow(/expected 202 Accepted/);
  });

  it("throws when the 202 response has no Location header", async () => {
    const { client } = makeClient({
      source: { id: "src", name: "x" },
      copyResponse: { status: 202, locationHeader: null },
    });

    await expect(
      copyFileTool.handler(client, {
        item_id: "src",
        target_parent_id: "tgt",
      }),
    ).rejects.toThrow(/missing Location header/);
  });

  it("requires item_id + target_parent_id", async () => {
    const { client } = makeClient({});
    await expect(copyFileTool.handler(client, {})).rejects.toThrow(
      "'item_id' must be a non-empty string",
    );
    await expect(copyFileTool.handler(client, { item_id: "x" })).rejects.toThrow(
      "'target_parent_id' must be a non-empty string",
    );
  });
});
