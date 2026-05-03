import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { expandEvent, getEventTool } from "../../src/tools/get_event.js";

describe("expandEvent", () => {
  it("includes the body and body_content_type", () => {
    const event = {
      id: "x",
      subject: "test",
      body: { content: "hello world", contentType: "html" },
    };
    const r = expandEvent(event);
    expect(r.body).toBe("hello world");
    expect(r.body_content_type).toBe("html");
    expect(r.body_truncated).toBe(false);
  });

  it("truncates body at the 8000-char cap and flags it", () => {
    const longBody = "x".repeat(20_000);
    const r = expandEvent({
      id: "x",
      subject: "x",
      body: { content: longBody, contentType: "text" },
    });
    expect(r.body.length).toBe(8000);
    expect(r.body_truncated).toBe(true);
  });

  it("preserves recurrence rule when present", () => {
    const recurrence = { pattern: { type: "weekly", interval: 1 } };
    const r = expandEvent({ id: "x", subject: "x", recurrence });
    expect(r.recurrence).toBe(recurrence);
  });

  it("returns null recurrence when absent", () => {
    expect(expandEvent({ id: "x", subject: "x" }).recurrence).toBeNull();
  });
});

describe("getEventTool handler", () => {
  it("requires event_id", async () => {
    const get = vi.fn().mockResolvedValue({ id: "x" });
    const api = vi.fn().mockReturnValue({ get });
    const client = { api } as unknown as Client;
    await expect(getEventTool.handler(client, {})).rejects.toThrow(
      "'event_id' must be a non-empty string",
    );
  });

  it("URL-encodes the event_id", async () => {
    const calls: string[] = [];
    const get = vi.fn().mockResolvedValue({ id: "x", subject: "x" });
    const api = vi.fn().mockImplementation((path: string) => {
      calls.push(path);
      return { get };
    });
    const client = { api } as unknown as Client;

    await getEventTool.handler(client, { event_id: "evt/with/slash" });
    expect(calls[0]).toBe("/me/events/evt%2Fwith%2Fslash");
  });
});
