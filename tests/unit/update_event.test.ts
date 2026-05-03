import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { buildPatchBody, updateEventTool } from "../../src/tools/update_event.js";

describe("buildPatchBody", () => {
  it("returns empty object when no fields given", () => {
    expect(buildPatchBody({})).toEqual({});
  });

  it("includes subject when provided", () => {
    expect(buildPatchBody({ subject: "New" })).toEqual({ subject: "New" });
  });

  it("requires timezone when start is updated", () => {
    expect(() => buildPatchBody({ start: "2026-05-04T10:00:00" })).toThrow("'timezone' is required");
  });

  it("requires timezone when end is updated", () => {
    expect(() => buildPatchBody({ end: "2026-05-04T10:00:00" })).toThrow();
  });

  it("builds start with timezone", () => {
    const r = buildPatchBody({ start: "2026-05-04T10:00:00", timezone: "Europe/Rome" });
    expect(r.start).toEqual({ dateTime: "2026-05-04T10:00:00", timeZone: "Europe/Rome" });
  });

  it("builds body with default contentType=text", () => {
    const r = buildPatchBody({ body: "Notes" });
    expect(r.body).toEqual({ content: "Notes", contentType: "text" });
  });

  it("validates body_content_type enum", () => {
    expect(() => buildPatchBody({ body: "x", body_content_type: "markdown" })).toThrow("must be one of");
  });

  it("builds location object", () => {
    expect(buildPatchBody({ location: "Room 1" }).location).toEqual({ displayName: "Room 1" });
  });

  it("builds attendees array", () => {
    const r = buildPatchBody({
      attendees: [{ email: "alice@x.com", name: "Alice", type: "required" }],
    });
    expect(r.attendees).toEqual([
      {
        emailAddress: { address: "alice@x.com", name: "Alice" },
        type: "required",
      },
    ]);
  });

  it("rejects non-array attendees", () => {
    expect(() => buildPatchBody({ attendees: "x" })).toThrow("must be an array");
  });

  it("validates is_all_day type", () => {
    expect(() => buildPatchBody({ is_all_day: "yes" })).toThrow("must be boolean");
    expect(buildPatchBody({ is_all_day: true }).isAllDay).toBe(true);
  });
});

describe("updateEventTool handler", () => {
  it("requires event_id", async () => {
    const api = vi.fn();
    const client = { api } as unknown as Client;
    await expect(updateEventTool.handler(client, {})).rejects.toThrow(
      "'event_id' must be a non-empty string",
    );
  });

  it("rejects empty patch (only event_id given)", async () => {
    const api = vi.fn();
    const client = { api } as unknown as Client;
    await expect(updateEventTool.handler(client, { event_id: "evt-1" })).rejects.toThrow(
      "at least one field to update",
    );
  });

  it("PATCHes /me/events/{id} with the body", async () => {
    const calls: string[] = [];
    const bodies: unknown[] = [];
    const patch = vi.fn().mockImplementation((b: unknown) => {
      bodies.push(b);
      return Promise.resolve({
        id: "evt-1",
        subject: "Updated",
        start: {},
        end: {},
      });
    });
    const api = vi.fn().mockImplementation((path: string) => {
      calls.push(path);
      return { patch };
    });
    const client = { api } as unknown as Client;

    await updateEventTool.handler(client, { event_id: "evt-1", subject: "Updated" });
    expect(calls[0]).toBe("/me/events/evt-1");
    expect(bodies[0]).toEqual({ subject: "Updated" });
  });

  it("URL-encodes the event_id", async () => {
    const calls: string[] = [];
    const patch = vi.fn().mockResolvedValue({ id: "x", subject: "x", start: {}, end: {} });
    const api = vi.fn().mockImplementation((path: string) => {
      calls.push(path);
      return { patch };
    });
    const client = { api } as unknown as Client;

    await updateEventTool.handler(client, { event_id: "evt/with/slash", subject: "x" });
    expect(calls[0]).toBe("/me/events/evt%2Fwith%2Fslash");
  });
});
