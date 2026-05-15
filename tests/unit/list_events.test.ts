import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { listEventsTool, summarizeEvent } from "../../src/tools/list_events.js";

describe("summarizeEvent", () => {
  it("extracts the canonical event fields", () => {
    const event = {
      id: "evt-1",
      subject: "Standup",
      bodyPreview: "Daily sync",
      start: { dateTime: "2026-05-04T09:00:00", timeZone: "Europe/Rome" },
      end: { dateTime: "2026-05-04T09:15:00", timeZone: "Europe/Rome" },
      isAllDay: false,
      location: { displayName: "Room 101" },
      organizer: { emailAddress: { name: "Alice", address: "alice@x.com" } },
      attendees: [
        {
          emailAddress: { name: "Bob", address: "bob@x.com" },
          type: "required",
          status: { response: "accepted" },
        },
      ],
      webLink: "https://outlook/...",
    };
    const s = summarizeEvent(event);
    expect(s.subject).toBe("Standup");
    expect(s.start.timezone).toBe("Europe/Rome");
    expect(s.location).toBe("Room 101");
    expect(s.organizer_email).toBe("alice@x.com");
    expect(s.attendees).toHaveLength(1);
    expect(s.attendees[0]).toEqual({
      name: "Bob",
      email: "bob@x.com",
      type: "required",
      response: "accepted",
    });
  });

  it("returns null location when location.displayName is missing", () => {
    expect(summarizeEvent({ id: "x", subject: "x" }).location).toBeNull();
  });

  it("handles missing attendees as empty array", () => {
    expect(summarizeEvent({ id: "x", subject: "x" }).attendees).toEqual([]);
  });

  it("handles missing organizer gracefully", () => {
    const s = summarizeEvent({ id: "x", subject: "x" });
    expect(s.organizer_name).toBeNull();
    expect(s.organizer_email).toBeNull();
  });
});

function captureRequest(returnValue: unknown): {
  apiCalls: string[];
  queries: Record<string, unknown>[];
  client: Client;
} {
  const apiCalls: string[] = [];
  const queries: Record<string, unknown>[] = [];
  const get = vi.fn().mockResolvedValue(returnValue);
  const orderby = vi.fn().mockReturnValue({ get });
  const top = vi.fn().mockReturnValue({ orderby });
  const select = vi.fn().mockReturnValue({ top, orderby });
  const query = vi.fn().mockImplementation((q: Record<string, unknown>) => {
    queries.push(q);
    return { top, select };
  });
  const api = vi.fn().mockImplementation((path: string) => {
    apiCalls.push(path);
    return { query };
  });
  return { apiCalls, queries, client: { api } as unknown as Client };
}

describe("listEventsTool handler", () => {
  const validArgs = { start: "2026-05-04", end: "2026-05-05" };

  it("requires start + end", async () => {
    const { client } = captureRequest({ value: [] });
    await expect(listEventsTool.handler(client, {})).rejects.toThrow(
      "'start' must be a non-empty string",
    );
    await expect(listEventsTool.handler(client, { start: validArgs.start })).rejects.toThrow(
      "'end' must be a non-empty string",
    );
  });

  it("rejects malformed ISO date", async () => {
    const { client } = captureRequest({ value: [] });
    await expect(
      listEventsTool.handler(client, { start: "yesterday", end: validArgs.end }),
    ).rejects.toThrow("ISO 8601");
  });

  it("uses /me/calendarView when no calendar_id", async () => {
    const { apiCalls, queries, client } = captureRequest({ value: [] });
    await listEventsTool.handler(client, validArgs);
    expect(apiCalls).toEqual(["/me/calendarView"]);
    expect(queries[0]).toEqual({
      startDateTime: "2026-05-04",
      endDateTime: "2026-05-05",
    });
  });

  it("scopes to /me/calendars/{id}/calendarView when calendar_id given", async () => {
    const { apiCalls, client } = captureRequest({ value: [] });
    await listEventsTool.handler(client, { ...validArgs, calendar_id: "cal-with space" });
    expect(apiCalls[0]).toBe("/me/calendars/cal-with%20space/calendarView");
  });

  it("returns count + events + window in response", async () => {
    const { client } = captureRequest({
      value: [
        { id: "e1", subject: "Standup", start: {}, end: {} },
        { id: "e2", subject: "Lunch", start: {}, end: {} },
      ],
    });
    const resp = await listEventsTool.handler(client, validArgs);
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.count).toBe(2);
    expect(parsed.window).toEqual({ start: "2026-05-04", end: "2026-05-05" });
  });
});
