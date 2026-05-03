import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { createEventTool } from "../../src/tools/create_event.js";

function captureRequest(returnValue: unknown): {
  apiCalls: string[];
  bodies: unknown[];
  client: Client;
} {
  const apiCalls: string[] = [];
  const bodies: unknown[] = [];
  const post = vi.fn().mockImplementation((body: unknown) => {
    bodies.push(body);
    return Promise.resolve(returnValue);
  });
  const api = vi.fn().mockImplementation((path: string) => {
    apiCalls.push(path);
    return { post };
  });
  return { apiCalls, bodies, client: { api } as unknown as Client };
}

const validArgs = {
  subject: "Test",
  start: "2026-05-04T10:00:00",
  end: "2026-05-04T11:00:00",
};

const validResponse = {
  id: "evt-1",
  subject: "Test",
  start: { dateTime: "2026-05-04T10:00:00", timeZone: "UTC" },
  end: { dateTime: "2026-05-04T11:00:00", timeZone: "UTC" },
};

describe("createEventTool handler", () => {
  it("requires subject + start + end", async () => {
    const { client } = captureRequest(validResponse);
    await expect(createEventTool.handler(client, {})).rejects.toThrow(
      "'subject' must be a non-empty string",
    );
    await expect(createEventTool.handler(client, { subject: "x" })).rejects.toThrow(
      "'start' must be a non-empty string",
    );
    await expect(
      createEventTool.handler(client, { subject: "x", start: validArgs.start }),
    ).rejects.toThrow("'end' must be a non-empty string");
  });

  it("rejects malformed ISO dates", async () => {
    const { client } = captureRequest(validResponse);
    await expect(
      createEventTool.handler(client, { ...validArgs, start: "tomorrow" }),
    ).rejects.toThrow("ISO 8601");
  });

  it("posts to /me/events with no calendar_id", async () => {
    const { apiCalls, client } = captureRequest(validResponse);
    await createEventTool.handler(client, validArgs);
    expect(apiCalls[0]).toBe("/me/events");
  });

  it("posts to /me/calendars/{id}/events when calendar_id set", async () => {
    const { apiCalls, client } = captureRequest(validResponse);
    await createEventTool.handler(client, { ...validArgs, calendar_id: "cal-1" });
    expect(apiCalls[0]).toBe("/me/calendars/cal-1/events");
  });

  it("builds the Graph body with start/end + UTC timezone default", async () => {
    const { bodies, client } = captureRequest(validResponse);
    await createEventTool.handler(client, validArgs);
    const body = bodies[0] as Record<string, unknown>;
    expect(body.subject).toBe("Test");
    expect(body.start).toEqual({ dateTime: "2026-05-04T10:00:00", timeZone: "UTC" });
    expect(body.end).toEqual({ dateTime: "2026-05-04T11:00:00", timeZone: "UTC" });
    expect(body.isAllDay).toBe(false);
  });

  it("uses the provided timezone", async () => {
    const { bodies, client } = captureRequest(validResponse);
    await createEventTool.handler(client, { ...validArgs, timezone: "Europe/Rome" });
    const body = bodies[0] as { start: { timeZone: string } };
    expect(body.start.timeZone).toBe("Europe/Rome");
  });

  it("includes body when provided, defaulting contentType to text", async () => {
    const { bodies, client } = captureRequest(validResponse);
    await createEventTool.handler(client, { ...validArgs, body: "Notes here" });
    const body = bodies[0] as { body: { content: string; contentType: string } };
    expect(body.body).toEqual({ content: "Notes here", contentType: "text" });
  });

  it("validates body_content_type enum", async () => {
    const { client } = captureRequest(validResponse);
    await expect(
      createEventTool.handler(client, { ...validArgs, body: "x", body_content_type: "markdown" }),
    ).rejects.toThrow("must be one of");
  });

  it("includes location when provided", async () => {
    const { bodies, client } = captureRequest(validResponse);
    await createEventTool.handler(client, { ...validArgs, location: "Room 1" });
    const body = bodies[0] as { location: { displayName: string } };
    expect(body.location).toEqual({ displayName: "Room 1" });
  });

  it("builds attendees array with email + name + type", async () => {
    const { bodies, client } = captureRequest(validResponse);
    await createEventTool.handler(client, {
      ...validArgs,
      attendees: [
        { email: "alice@x.com", name: "Alice" },
        { email: "bob@x.com", type: "optional" },
      ],
    });
    const body = bodies[0] as { attendees: Array<Record<string, unknown>> };
    expect(body.attendees).toHaveLength(2);
    expect(body.attendees[0]).toEqual({
      emailAddress: { address: "alice@x.com", name: "Alice" },
      type: "required",
    });
    expect(body.attendees[1]).toEqual({
      emailAddress: { address: "bob@x.com" },
      type: "optional",
    });
  });

  it("rejects non-array attendees", async () => {
    const { client } = captureRequest(validResponse);
    await expect(
      createEventTool.handler(client, { ...validArgs, attendees: "alice@x.com" }),
    ).rejects.toThrow("'attendees' must be an array");
  });

  it("rejects attendee item missing email", async () => {
    const { client } = captureRequest(validResponse);
    await expect(
      createEventTool.handler(client, { ...validArgs, attendees: [{ name: "Alice" }] }),
    ).rejects.toThrow("'attendees[0].email' must be a non-empty string");
  });

  it("rejects attendee with invalid type", async () => {
    const { client } = captureRequest(validResponse);
    await expect(
      createEventTool.handler(client, {
        ...validArgs,
        attendees: [{ email: "x@x.com", type: "boss" }],
      }),
    ).rejects.toThrow("must be one of");
  });
});
