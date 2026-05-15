import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { listMeetingTranscriptsTool } from "../../src/tools/list_meeting_transcripts.js";

const JOIN_URL = "https://teams.microsoft.com/l/meetup-join/test-meeting-123";

/**
 * Build a mock Graph client that routes calls by path substring.
 * Supports chaining: .select().filter().get()
 */
function makeMultiApiClient(responses: Record<string, unknown>): Client {
  const api = vi.fn().mockImplementation((path: string) => {
    // Longest-match-first so /me/onlineMeetings/meet-1/transcripts
    // doesn't accidentally match the shorter /me/onlineMeetings key.
    const key = Object.keys(responses)
      .sort((a, b) => b.length - a.length)
      .find((k) => path.includes(k)) ?? path;
    const response = responses[key] ?? {};
    const get = vi.fn().mockResolvedValue(response);
    const chain: Record<string, unknown> = { get };
    // Make all chainable methods return `chain` itself
    chain.select = vi.fn().mockReturnValue(chain);
    chain.filter = vi.fn().mockReturnValue(chain);
    chain.query = vi.fn().mockReturnValue(chain);
    return chain;
  });
  return { api } as unknown as Client;
}

describe("listMeetingTranscriptsTool handler", () => {
  it("requires event_id", async () => {
    const client = makeMultiApiClient({});
    await expect(listMeetingTranscriptsTool.handler(client, {})).rejects.toThrow(
      "'event_id' must be a non-empty string",
    );
  });

  it("returns error when event is not an online meeting", async () => {
    const client = makeMultiApiClient({
      "/me/events/": { id: "evt1", subject: "Standup", isOnlineMeeting: false },
    });
    const result = await listMeetingTranscriptsTool.handler(client, { event_id: "evt1" });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.error).toBe("not_an_online_meeting");
  });

  it("returns error when joinWebUrl is missing", async () => {
    const client = makeMultiApiClient({
      "/me/events/": { id: "evt1", subject: "Call", isOnlineMeeting: true, onlineMeeting: null },
    });
    const result = await listMeetingTranscriptsTool.handler(client, { event_id: "evt1" });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.error).toBe("meeting_id_unavailable");
  });

  it("returns error when onlineMeeting filter returns no results", async () => {
    const client = makeMultiApiClient({
      "/me/events/": { id: "evt1", isOnlineMeeting: true, onlineMeeting: { joinWebUrl: JOIN_URL } },
      "/me/onlineMeetings": { value: [] },
    });
    const result = await listMeetingTranscriptsTool.handler(client, { event_id: "evt1" });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.error).toBe("meeting_id_unavailable");
  });

  it("returns empty list with note when no transcripts", async () => {
    const client = makeMultiApiClient({
      "/me/events/": { id: "evt1", isOnlineMeeting: true, onlineMeeting: { joinWebUrl: JOIN_URL } },
      "/me/onlineMeetings": { value: [{ id: "meet-1" }] },
      "/me/onlineMeetings/meet-1/transcripts": { value: [] },
    });
    const result = await listMeetingTranscriptsTool.handler(client, { event_id: "evt1" });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.count).toBe(0);
    expect(parsed.note).toBeDefined();
  });

  it("returns transcript list with metadata", async () => {
    const client = makeMultiApiClient({
      "/me/events/": { id: "evt1", isOnlineMeeting: true, onlineMeeting: { joinWebUrl: JOIN_URL } },
      "/me/onlineMeetings": { value: [{ id: "meet-1" }] },
      "/me/onlineMeetings/meet-1/transcripts": {
        value: [
          { id: "trans-1", meetingId: "meet-1", createdDateTime: "2026-05-15T10:00:00Z", endDateTime: "2026-05-15T11:00:00Z" },
        ],
      },
    });
    const result = await listMeetingTranscriptsTool.handler(client, { event_id: "evt1" });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.count).toBe(1);
    expect(parsed.transcripts[0].id).toBe("trans-1");
    expect(parsed.transcripts[0].meeting_id).toBe("meet-1");
    expect(parsed.transcripts[0].created_at).toBe("2026-05-15T10:00:00Z");
  });

  it("includes meeting_id in output for use by get_transcript", async () => {
    const client = makeMultiApiClient({
      "/me/events/": { id: "evt1", isOnlineMeeting: true, onlineMeeting: { joinWebUrl: JOIN_URL } },
      "/me/onlineMeetings": { value: [{ id: "meet-42" }] },
      "/me/onlineMeetings/meet-42/transcripts": { value: [{ id: "t1", createdDateTime: null, endDateTime: null }] },
    });
    const result = await listMeetingTranscriptsTool.handler(client, { event_id: "evt1" });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.meeting_id).toBe("meet-42");
  });
});
