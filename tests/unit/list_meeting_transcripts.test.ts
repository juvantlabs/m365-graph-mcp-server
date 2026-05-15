import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { listMeetingTranscriptsTool } from "../../src/tools/list_meeting_transcripts.js";

function makeMultiApiClient(responses: Record<string, unknown>): Client {
  const api = vi.fn().mockImplementation((path: string) => {
    const key = Object.keys(responses).find((k) => path.includes(k)) ?? path;
    const get = vi.fn().mockResolvedValue(responses[key] ?? {});
    return { get, select: vi.fn().mockReturnThis(), query: vi.fn().mockReturnThis() };
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

  it("returns error when onlineMeetingId is missing", async () => {
    const client = makeMultiApiClient({
      "/me/events/": { id: "evt1", subject: "Call", isOnlineMeeting: true, onlineMeetingId: undefined },
    });
    const result = await listMeetingTranscriptsTool.handler(client, { event_id: "evt1" });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.error).toBe("meeting_id_unavailable");
  });

  it("returns empty list with note when no transcripts", async () => {
    const client = makeMultiApiClient({
      "/me/events/": { id: "evt1", isOnlineMeeting: true, onlineMeetingId: "meet-1" },
      "/me/onlineMeetings/": { value: [] },
    });
    const result = await listMeetingTranscriptsTool.handler(client, { event_id: "evt1" });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.count).toBe(0);
    expect(parsed.note).toBeDefined();
  });

  it("returns transcript list with metadata", async () => {
    const client = makeMultiApiClient({
      "/me/events/": { id: "evt1", isOnlineMeeting: true, onlineMeetingId: "meet-1" },
      "/me/onlineMeetings/": {
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
      "/me/events/": { id: "evt1", isOnlineMeeting: true, onlineMeetingId: "meet-42" },
      "/me/onlineMeetings/": { value: [{ id: "t1", createdDateTime: null, endDateTime: null }] },
    });
    const result = await listMeetingTranscriptsTool.handler(client, { event_id: "evt1" });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.meeting_id).toBe("meet-42");
  });
});
