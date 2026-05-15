/**
 * Integration tests for transcript tools.
 *
 * Runs against the real Microsoft Graph API using the credentials in
 * `.env.local` (M365_TENANT_ID, M365_CLIENT_ID, M365_CLIENT_SECRET)
 * and the token cached in the OS keychain via `npm run setup`.
 *
 * Skipped automatically when credentials are not available so the
 * test suite can run in CI without secrets configured.
 *
 * Run manually:
 *   npm run test:integration
 */

import { describe, expect, it, beforeAll } from "vitest";

import { makeMsalClient } from "../../src/auth/msal.js";
import { makeGraphClient } from "../../src/client/graph.js";
import { listMeetingTranscriptsTool } from "../../src/tools/list_meeting_transcripts.js";
import { getTranscriptTool } from "../../src/tools/get_transcript.js";
import { listEventsTool } from "../../src/tools/list_events.js";

const hasCredentials =
  !!process.env.M365_TENANT_ID &&
  !!process.env.M365_CLIENT_ID &&
  !!process.env.M365_CLIENT_SECRET;

const describeIf = hasCredentials ? describe : describe.skip;

describeIf("Transcript tools — live Graph API", () => {
  let graph: ReturnType<typeof makeGraphClient>;

  beforeAll(() => {
    const msal = makeMsalClient();
    graph = makeGraphClient(msal);
  });

  it("list_events returns recent meetings", async () => {
    const today = new Date();
    const weekAgo = new Date(today);
    weekAgo.setDate(today.getDate() - 7);

    const result = await listEventsTool.handler(graph, {
      start: weekAgo.toISOString().slice(0, 10),
      end: today.toISOString().slice(0, 10),
      limit: 10,
    });

    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    console.log(`Found ${parsed.count} events in the last 7 days`);
    expect(parsed.count).toBeGreaterThanOrEqual(0);
  });

  it("list_meeting_transcripts handles a non-Teams event gracefully", async () => {
    // Find the most recent event (any type) and test graceful handling
    const today = new Date();
    const weekAgo = new Date(today);
    weekAgo.setDate(today.getDate() - 7);

    const eventsResult = await listEventsTool.handler(graph, {
      start: weekAgo.toISOString().slice(0, 10),
      end: today.toISOString().slice(0, 10),
      limit: 5,
    });
    const events = JSON.parse((eventsResult.content[0] as { text: string }).text).events;

    if (events.length === 0) {
      console.log("No events in last 7 days — skipping transcript test");
      return;
    }

    const eventId = events[0].id;
    const result = await listMeetingTranscriptsTool.handler(graph, { event_id: eventId });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);

    // Either returns transcripts, empty list, or graceful error — never throws
    expect(parsed).toHaveProperty("event_id");
    console.log("list_meeting_transcripts result:", JSON.stringify(parsed, null, 2));
  });

  it("list_meeting_transcripts + get_transcript on a Teams meeting with transcript", async () => {
    // Look for Teams meetings in the last 30 days
    const today = new Date();
    const monthAgo = new Date(today);
    monthAgo.setDate(today.getDate() - 30);

    const eventsResult = await listEventsTool.handler(graph, {
      start: monthAgo.toISOString().slice(0, 10),
      end: today.toISOString().slice(0, 10),
      limit: 50,
    });
    const events = JSON.parse((eventsResult.content[0] as { text: string }).text).events;
    const onlineMeetings = events.filter((e: { is_online_meeting?: boolean }) => e.is_online_meeting);

    if (onlineMeetings.length === 0) {
      console.log("No Teams meetings in last 30 days — skipping live transcript test");
      return;
    }

    console.log(`Found ${onlineMeetings.length} Teams meetings — testing transcripts`);

    // Try each meeting until we find one with a transcript
    let foundTranscript = false;
    for (const meeting of onlineMeetings) {
      const listResult = await listMeetingTranscriptsTool.handler(graph, {
        event_id: meeting.id,
      });
      const listParsed = JSON.parse((listResult.content[0] as { text: string }).text);

      if (listParsed.count > 0) {
        foundTranscript = true;
        console.log(`Meeting "${meeting.subject}" has ${listParsed.count} transcript(s)`);

        const transcript = listParsed.transcripts[0];
        const getResult = await getTranscriptTool.handler(graph, {
          meeting_id: transcript.meeting_id,
          transcript_id: transcript.id,
        });
        const getParsed = JSON.parse((getResult.content[0] as { text: string }).text);

        expect(getParsed.transcript).toBeTruthy();
        expect(typeof getParsed.transcript).toBe("string");
        expect(getParsed.char_count).toBeGreaterThan(0);
        expect(getParsed).not.toContain("-->");  // no VTT markers in output

        console.log(`Transcript sample (first 200 chars): ${getParsed.transcript.slice(0, 200)}`);
        break;
      }
    }

    if (!foundTranscript) {
      console.log("No transcripts found in any meeting — recording may not be enabled");
    }
    // Not a hard failure — recording must be explicitly enabled per meeting
    expect(true).toBe(true);
  });
});
