/**
 * Tool: m365-graph:list_meeting_transcripts
 *
 * List available transcripts for a Teams meeting, identified by its
 * calendar event ID. Transcripts are available post-meeting and only
 * when recording was enabled by the meeting organizer.
 *
 * Required Graph scopes (delegated, admin consent required):
 *   Calendars.Read          — resolve event → onlineMeeting join URL
 *   OnlineMeetings.Read     — filter /me/onlineMeetings by JoinWebUrl
 *   OnlineMeetingTranscript.Read.All — list transcripts
 *
 * Input:
 *   event_id  (string, required) — calendar event id from list_events /
 *             search_events / get_event
 *
 * Output: list of transcript metadata with id, created_at, end_at.
 *   Pass transcript id + meeting_id to get_transcript to read the content.
 *
 * Returns an empty list (not an error) when the meeting ended without
 * a transcript (recording disabled or not yet processed).
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import { validateRequiredString } from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const SUBJECT_MAX = 120;

const definition: ToolDefinition = {
  name: "m365-graph:list_meeting_transcripts",
  description:
    "List available transcripts for a Teams meeting identified by its calendar event ID. " +
    "Transcripts are post-meeting only and require recording to have been enabled. " +
    "Returns transcript IDs to pass to get_transcript. Read-only.",
  inputSchema: {
    type: "object",
    properties: {
      event_id: {
        type: "string",
        description: "Calendar event ID from list_events, search_events, or get_event.",
      },
    },
    required: ["event_id"],
  },
};

interface TranscriptSummary {
  id: string;
  meeting_id: string;
  created_at: string | null;
  end_at: string | null;
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const eventId = validateRequiredString(args.event_id, "event_id");

  // Step 1: resolve joinUrl from the calendar event.
  // Graph returns joinUrl (not joinWebUrl) on the event's OnlineMeetingInfo object.
  // onlineMeetingId is not selectable via $select in all tenants; we use
  // onlineMeeting.joinUrl and resolve the onlineMeeting resource via filter instead.
  const event = await graph
    .api(`/me/events/${encodeURIComponent(eventId)}`)
    .select("id,isOnlineMeeting,onlineMeeting,subject")
    .get();

  if (!event.isOnlineMeeting) {
    const subject = String(event.subject ?? eventId).slice(0, SUBJECT_MAX);
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify({
            error: "not_an_online_meeting",
            message: `Event '${subject}' is not a Teams online meeting. Transcripts are only available for online meetings.`,
          }),
        },
      ],
    };
  }

  const onlineMeeting = event.onlineMeeting as Record<string, unknown> | undefined;
  const joinUrl = (onlineMeeting?.joinUrl ?? onlineMeeting?.joinWebUrl) as string | undefined;
  if (!joinUrl) {
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify({
            error: "meeting_id_unavailable",
            message: "Could not resolve Teams meeting join URL from the event.",
          }),
        },
      ],
    };
  }

  // Step 2: resolve onlineMeeting ID via filter on joinUrl.
  // OData string literals: single quotes are escaped by doubling them ('').
  // $select is not supported on /me/onlineMeetings when combined with $filter.
  const escapedJoinUrl = joinUrl.replace(/'/g, "''");
  const meetingResponse = await graph
    .api("/me/onlineMeetings")
    .filter(`JoinWebUrl eq '${escapedJoinUrl}'`)
    .get();

  const meetings: Record<string, unknown>[] = Array.isArray(meetingResponse?.value)
    ? meetingResponse.value
    : [];

  if (meetings.length === 0) {
    return {
      content: [
        {
          type: "text",
          text: JSON.stringify({
            error: "meeting_id_unavailable",
            message: "Could not resolve onlineMeeting from the event's join URL. The meeting may have expired or transcript access may require admin consent (OnlineMeetings.Read + OnlineMeetingTranscript.Read.All).",
          }),
        },
      ],
    };
  }

  const meetingId = String(meetings[0].id ?? "");

  // Step 3: list transcripts for the meeting.
  const response = await graph
    .api(`/me/onlineMeetings/${encodeURIComponent(meetingId)}/transcripts`)
    .select("id,meetingId,createdDateTime,endDateTime")
    .get();

  const items: Record<string, unknown>[] = Array.isArray(response?.value)
    ? response.value
    : [];

  const transcripts: TranscriptSummary[] = items.map((t) => ({
    id: String(t.id ?? ""),
    meeting_id: meetingId,
    created_at: (t.createdDateTime as string | undefined) ?? null,
    end_at: (t.endDateTime as string | undefined) ?? null,
  }));

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            event_id: eventId,
            meeting_id: meetingId,
            count: transcripts.length,
            transcripts,
            note:
              transcripts.length === 0
                ? "No transcripts found. Recording may not have been enabled, or the transcript is still processing."
                : undefined,
          },
          null,
          2,
        ),
      },
    ],
  };
};

export const listMeetingTranscriptsTool: Tool = { category: "read", definition, handler };
