/**
 * Tool: m365-graph:get_transcript
 *
 * Fetch the text content of a Teams meeting transcript. The Graph API
 * returns VTT (WebVTT subtitle format); this tool strips the timing
 * markers and returns clean readable text, capped at 30 000 chars.
 *
 * Required Graph scope: OnlineMeetingTranscript.Read.All (delegated,
 * admin consent required).
 *
 * Input:
 *   meeting_id     (string, required) — onlineMeeting id from
 *                  list_meeting_transcripts
 *   transcript_id  (string, required) — transcript id from
 *                  list_meeting_transcripts
 *
 * Output: plain text transcript content (VTT markup removed) plus
 *   metadata (meeting_id, transcript_id, char_count, truncated).
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import { validateRequiredString } from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

const CONTENT_CHAR_CAP = 30_000;

const definition: ToolDefinition = {
  name: "m365-graph:get_transcript",
  description:
    "Fetch the text content of a Teams meeting transcript. VTT timing markers " +
    "are stripped; returns clean readable text capped at 30 000 chars. " +
    "Use list_meeting_transcripts to get the meeting_id and transcript_id. Read-only.",
  inputSchema: {
    type: "object",
    properties: {
      meeting_id: {
        type: "string",
        description: "onlineMeeting ID from list_meeting_transcripts.",
      },
      transcript_id: {
        type: "string",
        description: "Transcript ID from list_meeting_transcripts.",
      },
    },
    required: ["meeting_id", "transcript_id"],
  },
};

/**
 * Parse WebVTT content into plain text.
 * Strips: WEBVTT header, NOTE blocks, timestamp lines (HH:MM:SS.mmm --> ...),
 * numeric cue sequence lines, and blank lines between cues.
 * Keeps: cue text lines (speaker names + spoken content).
 */
export function parseVtt(vtt: string): string {
  const lines = vtt.split(/\r?\n/);
  const textLines: string[] = [];
  const TIMESTAMP_RE = /^\d{2}:\d{2}:\d{2}\.\d{3}\s+-->\s+\d{2}:\d{2}:\d{2}\.\d{3}/;
  const SEQUENCE_RE = /^\d+\s*$/;

  let skipNote = false;
  for (const line of lines) {
    if (line.startsWith("WEBVTT")) { skipNote = false; continue; }
    if (line.startsWith("NOTE")) { skipNote = true; continue; }
    if (skipNote && line.trim() === "") { skipNote = false; continue; }
    if (skipNote) continue;
    if (TIMESTAMP_RE.test(line)) continue;
    if (SEQUENCE_RE.test(line)) continue;
    if (line.trim() === "") continue;
    textLines.push(line.trim());
  }
  return textLines.join("\n");
}

const handler: ToolHandler = async (
  graph: Client,
  args: Record<string, unknown>,
): Promise<ToolResponse> => {
  const meetingId = validateRequiredString(args.meeting_id, "meeting_id");
  const transcriptId = validateRequiredString(args.transcript_id, "transcript_id");

  const endpoint =
    `/me/onlineMeetings/${encodeURIComponent(meetingId)}` +
    `/transcripts/${encodeURIComponent(transcriptId)}/content`;

  const rawVtt: string = await graph
    .api(endpoint)
    .query({ $format: "text/vtt" })
    .get();

  const text = parseVtt(String(rawVtt ?? ""));
  const truncated = text.length > CONTENT_CHAR_CAP;
  const content = truncated ? text.slice(0, CONTENT_CHAR_CAP) : text;

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            meeting_id: meetingId,
            transcript_id: transcriptId,
            char_count: content.length,
            truncated,
            transcript: content,
          },
          null,
          2,
        ),
      },
    ],
  };
};

export const getTranscriptTool: Tool = { category: "read", definition, handler };
