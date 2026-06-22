/**
 * Tool: m365-graph:get_transcript
 *
 * Fetch the text content of a Teams meeting transcript. The Graph API
 * returns VTT (WebVTT subtitle format); this tool strips the timing
 * markers and returns clean readable text.
 *
 * Required Graph scope: OnlineMeetingTranscript.Read.All (delegated,
 * admin consent required).
 *
 * Caps (defense-in-depth bounds, configurable via env):
 *   - Raw VTT stream:   M365_TRANSCRIPT_MAX_BYTES (default 10 000 000 / 10 MB)
 *   - Parsed-text page: M365_TRANSCRIPT_MAX_CHARS (default 200 000 chars)
 *
 * Paging: when a transcript exceeds the per-call char cap, callers
 * iterate by passing the returned `next_offset` back as `offset` on
 * the next call. The Graph content endpoint returns the whole VTT in
 * one blob — there is no server-side paging primitive — so paging is
 * implemented client-side by slicing the parsed text.
 *
 * Input:
 *   meeting_id     (string, required)  — onlineMeeting id from
 *                  list_meeting_transcripts
 *   transcript_id  (string, required)  — transcript id from
 *                  list_meeting_transcripts
 *   offset         (integer, optional, 0..2_000_000_000, default 0)
 *                  — character offset into the parsed transcript text
 *   max_chars      (integer, optional, 1..M365_TRANSCRIPT_MAX_CHARS,
 *                  default M365_TRANSCRIPT_MAX_CHARS) — max chars to
 *                  return in this response
 *
 * Output: plain-text transcript slice (VTT markup removed) plus
 *   metadata (meeting_id, transcript_id, offset, char_count,
 *   next_offset, total_char_count, truncated, vtt_truncated).
 */

import type { Client } from "@microsoft/microsoft-graph-client";

import { validateRequiredString, validateOptionalInteger } from "../types/validators.js";
import type { Tool, ToolDefinition, ToolHandler, ToolResponse } from "../types/tool.js";

// Defaults are generous enough for multi-hour meetings but still bounded
// against the audit-S7 whole-file-buffering anti-pattern. Both are
// overridable via env vars at process start (per-tenant subprocess, so
// the env is the right knob — no per-call override of the absolute cap).
const DEFAULT_MAX_VTT_BYTES = 10_000_000; // 10 MB raw VTT (~8h+ of speech)
const DEFAULT_CONTENT_CHAR_CAP = 200_000; // 200k chars parsed (~2h of speech)

function readPositiveIntEnv(name: string, fallback: number): number {
  const raw = process.env[name];
  if (raw === undefined || raw === "") return fallback;
  const n = Number(raw);
  if (!Number.isInteger(n) || n <= 0) return fallback;
  return n;
}

function getMaxVttBytes(): number {
  return readPositiveIntEnv("M365_TRANSCRIPT_MAX_BYTES", DEFAULT_MAX_VTT_BYTES);
}

function getMaxChars(): number {
  return readPositiveIntEnv("M365_TRANSCRIPT_MAX_CHARS", DEFAULT_CONTENT_CHAR_CAP);
}

const definition: ToolDefinition = {
  name: "m365-graph:get_transcript",
  description:
    "Fetch the text content of a Teams meeting transcript. VTT timing markers " +
    "are stripped; returns clean readable text. Long transcripts can be paged " +
    "via the offset + max_chars inputs (see next_offset in the response). " +
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
      offset: {
        type: "integer",
        minimum: 0,
        maximum: 2_000_000_000,
        description:
          "Character offset into the parsed transcript text (default 0). " +
          "Pass the previous response's next_offset to fetch the next page.",
      },
      max_chars: {
        type: "integer",
        minimum: 1,
        description:
          "Maximum number of transcript characters to return in this response. " +
          "Defaults to and is capped by M365_TRANSCRIPT_MAX_CHARS " +
          "(default 200 000).",
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

  const maxVttBytes = getMaxVttBytes();
  const maxChars = getMaxChars();

  const offset = validateOptionalInteger(args.offset, "offset", {
    min: 0,
    max: 2_000_000_000,
    default: 0,
  });
  const pageChars = validateOptionalInteger(args.max_chars, "max_chars", {
    min: 1,
    max: maxChars,
    default: maxChars,
  });

  const endpoint =
    `/me/onlineMeetings/${encodeURIComponent(meetingId)}` +
    `/transcripts/${encodeURIComponent(transcriptId)}/content`;

  const rawResponse: unknown = await graph
    .api(endpoint)
    .query({ $format: "text/vtt" })
    .get();

  // Graph SDK may return a ReadableStream for binary/text content types.
  // Cap raw VTT at maxVttBytes before parsing (audit S7: whole-file
  // buffering bound). We flag `vtt_truncated` so callers know the tail
  // of the *upstream* VTT was not read (distinct from output paging).
  let vttString: string;
  let vttTruncated = false;
  if (rawResponse instanceof ReadableStream) {
    const reader = rawResponse.getReader();
    const chunks: Buffer[] = [];
    let totalBytes = 0;
    let done = false;
    while (!done && totalBytes < maxVttBytes) {
      const { done: d, value } = await reader.read();
      done = d;
      if (value) {
        const buf = Buffer.from(value);
        chunks.push(buf);
        totalBytes += buf.byteLength;
      }
    }
    if (!done) {
      vttTruncated = true;
      reader.cancel().catch(() => undefined);
    }
    vttString = Buffer.concat(chunks).toString("utf-8");
    if (vttString.length > maxVttBytes) {
      vttString = vttString.slice(0, maxVttBytes);
      vttTruncated = true;
    }
  } else {
    const s = String(rawResponse ?? "");
    if (s.length > maxVttBytes) {
      vttString = s.slice(0, maxVttBytes);
      vttTruncated = true;
    } else {
      vttString = s;
    }
  }

  const fullText = parseVtt(vttString);
  const totalChars = fullText.length;

  // Slice the requested page out of the parsed text.
  const sliceStart = Math.min(offset, totalChars);
  const sliceEnd = Math.min(sliceStart + pageChars, totalChars);
  const content = fullText.slice(sliceStart, sliceEnd);
  const nextOffset = sliceEnd < totalChars ? sliceEnd : null;
  // `truncated` retained for backward-compat with v0.2.x callers: true when
  // the returned slice does not reach the end of the (in-memory) parsed
  // text. Combined with `vtt_truncated`, callers can distinguish a
  // page-of-many response from a hard upstream-byte loss.
  const truncated = nextOffset !== null;

  return {
    content: [
      {
        type: "text",
        text: JSON.stringify(
          {
            meeting_id: meetingId,
            transcript_id: transcriptId,
            offset: sliceStart,
            char_count: content.length,
            next_offset: nextOffset,
            total_char_count: totalChars,
            truncated,
            vtt_truncated: vttTruncated,
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
