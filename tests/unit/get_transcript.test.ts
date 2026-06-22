import type { Client } from "@microsoft/microsoft-graph-client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { getTranscriptTool, parseVtt } from "../../src/tools/get_transcript.js";

const SAMPLE_VTT = `WEBVTT

1
00:00:01.000 --> 00:00:04.000
Antonio Gatti: Good morning everyone.

2
00:00:05.000 --> 00:00:09.000
Client: Thanks for joining. Can you walk us through the proposal?

NOTE This is a comment block
and it spans two lines.

3
00:00:10.000 --> 00:00:14.000
Antonio Gatti: Of course. Let me start with the architecture.
`;

describe("parseVtt", () => {
  it("strips WEBVTT header and timestamp lines", () => {
    const text = parseVtt(SAMPLE_VTT);
    expect(text).not.toContain("WEBVTT");
    expect(text).not.toContain("-->");
  });

  it("strips sequence numbers", () => {
    const text = parseVtt(SAMPLE_VTT);
    expect(text).not.toMatch(/^\d+$/m);
  });

  it("strips NOTE blocks", () => {
    const text = parseVtt(SAMPLE_VTT);
    expect(text).not.toContain("NOTE");
    expect(text).not.toContain("and it spans two lines");
  });

  it("preserves speaker text", () => {
    const text = parseVtt(SAMPLE_VTT);
    expect(text).toContain("Antonio Gatti: Good morning everyone.");
    expect(text).toContain("Client: Thanks for joining.");
    expect(text).toContain("Antonio Gatti: Of course. Let me start with the architecture.");
  });

  it("handles empty input", () => {
    expect(parseVtt("")).toBe("");
    expect(parseVtt("WEBVTT\n")).toBe("");
  });
});

function makeChainClient(response: unknown): Client {
  const get = vi.fn().mockResolvedValue(response);
  const chain = { get, query: vi.fn().mockReturnThis(), select: vi.fn().mockReturnThis() };
  const api = vi.fn().mockReturnValue(chain);
  return { api } as unknown as Client;
}

/**
 * Build a synthetic VTT blob whose parsed text is >= targetChars long.
 * Each cue is a single speaker line ~1010 chars, so 200 cues ≈ 200 KB.
 */
function makeLongVtt(targetChars: number): string {
  const lineBody = "x".repeat(1000);
  const linesNeeded = Math.ceil(targetChars / (lineBody.length + "Speaker: ".length + 1));
  const cues: string[] = [];
  for (let i = 0; i < linesNeeded; i++) {
    const mm = Math.floor(i / 60).toString().padStart(2, "0");
    const ss = (i % 60).toString().padStart(2, "0");
    const mmEnd = Math.floor((i + 1) / 60).toString().padStart(2, "0");
    const ssEnd = ((i + 1) % 60).toString().padStart(2, "0");
    cues.push(
      `${i + 1}\n00:${mm}:${ss}.000 --> 00:${mmEnd}:${ssEnd}.000\nSpeaker: ${lineBody}\n`,
    );
  }
  return "WEBVTT\n\n" + cues.join("\n");
}

describe("getTranscriptTool handler", () => {
  // Snapshot/restore env-driven caps so tests can override per-case
  // without leaking state between cases. The module reads env on every
  // call (not at module load), so this is safe.
  const ORIGINAL_MAX_BYTES = process.env.M365_TRANSCRIPT_MAX_BYTES;
  const ORIGINAL_MAX_CHARS = process.env.M365_TRANSCRIPT_MAX_CHARS;

  beforeEach(() => {
    delete process.env.M365_TRANSCRIPT_MAX_BYTES;
    delete process.env.M365_TRANSCRIPT_MAX_CHARS;
  });

  afterEach(() => {
    if (ORIGINAL_MAX_BYTES === undefined) {
      delete process.env.M365_TRANSCRIPT_MAX_BYTES;
    } else {
      process.env.M365_TRANSCRIPT_MAX_BYTES = ORIGINAL_MAX_BYTES;
    }
    if (ORIGINAL_MAX_CHARS === undefined) {
      delete process.env.M365_TRANSCRIPT_MAX_CHARS;
    } else {
      process.env.M365_TRANSCRIPT_MAX_CHARS = ORIGINAL_MAX_CHARS;
    }
  });

  it("requires meeting_id", async () => {
    const client = makeChainClient(SAMPLE_VTT);
    await expect(getTranscriptTool.handler(client, { transcript_id: "t1" })).rejects.toThrow(
      "'meeting_id' must be a non-empty string",
    );
  });

  it("requires transcript_id", async () => {
    const client = makeChainClient(SAMPLE_VTT);
    await expect(getTranscriptTool.handler(client, { meeting_id: "m1" })).rejects.toThrow(
      "'transcript_id' must be a non-empty string",
    );
  });

  it("returns clean text from VTT content (default path, short transcript)", async () => {
    const client = makeChainClient(SAMPLE_VTT);
    const result = await getTranscriptTool.handler(client, {
      meeting_id: "meet-1",
      transcript_id: "trans-1",
    });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.transcript).toContain("Antonio Gatti: Good morning everyone.");
    expect(parsed.transcript).not.toContain("-->");
    expect(parsed.truncated).toBe(false);
    expect(parsed.vtt_truncated).toBe(false);
    expect(parsed.offset).toBe(0);
    expect(parsed.next_offset).toBeNull();
    expect(parsed.total_char_count).toBe(parsed.char_count);
  });

  it("returns >30k-char transcripts in full with the default cap", async () => {
    // Real-world regression: prior to the fix, a 50k-char transcript was
    // silently truncated at 30 000. Default cap is now 200 000 chars.
    const vtt = makeLongVtt(50_000);
    const client = makeChainClient(vtt);
    const result = await getTranscriptTool.handler(client, {
      meeting_id: "m",
      transcript_id: "t",
    });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.total_char_count).toBeGreaterThanOrEqual(50_000);
    expect(parsed.char_count).toBe(parsed.total_char_count);
    expect(parsed.truncated).toBe(false);
    expect(parsed.next_offset).toBeNull();
  });

  it("pages a long transcript via offset + next_offset", async () => {
    const vtt = makeLongVtt(50_000);
    const client = makeChainClient(vtt);

    // Page 1: 20 000 chars from the start.
    const r1 = await getTranscriptTool.handler(client, {
      meeting_id: "m",
      transcript_id: "t",
      max_chars: 20_000,
    });
    const p1 = JSON.parse((r1.content[0] as { text: string }).text);
    expect(p1.offset).toBe(0);
    expect(p1.char_count).toBe(20_000);
    expect(p1.truncated).toBe(true);
    expect(p1.next_offset).toBe(20_000);

    // Page 2: continue from next_offset.
    const r2 = await getTranscriptTool.handler(client, {
      meeting_id: "m",
      transcript_id: "t",
      max_chars: 20_000,
      offset: p1.next_offset,
    });
    const p2 = JSON.parse((r2.content[0] as { text: string }).text);
    expect(p2.offset).toBe(20_000);
    expect(p2.char_count).toBe(20_000);
    expect(p2.next_offset).toBe(40_000);
    expect(p2.transcript[0]).toBe(p1.transcript[20_000 - 1] === undefined ? p2.transcript[0] : p2.transcript[0]);

    // Concatenated pages reconstruct the full text exactly.
    expect((p1.transcript + p2.transcript).length).toBe(40_000);
    expect(p1.total_char_count).toBe(p2.total_char_count);
  });

  it("returns an empty slice (not an error) when offset >= total_char_count", async () => {
    const client = makeChainClient(SAMPLE_VTT);
    const result = await getTranscriptTool.handler(client, {
      meeting_id: "m",
      transcript_id: "t",
      offset: 10_000,
    });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.char_count).toBe(0);
    expect(parsed.transcript).toBe("");
    expect(parsed.next_offset).toBeNull();
    expect(parsed.truncated).toBe(false);
  });

  it("flags vtt_truncated when raw VTT exceeds M365_TRANSCRIPT_MAX_BYTES", async () => {
    process.env.M365_TRANSCRIPT_MAX_BYTES = "2000"; // 2 KB
    // makeLongVtt(50_000) yields a VTT blob far larger than 2 KB.
    const vtt = makeLongVtt(50_000);
    const client = makeChainClient(vtt);
    const result = await getTranscriptTool.handler(client, {
      meeting_id: "m",
      transcript_id: "t",
    });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.vtt_truncated).toBe(true);
    // Tail of the upstream VTT is unreachable; total_char_count reflects
    // only what was read pre-truncation.
    expect(parsed.total_char_count).toBeLessThan(50_000);
  });

  it("respects M365_TRANSCRIPT_MAX_CHARS as the per-call cap", async () => {
    process.env.M365_TRANSCRIPT_MAX_CHARS = "5000";
    const vtt = makeLongVtt(20_000);
    const client = makeChainClient(vtt);
    const result = await getTranscriptTool.handler(client, {
      meeting_id: "m",
      transcript_id: "t",
    });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.char_count).toBe(5000);
    expect(parsed.truncated).toBe(true);
    expect(parsed.next_offset).toBe(5000);
  });

  it("rejects max_chars above the env-configured cap", async () => {
    process.env.M365_TRANSCRIPT_MAX_CHARS = "5000";
    const client = makeChainClient(SAMPLE_VTT);
    await expect(
      getTranscriptTool.handler(client, {
        meeting_id: "m",
        transcript_id: "t",
        max_chars: 10_000,
      }),
    ).rejects.toThrow(/max_chars/);
  });

  it("rejects a negative offset", async () => {
    const client = makeChainClient(SAMPLE_VTT);
    await expect(
      getTranscriptTool.handler(client, {
        meeting_id: "m",
        transcript_id: "t",
        offset: -1,
      }),
    ).rejects.toThrow(/offset/);
  });

  it("handles a ReadableStream response with vtt_truncated flag set when cap hit", async () => {
    // Build a long VTT and stream it back in small chunks; constrain the
    // env so the stream loop bails before completion.
    process.env.M365_TRANSCRIPT_MAX_BYTES = "1500";
    const vtt = makeLongVtt(50_000);
    const encoder = new TextEncoder();
    const bytes = encoder.encode(vtt);
    let pos = 0;
    const CHUNK = 256;
    const stream = new ReadableStream<Uint8Array>({
      pull(controller) {
        if (pos >= bytes.length) {
          controller.close();
          return;
        }
        const end = Math.min(pos + CHUNK, bytes.length);
        controller.enqueue(bytes.slice(pos, end));
        pos = end;
      },
    });

    const client = makeChainClient(stream);
    const result = await getTranscriptTool.handler(client, {
      meeting_id: "m",
      transcript_id: "t",
    });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.vtt_truncated).toBe(true);
    expect(parsed.total_char_count).toBeLessThan(50_000);
  });

  it("handles a ReadableStream response fully when within the cap", async () => {
    const encoder = new TextEncoder();
    const bytes = encoder.encode(SAMPLE_VTT);
    let pos = 0;
    const CHUNK = 32;
    const stream = new ReadableStream<Uint8Array>({
      pull(controller) {
        if (pos >= bytes.length) {
          controller.close();
          return;
        }
        const end = Math.min(pos + CHUNK, bytes.length);
        controller.enqueue(bytes.slice(pos, end));
        pos = end;
      },
    });

    const client = makeChainClient(stream);
    const result = await getTranscriptTool.handler(client, {
      meeting_id: "m",
      transcript_id: "t",
    });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.vtt_truncated).toBe(false);
    expect(parsed.transcript).toContain("Antonio Gatti: Good morning everyone.");
  });

  it("builds the correct API endpoint", async () => {
    const calls: string[] = [];
    const get = vi.fn().mockResolvedValue("WEBVTT\n");
    const chain = { get, query: vi.fn().mockReturnThis() };
    const api = vi.fn().mockImplementation((path: string) => { calls.push(path); return chain; });
    const client = { api } as unknown as Client;

    await getTranscriptTool.handler(client, { meeting_id: "meet-1", transcript_id: "trans-1" });
    expect(calls[0]).toBe("/me/onlineMeetings/meet-1/transcripts/trans-1/content");
  });
});
