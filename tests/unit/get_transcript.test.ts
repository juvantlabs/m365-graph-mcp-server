import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

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

describe("getTranscriptTool handler", () => {
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

  it("returns clean text from VTT content", async () => {
    const client = makeChainClient(SAMPLE_VTT);
    const result = await getTranscriptTool.handler(client, {
      meeting_id: "meet-1",
      transcript_id: "trans-1",
    });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.transcript).toContain("Antonio Gatti: Good morning everyone.");
    expect(parsed.transcript).not.toContain("-->");
    expect(parsed.truncated).toBe(false);
  });

  it("truncates transcript at 30 000 chars", async () => {
    const longLine = "Speaker: " + "x".repeat(1000);
    const vtt = "WEBVTT\n\n" + Array.from({ length: 50 }, (_, i) =>
      `${i + 1}\n00:0${Math.floor(i / 10)}:0${i % 10}.000 --> 00:0${Math.floor(i / 10)}:0${(i % 10) + 1}.000\n${longLine}\n`
    ).join("\n");

    const client = makeChainClient(vtt);
    const result = await getTranscriptTool.handler(client, {
      meeting_id: "m",
      transcript_id: "t",
    });
    const parsed = JSON.parse((result.content[0] as { text: string }).text);
    expect(parsed.truncated).toBe(true);
    expect(parsed.char_count).toBeLessThanOrEqual(30_000);
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
