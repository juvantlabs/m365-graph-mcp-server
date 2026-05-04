import type { Client } from "@microsoft/microsoft-graph-client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { _resetConfirmationTokens } from "../../src/auth/confirmation_tokens.js";
import { cancelEventTool } from "../../src/tools/cancel_event.js";

beforeEach(() => _resetConfirmationTokens());
afterEach(() => _resetConfirmationTokens());

function makeClient(eventMetadata: unknown): {
  apiCalls: string[];
  postBodies: unknown[];
  client: Client;
} {
  const apiCalls: string[] = [];
  const postBodies: unknown[] = [];
  const get = vi.fn().mockResolvedValue(eventMetadata);
  const post = vi.fn().mockImplementation((b: unknown) => {
    postBodies.push(b);
    return Promise.resolve({});
  });
  const api = vi.fn().mockImplementation((path: string) => {
    apiCalls.push(path);
    return { get, post };
  });
  return { apiCalls, postBodies, client: { api } as unknown as Client };
}

describe("cancelEventTool — phase 1 (preview)", () => {
  it("requires event_id", async () => {
    const { client } = makeClient({});
    await expect(cancelEventTool.handler(client, {})).rejects.toThrow(
      "'event_id' must be a non-empty string",
    );
  });

  it("returns event preview + token, does NOT call /cancel", async () => {
    const event = { id: "e1", subject: "Standup", start: {}, end: {} };
    const c = makeClient(event);
    const resp = await cancelEventTool.handler(c.client, { event_id: "e1" });

    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.preview).toBeDefined();
    expect(parsed.preview.event.subject).toBe("Standup");
    expect(parsed.preview.confirmation_token).toMatch(/^[0-9a-f]{32}$/);
    expect(c.apiCalls).toEqual(["/me/events/e1"]);
    expect(c.postBodies).toHaveLength(0);
  });
});

describe("cancelEventTool — phase 2 (execute)", () => {
  it("rejects unknown token", async () => {
    const c = makeClient({});
    await expect(
      cancelEventTool.handler(c.client, {
        event_id: "e1",
        confirmation_token: "0".repeat(32),
      }),
    ).rejects.toThrow("token_unknown");
  });

  it("includes Comment in the POST body when comment was given in preview", async () => {
    const event = { id: "e1", subject: "x", start: {}, end: {} };
    const c = makeClient(event);

    const phase1 = await cancelEventTool.handler(c.client, {
      event_id: "e1",
      comment: "Sorry, sick today",
    });
    const { confirmation_token } = JSON.parse(
      (phase1.content[0] as { type: string; text: string }).text,
    ).preview;

    const phase2 = await cancelEventTool.handler(c.client, {
      event_id: "e1",
      comment: "Sorry, sick today",
      confirmation_token,
    });
    const parsed = JSON.parse((phase2.content[0] as { type: string; text: string }).text);
    expect(parsed.cancelled.event_id).toBe("e1");
    expect(c.apiCalls).toContain("/me/events/e1/cancel");
    expect(c.postBodies[0]).toEqual({ Comment: "Sorry, sick today" });
  });

  it("rejects when comment differs between preview and execute (spec mismatch)", async () => {
    const event = { id: "e1", subject: "x", start: {}, end: {} };
    const c = makeClient(event);

    const phase1 = await cancelEventTool.handler(c.client, {
      event_id: "e1",
      comment: "Original",
    });
    const { confirmation_token } = JSON.parse(
      (phase1.content[0] as { type: string; text: string }).text,
    ).preview;

    await expect(
      cancelEventTool.handler(c.client, {
        event_id: "e1",
        comment: "Different comment",
        confirmation_token,
      }),
    ).rejects.toThrow("spec_mismatch");
  });
});
