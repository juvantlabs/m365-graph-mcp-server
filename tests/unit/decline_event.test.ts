import type { Client } from "@microsoft/microsoft-graph-client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { _resetConfirmationTokens } from "../../src/auth/confirmation_tokens.js";
import { declineEventTool } from "../../src/tools/decline_event.js";

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

describe("declineEventTool — phase 1 (preview)", () => {
  it("requires event_id", async () => {
    const { client } = makeClient({});
    await expect(declineEventTool.handler(client, {})).rejects.toThrow(
      "'event_id' must be a non-empty string",
    );
  });

  it("returns event preview + token, does NOT call /decline", async () => {
    const event = { id: "e1", subject: "Standup", start: {}, end: {} };
    const c = makeClient(event);
    const resp = await declineEventTool.handler(c.client, { event_id: "e1" });

    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.preview.event.subject).toBe("Standup");
    expect(parsed.preview.send_response).toBe(true); // default
    expect(parsed.preview.confirmation_token).toMatch(/^[0-9a-f]{32}$/);
    expect(c.apiCalls).toEqual(["/me/events/e1"]);
    expect(c.postBodies).toHaveLength(0);
  });

  it("send_response=false is recorded in the preview + spec", async () => {
    const event = { id: "e1", subject: "x", start: {}, end: {} };
    const c = makeClient(event);
    const resp = await declineEventTool.handler(c.client, {
      event_id: "e1",
      send_response: false,
    });
    const preview = JSON.parse((resp.content[0] as { type: string; text: string }).text).preview;
    expect(preview.send_response).toBe(false);
  });
});

describe("declineEventTool — phase 2 (execute)", () => {
  it("rejects unknown token", async () => {
    const c = makeClient({});
    await expect(
      declineEventTool.handler(c.client, {
        event_id: "e1",
        confirmation_token: "0".repeat(32),
      }),
    ).rejects.toThrow("token_unknown");
  });

  it("POSTs to /decline with sendResponse + comment from the spec", async () => {
    const event = { id: "e1", subject: "x", start: {}, end: {} };
    const c = makeClient(event);

    const phase1 = await declineEventTool.handler(c.client, {
      event_id: "e1",
      comment: "Conflict",
      send_response: true,
    });
    const { confirmation_token } = JSON.parse(
      (phase1.content[0] as { type: string; text: string }).text,
    ).preview;

    const phase2 = await declineEventTool.handler(c.client, {
      event_id: "e1",
      comment: "Conflict",
      send_response: true,
      confirmation_token,
    });
    const parsed = JSON.parse((phase2.content[0] as { type: string; text: string }).text);
    expect(parsed.declined.event_id).toBe("e1");
    expect(c.apiCalls).toContain("/me/events/e1/decline");
    expect(c.postBodies[0]).toEqual({ sendResponse: true, comment: "Conflict" });
  });

  it("silent decline (send_response=false) sends sendResponse:false", async () => {
    const event = { id: "e1", subject: "x", start: {}, end: {} };
    const c = makeClient(event);

    const phase1 = await declineEventTool.handler(c.client, {
      event_id: "e1",
      send_response: false,
    });
    const { confirmation_token } = JSON.parse(
      (phase1.content[0] as { type: string; text: string }).text,
    ).preview;

    await declineEventTool.handler(c.client, {
      event_id: "e1",
      send_response: false,
      confirmation_token,
    });
    expect(c.postBodies[0]).toEqual({ sendResponse: false });
  });

  it("rejects when send_response differs between preview and execute (spec mismatch)", async () => {
    const event = { id: "e1", subject: "x", start: {}, end: {} };
    const c = makeClient(event);

    const phase1 = await declineEventTool.handler(c.client, {
      event_id: "e1",
      send_response: true,
    });
    const { confirmation_token } = JSON.parse(
      (phase1.content[0] as { type: string; text: string }).text,
    ).preview;

    await expect(
      declineEventTool.handler(c.client, {
        event_id: "e1",
        send_response: false,
        confirmation_token,
      }),
    ).rejects.toThrow("spec_mismatch");
  });
});
