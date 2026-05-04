import type { Client } from "@microsoft/microsoft-graph-client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { _resetConfirmationTokens } from "../../src/auth/confirmation_tokens.js";
import { deleteFileTool } from "../../src/tools/delete_file.js";

beforeEach(() => _resetConfirmationTokens());
afterEach(() => _resetConfirmationTokens());

function makeClient(itemMetadata: unknown): {
  apiCalls: string[];
  deleteCalled: number;
  client: Client;
} {
  const apiCalls: string[] = [];
  let deleteCalled = 0;
  const get = vi.fn().mockResolvedValue(itemMetadata);
  const del = vi.fn().mockImplementation(() => {
    deleteCalled++;
    return Promise.resolve(undefined);
  });
  const api = vi.fn().mockImplementation((path: string) => {
    apiCalls.push(path);
    return { get, delete: del };
  });
  return {
    apiCalls,
    get deleteCalled() {
      return deleteCalled;
    },
    client: { api } as unknown as Client,
  };
}

describe("deleteFileTool — phase 1 (preview)", () => {
  it("requires item_id", async () => {
    const { client } = makeClient({});
    await expect(deleteFileTool.handler(client, {})).rejects.toThrow(
      "'item_id' must be a non-empty string",
    );
  });

  it("returns a preview + confirmation_token, does NOT delete", async () => {
    const item = {
      id: "i1",
      name: "report.pdf",
      size: 2048,
      webUrl: "https://example/i1",
      parentReference: { path: "/drive/root:/Documents" },
    };
    const c = makeClient(item);
    const resp = await deleteFileTool.handler(c.client, { item_id: "i1" });

    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.preview).toBeDefined();
    expect(parsed.preview.item.name).toBe("report.pdf");
    expect(parsed.preview.item.type).toBe("file");
    expect(parsed.preview.confirmation_token).toMatch(/^[0-9a-f]{32}$/);
    expect(parsed.preview.expires_in_seconds).toBe(300);
    expect(c.deleteCalled).toBe(0);
  });
});

describe("deleteFileTool — phase 2 (execute)", () => {
  it("rejects an unknown token with token_unknown", async () => {
    const c = makeClient({});
    await expect(
      deleteFileTool.handler(c.client, {
        item_id: "i1",
        confirmation_token: "deadbeef".repeat(4),
      }),
    ).rejects.toThrow("token_unknown");
    expect(c.deleteCalled).toBe(0);
  });

  it("rejects a token that was issued for a different item_id", async () => {
    const item = { id: "i1", name: "x", size: 0 };
    const c = makeClient(item);

    // Phase 1: preview for item A
    const phase1 = await deleteFileTool.handler(c.client, { item_id: "A" });
    const { confirmation_token } = JSON.parse(
      (phase1.content[0] as { type: string; text: string }).text,
    ).preview;

    // Phase 2: try to use token with item B
    await expect(
      deleteFileTool.handler(c.client, {
        item_id: "B",
        confirmation_token,
      }),
    ).rejects.toThrow("spec_mismatch");
    expect(c.deleteCalled).toBe(0);
  });

  it("executes DELETE on a valid token + matching spec", async () => {
    const item = { id: "i1", name: "x", size: 0 };
    const c = makeClient(item);

    const phase1 = await deleteFileTool.handler(c.client, { item_id: "i1" });
    const { confirmation_token } = JSON.parse(
      (phase1.content[0] as { type: string; text: string }).text,
    ).preview;

    const phase2 = await deleteFileTool.handler(c.client, {
      item_id: "i1",
      confirmation_token,
    });
    const parsed = JSON.parse((phase2.content[0] as { type: string; text: string }).text);
    expect(parsed.deleted.item_id).toBe("i1");
    expect(c.deleteCalled).toBe(1);
  });

  it("token is single-use — second attempt fails", async () => {
    const item = { id: "i1", name: "x", size: 0 };
    const c = makeClient(item);

    const phase1 = await deleteFileTool.handler(c.client, { item_id: "i1" });
    const { confirmation_token } = JSON.parse(
      (phase1.content[0] as { type: string; text: string }).text,
    ).preview;

    await deleteFileTool.handler(c.client, { item_id: "i1", confirmation_token });
    await expect(
      deleteFileTool.handler(c.client, { item_id: "i1", confirmation_token }),
    ).rejects.toThrow("token_unknown");
    expect(c.deleteCalled).toBe(1);
  });
});
