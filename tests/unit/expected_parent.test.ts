import type { Client } from "@microsoft/microsoft-graph-client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { _resetConfirmationTokens } from "../../src/auth/confirmation_tokens.js";
import { deleteFileTool } from "../../src/tools/delete_file.js";
import { moveFileTool } from "../../src/tools/move_file.js";

beforeEach(() => _resetConfirmationTokens());
afterEach(() => _resetConfirmationTokens());

/** A fake Graph whose item lives in `parent` (mutable, to simulate a move between calls). */
function fakeGraph(state: { parent: string }): { calls: { get: number; delete: number; patch: number }; client: Client } {
  const calls = { get: 0, delete: 0, patch: 0 };
  const api = vi.fn().mockImplementation(() => ({
    get: vi.fn().mockImplementation(() => {
      calls.get++;
      return Promise.resolve({ id: "i1", name: "a.txt", size: 1, parentReference: { id: state.parent, path: "/drive/root:/x" } });
    }),
    delete: vi.fn().mockImplementation(() => {
      calls.delete++;
      return Promise.resolve(undefined);
    }),
    patch: vi.fn().mockImplementation(() => {
      calls.patch++;
      return Promise.resolve({ id: "i1", name: "a.txt", parentReference: { id: "dest" } });
    }),
  }));
  return { calls, client: { api } as unknown as Client };
}

const text = (r: { content: unknown[] }): { preview: { confirmation_token: string } } => JSON.parse((r.content[0] as { text: string }).text);

describe("delete_file with expected_parent_id", () => {
  it("previews and deletes when the item is in the expected folder", async () => {
    const g = fakeGraph({ parent: "F1" });
    const preview = text(await deleteFileTool.handler(g.client, { item_id: "i1", expected_parent_id: "F1" })).preview;
    expect(preview.confirmation_token).toMatch(/^[0-9a-f]{32}$/);
    await deleteFileTool.handler(g.client, {
      item_id: "i1",
      expected_parent_id: "F1",
      confirmation_token: preview.confirmation_token,
    });
    expect(g.calls.delete).toBe(1);
  });

  it("refuses at preview when the item is elsewhere, and issues no token", async () => {
    const g = fakeGraph({ parent: "OTHER" });
    await expect(deleteFileTool.handler(g.client, { item_id: "i1", expected_parent_id: "F1" })).rejects.toThrow(
      "not in expected_parent_id F1",
    );
    expect(g.calls.delete).toBe(0);
  });

  it("re-checks at execution: an item moved after the preview is not deleted", async () => {
    const state = { parent: "F1" };
    const g = fakeGraph(state);
    const { confirmation_token } = text(
      await deleteFileTool.handler(g.client, { item_id: "i1", expected_parent_id: "F1" }),
    ).preview;
    state.parent = "OTHER";
    await expect(
      deleteFileTool.handler(g.client, { item_id: "i1", expected_parent_id: "F1", confirmation_token }),
    ).rejects.toThrow("Nothing was changed");
    expect(g.calls.delete).toBe(0);
  });

  it("binds the token to expected_parent_id", async () => {
    const g = fakeGraph({ parent: "F1" });
    const { confirmation_token } = text(
      await deleteFileTool.handler(g.client, { item_id: "i1", expected_parent_id: "F1" }),
    ).preview;
    await expect(deleteFileTool.handler(g.client, { item_id: "i1", confirmation_token })).rejects.toThrow(
      "Refusing to delete",
    );
    expect(g.calls.delete).toBe(0);
  });

  it("without the argument behaves as before", async () => {
    const g = fakeGraph({ parent: "ANY" });
    const { confirmation_token } = text(await deleteFileTool.handler(g.client, { item_id: "i1" })).preview;
    await deleteFileTool.handler(g.client, { item_id: "i1", confirmation_token });
    expect(g.calls.delete).toBe(1);
  });
});

describe("move_file with expected_parent_id", () => {
  it("moves when the item is in the expected folder", async () => {
    const g = fakeGraph({ parent: "F1" });
    await moveFileTool.handler(g.client, { item_id: "i1", target_parent_id: "dest", expected_parent_id: "F1" });
    expect(g.calls.patch).toBe(1);
  });

  it("refuses and changes nothing when the item is elsewhere", async () => {
    const g = fakeGraph({ parent: "OTHER" });
    await expect(
      moveFileTool.handler(g.client, { item_id: "i1", target_parent_id: "dest", expected_parent_id: "F1" }),
    ).rejects.toThrow("Refusing to move");
    expect(g.calls.patch).toBe(0);
  });

  it("without the argument does not read the item first", async () => {
    const g = fakeGraph({ parent: "ANY" });
    await moveFileTool.handler(g.client, { item_id: "i1", target_parent_id: "dest" });
    expect(g.calls.get).toBe(0);
    expect(g.calls.patch).toBe(1);
  });
});
