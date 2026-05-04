import { afterEach, beforeEach, describe, expect, it } from "vitest";

import {
  _resetConfirmationTokens,
  consumeConfirmation,
  issueConfirmation,
} from "../../src/auth/confirmation_tokens.js";

beforeEach(() => {
  _resetConfirmationTokens();
});

afterEach(() => {
  _resetConfirmationTokens();
});

describe("issueConfirmation", () => {
  it("returns a token, expires_at, and expires_in_seconds", () => {
    const r = issueConfirmation("tool:foo", { item_id: "x" });
    expect(typeof r.confirmation_token).toBe("string");
    expect(r.confirmation_token.length).toBeGreaterThanOrEqual(32);
    expect(r.expires_in_seconds).toBe(300);
    expect(new Date(r.expires_at).getTime()).toBeGreaterThan(Date.now());
  });

  it("issues a different token each call (random)", () => {
    const a = issueConfirmation("t", { id: 1 });
    const b = issueConfirmation("t", { id: 1 });
    expect(a.confirmation_token).not.toBe(b.confirmation_token);
  });
});

describe("consumeConfirmation", () => {
  it("succeeds with the right token + tool + matching spec", () => {
    const { confirmation_token } = issueConfirmation("tool:foo", { item_id: "x" });
    expect(consumeConfirmation(confirmation_token, "tool:foo", { item_id: "x" })).toEqual({ ok: true });
  });

  it("rejects unknown token", () => {
    expect(consumeConfirmation("nope", "tool:foo", { item_id: "x" })).toEqual({
      ok: false,
      error: "token_unknown",
    });
  });

  it("rejects token issued for a different tool", () => {
    const { confirmation_token } = issueConfirmation("tool:foo", { item_id: "x" });
    expect(consumeConfirmation(confirmation_token, "tool:bar", { item_id: "x" })).toEqual({
      ok: false,
      error: "token_wrong_tool",
    });
  });

  it("rejects token paired with mismatched spec (different item_id)", () => {
    const { confirmation_token } = issueConfirmation("tool:foo", { item_id: "A" });
    expect(consumeConfirmation(confirmation_token, "tool:foo", { item_id: "B" })).toEqual({
      ok: false,
      error: "spec_mismatch",
    });
  });

  it("treats key reordering as the same spec (canonical hash)", () => {
    const { confirmation_token } = issueConfirmation("t", { a: "1", b: "2" });
    expect(consumeConfirmation(confirmation_token, "t", { b: "2", a: "1" })).toEqual({ ok: true });
  });

  it("is single-use (second consume returns token_unknown)", () => {
    const { confirmation_token } = issueConfirmation("t", { x: 1 });
    expect(consumeConfirmation(confirmation_token, "t", { x: 1 })).toEqual({ ok: true });
    expect(consumeConfirmation(confirmation_token, "t", { x: 1 })).toEqual({
      ok: false,
      error: "token_unknown",
    });
  });
});
