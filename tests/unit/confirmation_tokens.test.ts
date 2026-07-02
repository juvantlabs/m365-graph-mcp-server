import { afterEach, beforeEach, describe, expect, it } from "vitest";

import {
  _injectExpiredConfirmation,
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

  it("rejects an expired token with token_expired (distinct from token_unknown)", () => {
    // Regression: pre-fix, gc() ran before the entry lookup so an expired-
    // but-still-recorded entry was deleted first and the caller saw
    // token_unknown. The two states are semantically distinct — expired
    // says "you're late", unknown says "there's no such token" — and both
    // are audit-relevant on a security-relevant two-phase gate
    // (delete_file, cancel_event, decline_event).
    _injectExpiredConfirmation("expired-token", "tool:foo", { item_id: "x" });
    expect(consumeConfirmation("expired-token", "tool:foo", { item_id: "x" })).toEqual({
      ok: false,
      error: "token_expired",
    });
    // Expired token is dropped from the store on consume; a second attempt
    // now falls through to token_unknown — proving the two outcomes are
    // reachable independently.
    expect(consumeConfirmation("expired-token", "tool:foo", { item_id: "x" })).toEqual({
      ok: false,
      error: "token_unknown",
    });
  });

  it("token_unknown is returned for a token that was never issued (distinct from token_expired)", () => {
    // Companion to the expired-token test: with no matching entry at all,
    // the result must be token_unknown, never token_expired.
    expect(consumeConfirmation("never-issued", "tool:foo", { item_id: "x" })).toEqual({
      ok: false,
      error: "token_unknown",
    });
  });
});
