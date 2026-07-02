/**
 * Server-side state for the spec/approval confirmation-token pattern
 * used by destructive tools (delete_file, cancel_event).
 *
 * Per the handbook MCP server spec § Tool design and ADR 0002:
 *   1. First call: agent submits a spec describing what to delete /
 *      cancel; tool returns a preview + a confirmation_token.
 *   2. Agent reviews the preview, returns a second call with the same
 *      destructive-op args plus the confirmation_token.
 *   3. Tool verifies (token exists, not expired, not used, tied to
 *      THIS tool, args match the original spec) and executes. Token
 *      is then consumed (single-use).
 *
 * Token lifetime: 5 minutes. State lives in a module-level Map keyed
 * by token. Cleared on process exit (per-tenant subprocess per
 * handbook spec — no cross-process leakage). Garbage-collected on
 * each issue/consume pass.
 *
 * Spec match is via SHA-256 of canonical JSON (keys sorted) so the
 * agent can't pass a token issued for {item_id: "A"} together with
 * args {item_id: "B"} and have the destructive call go through.
 */

import crypto from "node:crypto";

const EXPIRY_MS = 5 * 60 * 1000;

interface PendingConfirmation {
  toolName: string;
  specHash: string;
  expiresAt: number;
}

const pending: Map<string, PendingConfirmation> = new Map();

function canonicalize(spec: Record<string, unknown>): string {
  // Stable JSON: keys sorted alphabetically. Sufficient for our specs
  // which are flat objects of primitives (strings, numbers, booleans);
  // we don't need full recursive canonicalization.
  const sortedKeys = Object.keys(spec).sort();
  const sorted: Record<string, unknown> = {};
  for (const k of sortedKeys) sorted[k] = spec[k];
  return JSON.stringify(sorted);
}

function hashSpec(spec: Record<string, unknown>): string {
  return crypto.createHash("sha256").update(canonicalize(spec)).digest("hex");
}

function gc(): void {
  const now = Date.now();
  for (const [k, v] of pending) {
    if (v.expiresAt <= now) pending.delete(k);
  }
}

export interface IssuedToken {
  confirmation_token: string;
  expires_at: string;
  expires_in_seconds: number;
}

export function issueConfirmation(
  toolName: string,
  spec: Record<string, unknown>,
): IssuedToken {
  gc();
  const token = crypto.randomBytes(16).toString("hex");
  const expiresAt = Date.now() + EXPIRY_MS;
  pending.set(token, {
    toolName,
    specHash: hashSpec(spec),
    expiresAt,
  });
  return {
    confirmation_token: token,
    expires_at: new Date(expiresAt).toISOString(),
    expires_in_seconds: Math.floor(EXPIRY_MS / 1000),
  };
}

export type ConsumeError =
  | "token_unknown"
  | "token_expired"
  | "token_wrong_tool"
  | "spec_mismatch";

export type ConsumeResult =
  | { ok: true }
  | { ok: false; error: ConsumeError };

export function consumeConfirmation(
  token: string,
  toolName: string,
  spec: Record<string, unknown>,
): ConsumeResult {
  // Look up the token BEFORE the periodic GC pass so an expired-but-
  // still-recorded entry reports as `token_expired`, not `token_unknown`
  // — the two states are semantically distinct (expired says "you're
  // late"; unknown says "there's no such token"). Only after this
  // lookup do we GC other stale entries. For a security-relevant two-
  // phase gate (delete_file, cancel_event, decline_event) this matters
  // for the caller-facing UX and for audit signal fidelity.
  const entry = pending.get(token);
  if (!entry) {
    gc();
    return { ok: false, error: "token_unknown" };
  }
  if (entry.expiresAt <= Date.now()) {
    pending.delete(token);
    gc();
    return { ok: false, error: "token_expired" };
  }
  gc();
  if (entry.toolName !== toolName) {
    return { ok: false, error: "token_wrong_tool" };
  }
  if (entry.specHash !== hashSpec(spec)) {
    return { ok: false, error: "spec_mismatch" };
  }
  // Single-use: consume on success.
  pending.delete(token);
  return { ok: true };
}

// Test helper — clears the in-memory token store. Not exported as a
// tool. Tests import this directly to ensure isolation.
export function _resetConfirmationTokens(): void {
  pending.clear();
}

// Test helper — inject a pre-expired token so expiry paths can be
// exercised deterministically without racing wall-clock time.
export function _injectExpiredConfirmation(
  token: string,
  toolName: string,
  spec: Record<string, unknown>,
): void {
  pending.set(token, {
    toolName,
    specHash: hashSpec(spec),
    expiresAt: Date.now() - 1_000,
  });
}
