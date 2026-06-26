/**
 * decisions#210 Layer B — synthetic permission-surface canary.
 *
 * Static invariant over `ALL_TOOLS` (single source of truth from
 * `src/tools/index.ts`). For every exposed tool, exactly one of the
 * following must hold:
 *
 *   1. Scopable — the tool's `inputSchema.properties` declares at
 *      least one parameter the upstream harness pre-tool-use hook
 *      can use to scope a per-agent / per-target deny-list against
 *      (drive_id, site_id, item_id, event_id, meeting_id,
 *      calendar_id, parent_item_id, source_drive_id, target_drive_id,
 *      target_parent_id, transcript_id).
 *
 *   2. NO_TARGET_TOOLS — the tool is on an explicit allowlist of
 *      target-less operations (enumerations / mailbox-wide searches).
 *      The harness cannot scope these by target shape and must apply
 *      a post-call result filter or block the tool outright per
 *      policy. Adding a new entry to NO_TARGET_TOOLS is a deliberate,
 *      reviewable security event.
 *
 *   3. permission_mutating — the tool is classified as
 *      privilege-escalation-class and is gated by CI Layer A in
 *      `.github/workflows/ci.yml`. v0.4.0 baseline: zero such tools.
 *
 * A tool that fits NONE of the three buckets is structurally
 * invisible to the harness deny-list hook — that is the failure mode
 * a future v0.5+ contribution might silently introduce. This test
 * fails the build before such a tool can merge.
 *
 * Runs unconditionally (no tenant calls, no secrets) so PRs from
 * contributors without `VENDOR_SANDBOX_TOKEN` still get the gate.
 */

import { describe, expect, it } from "vitest";

import { ALL_TOOLS } from "../../src/tools/index.js";

/**
 * Parameter names that the upstream harness pre-tool-use hook can
 * read off a `tools/call` request to scope an allow/deny decision
 * against a known resource (driveId, siteId, calendarId, itemId,
 * eventId, meetingId, transcriptId).
 *
 * Order is insignificant; presence of ANY one of these in a tool's
 * `inputSchema.properties` qualifies the tool as scopable.
 */
const SCOPABLE_PARAMS: ReadonlyArray<string> = [
  "drive_id",
  "site_id",
  "item_id",
  "event_id",
  "meeting_id",
  "calendar_id",
  "parent_item_id",
  "source_drive_id",
  "target_drive_id",
  "target_parent_id",
  "transcript_id",
];

/**
 * Tools that intentionally accept no scopable target — enumerations
 * over the caller's full surface or mailbox-wide searches. The
 * harness must apply post-call filtering or deny by tool name; the
 * server cannot scope these by parameter shape.
 *
 * Keep this list MINIMAL. Every entry expands the harness's
 * non-scopable surface and must be justified in code review.
 */
const NO_TARGET_TOOLS: ReadonlySet<string> = new Set([
  "m365-graph:list_drives",
  "m365-graph:list_calendars",
  "m365-graph:search_events",
  "m365-graph:search_events_content",
]);

function isScopable(tool: (typeof ALL_TOOLS)[number]): boolean {
  const props = tool.definition.inputSchema.properties ?? {};
  return SCOPABLE_PARAMS.some((p) => Object.prototype.hasOwnProperty.call(props, p));
}

describe("decisions#210 Layer B — permission-surface invariant", () => {
  it("every tool is either scopable, on NO_TARGET_TOOLS, or classified permission_mutating", () => {
    const unclassified: string[] = [];
    for (const tool of ALL_TOOLS) {
      const name = tool.definition.name;
      const scopable = isScopable(tool);
      const noTarget = NO_TARGET_TOOLS.has(name);
      const permMut = tool.category === "permission_mutating";
      if (!scopable && !noTarget && !permMut) {
        unclassified.push(name);
      }
    }
    expect(
      unclassified,
      `These tools fit none of the three buckets (scopable / NO_TARGET_TOOLS / permission_mutating). ` +
        `They are structurally invisible to the upstream harness deny-list hook. ` +
        `Either add a scopable parameter to the tool's inputSchema, add the tool to ` +
        `NO_TARGET_TOOLS in tests/unit/permission_surface.test.ts (with security-review justification), ` +
        `or classify it as category: "permission_mutating" (which also requires a CI Layer A ` +
        `allowlist entry in .github/workflows/ci.yml). See decisions#210.`,
    ).toEqual([]);
  });

  it("NO_TARGET_TOOLS entries actually exist in ALL_TOOLS (no stale entries)", () => {
    const allNames = new Set(ALL_TOOLS.map((t) => t.definition.name));
    const stale = [...NO_TARGET_TOOLS].filter((n) => !allNames.has(n));
    expect(
      stale,
      `NO_TARGET_TOOLS contains entries that no longer correspond to a registered tool. ` +
        `Remove them from tests/unit/permission_surface.test.ts.`,
    ).toEqual([]);
  });

  it("v0.4.0 baseline — zero tools classified permission_mutating", () => {
    // This assertion is a deliberate floor, not a ceiling. When the
    // first permission_mutating tool is introduced (post-v0.4.0), this
    // expectation flips to `.toBeGreaterThan(0)` AND the workflow's
    // PERMISSION_MUTATING_ALLOWLIST gains the tool's filename in the
    // same PR. Updating this expectation without also updating the
    // workflow allowlist is the regression CI Layer A catches.
    const permMutCount = ALL_TOOLS.filter((t) => t.category === "permission_mutating").length;
    expect(
      permMutCount,
      `v0.4.0 ships with zero permission_mutating tools by design. ` +
        `If you are adding the first one, update this assertion AND add the tool's ` +
        `filename to PERMISSION_MUTATING_ALLOWLIST in .github/workflows/ci.yml.`,
    ).toBe(0);
  });
});
