/**
 * Shared types for MCP tool definitions in this server.
 *
 * Each tool exports a `definition` (the static metadata returned by
 * tools/list) and a `handler` (the function invoked by tools/call).
 * The dispatcher in src/index.ts switches on the tool name to call
 * the right handler.
 */

import type { Client } from "@microsoft/microsoft-graph-client";
import type { CallToolResult } from "@modelcontextprotocol/sdk/types.js";

export interface ToolDefinition {
  name: string;
  description: string;
  inputSchema: {
    type: "object";
    properties: Record<string, unknown>;
    required: string[];
  };
}

/**
 * Tool categorization per handbook ADR 0004 (Agent action guardrails).
 *
 * - `read` — no mutations.
 * - `write_idempotent` — mutations where re-running with the same args
 *   reaches the same effective state (or is reversible by another
 *   tool of this server within the same session, e.g. create_event
 *   reversible by cancel_event).
 * - `write_irreversible` — external mutation that cannot be reverted by
 *   a subsequent automated call from this server (notifications sent
 *   to recipients, deletions without a deterministic restore path,
 *   payments). MUST gate via the two-phase confirmation token pattern
 *   (preview → token → execute) per ADR 0002. CI enforces this.
 * - `permission_mutating` — operations that create/modify standing
 *   access, sharing, or ownership of a resource (sharing-link creation,
 *   guest invite, permission grant/revoke, sensitivity-label change,
 *   ownership transfer). Distinct from `write_irreversible` because
 *   the threat model is privilege escalation: a single call grants a
 *   third party persistent access that survives the session and
 *   bypasses content-level RW deny-lists. No tool in v0.4.0 carries
 *   this category — it is a pre-classification slot. CI Layer A
 *   (`.github/workflows/ci.yml`) fails the build if any future tool
 *   touches a permission/sharing Graph endpoint without being
 *   classified here AND added to the workflow's allowlist.
 *
 * Ambiguous categorizations resolve to the strictest. When in doubt,
 * gate it.
 */
export type ToolCategory =
  | "read"
  | "write_idempotent"
  | "write_irreversible"
  | "permission_mutating";

/**
 * Return shape for a tool handler. We use the SDK's `CallToolResult`
 * directly so the dispatcher in `src/index.ts` can return handler
 * results to `setRequestHandler(CallToolRequestSchema, …)` without a
 * type assertion.
 */
export type ToolResponse = CallToolResult;

export type ToolHandler = (
  graph: Client,
  args: Record<string, unknown>,
) => Promise<ToolResponse>;

export interface Tool {
  category: ToolCategory;
  definition: ToolDefinition;
  handler: ToolHandler;
}
