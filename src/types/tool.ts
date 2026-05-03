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
  definition: ToolDefinition;
  handler: ToolHandler;
}
