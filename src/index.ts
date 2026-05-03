#!/usr/bin/env node
/**
 * M365 Graph MCP Server — entrypoint.
 *
 * Two modes, dispatched on argv[2]:
 *   - `setup`: interactive OAuth flow → caches tokens in OS keychain
 *   - (default): stdio MCP server → reads cached tokens, serves tool
 *     calls over JSON-RPC framing
 *
 * Conforms to handbook docs/repo-types/mcp-server.md.
 *
 * IMPORTANT: All non-protocol output must go to stderr (`console.error`).
 * Writing to stdout corrupts the MCP stdio JSON-RPC framing. CI lint
 * rules enforce this in `juvantlabs/*-mcp-server` repos.
 */

import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";

import { makeMsalClient } from "./auth/msal.js";
import { runSetup } from "./auth/setup.js";
import { makeGraphClient } from "./client/graph.js";
import { ALL_TOOLS, buildHandlerMap } from "./tools/index.js";

const LOG_LEVEL = process.env.MCP_SERVER_LOG_LEVEL ?? "info";

const M365_CLIENT_ID = process.env.M365_CLIENT_ID;
const M365_CLIENT_SECRET = process.env.M365_CLIENT_SECRET;
const M365_TENANT_ID = process.env.M365_TENANT_ID;

const TENANT_ID_RE =
  /^(common|organizations|consumers|[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})$/;

function validateEnv(): void {
  const missing: string[] = [];
  if (!M365_CLIENT_ID) missing.push("M365_CLIENT_ID");
  if (!M365_CLIENT_SECRET) missing.push("M365_CLIENT_SECRET");
  if (!M365_TENANT_ID) missing.push("M365_TENANT_ID");
  if (missing.length > 0) {
    console.error(
      `[m365-graph-mcp-server] missing required env var(s): ${missing.join(", ")}`,
    );
    console.error("[m365-graph-mcp-server] see README.md § Environment variables");
    process.exit(1);
  }
  if (!TENANT_ID_RE.test(M365_TENANT_ID!)) {
    console.error(
      `[m365-graph-mcp-server] M365_TENANT_ID has invalid shape: ${M365_TENANT_ID}`,
    );
    console.error(
      "[m365-graph-mcp-server] expected: 'common' | 'organizations' | 'consumers' | <UUID>",
    );
    process.exit(1);
  }
}

async function runMcpServer(): Promise<void> {
  const msal = makeMsalClient();
  const graph = makeGraphClient(msal);
  const handlers = buildHandlerMap(ALL_TOOLS);

  const server = new Server(
    {
      name: "@juvantlabs/m365-graph-mcp-server",
      version: "0.0.1",
    },
    {
      capabilities: { tools: {} },
    },
  );

  server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: ALL_TOOLS.map((t) => t.definition),
  }));

  server.setRequestHandler(CallToolRequestSchema, async (request) => {
    const { name, arguments: rawArgs } = request.params;
    const handler = handlers.get(name);
    if (!handler) {
      throw new Error(`Unknown tool: ${name}`);
    }
    const args = (rawArgs ?? {}) as Record<string, unknown>;
    try {
      return await handler(graph, args);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return {
        content: [{ type: "text", text: `Error: ${message}` }],
        isError: true,
      };
    }
  });

  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error(
    `[m365-graph-mcp-server] running on stdio (log level: ${LOG_LEVEL}, tenant: ${M365_TENANT_ID}, tools: ${ALL_TOOLS.length})`,
  );
}

async function main(): Promise<void> {
  validateEnv();

  const subcommand = process.argv[2];
  if (subcommand === "setup") {
    await runSetup();
    return;
  }
  await runMcpServer();
}

main().catch((err) => {
  console.error("[m365-graph-mcp-server] fatal:", err);
  process.exit(1);
});
