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

import { realpathSync } from "node:fs";
import { pathToFileURL } from "node:url";

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

export const TENANT_ID_RE =
  /^(common|organizations|consumers|[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})$/;

/**
 * Validate the env vars the server needs at startup. Throws with a
 * specific message on failure; main() catches + exits with stderr +
 * code 1.
 *
 * Exported so tests can call directly with a fake env.
 */
export function checkEnv(env: NodeJS.ProcessEnv = process.env): void {
  const missing: string[] = [];
  if (!env.M365_CLIENT_ID) missing.push("M365_CLIENT_ID");
  if (!env.M365_CLIENT_SECRET) missing.push("M365_CLIENT_SECRET");
  if (!env.M365_TENANT_ID) missing.push("M365_TENANT_ID");
  if (missing.length > 0) {
    throw new Error(
      `missing required env var(s): ${missing.join(", ")}. See README.md § Environment variables.`,
    );
  }
  if (!TENANT_ID_RE.test(env.M365_TENANT_ID!)) {
    throw new Error(
      `M365_TENANT_ID has invalid shape: ${env.M365_TENANT_ID}. ` +
        `Expected: 'common' | 'organizations' | 'consumers' | <UUID>.`,
    );
  }
}

/**
 * Execute a single tools/call request against the registered handler
 * map. Wraps handler errors in an `isError: true` MCP response so the
 * agent sees a structured failure instead of a thrown exception that
 * tears down the JSON-RPC session.
 *
 * Exported for unit testing without spinning up the full MCP transport.
 */
export async function dispatchToolCall(
  graph: import("@microsoft/microsoft-graph-client").Client,
  handlers: Map<string, (typeof ALL_TOOLS)[number]["handler"]>,
  request: { params: { name: string; arguments?: unknown } },
): Promise<{ content: Array<{ type: "text"; text: string }>; isError?: boolean }> {
  const { name, arguments: rawArgs } = request.params;
  const handler = handlers.get(name);
  if (!handler) {
    throw new Error(`Unknown tool: ${name}`);
  }
  const args = (rawArgs ?? {}) as Record<string, unknown>;
  try {
    return (await handler(graph, args)) as {
      content: Array<{ type: "text"; text: string }>;
      isError?: boolean;
    };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return {
      content: [{ type: "text", text: `Error: ${message}` }],
      isError: true,
    };
  }
}

async function runMcpServer(): Promise<void> {
  const msal = makeMsalClient();
  const graph = makeGraphClient(msal);
  const handlers = buildHandlerMap(ALL_TOOLS);
  const logLevel = process.env.MCP_SERVER_LOG_LEVEL ?? "info";

  const server = new Server(
    {
      name: "@juvantlabs/m365-graph-mcp-server",
      version: "0.1.1",
    },
    {
      capabilities: { tools: {} },
    },
  );

  server.setRequestHandler(ListToolsRequestSchema, async () => ({
    tools: ALL_TOOLS.map((t) => t.definition),
  }));

  server.setRequestHandler(CallToolRequestSchema, async (request) =>
    dispatchToolCall(graph, handlers, request),
  );

  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error(
    `[m365-graph-mcp-server] running on stdio (log level: ${logLevel}, tenant: ${process.env.M365_TENANT_ID}, tools: ${ALL_TOOLS.length})`,
  );
}

/**
 * Subcommand dispatcher. Exported so tests can verify routing without
 * actually spawning the server.
 */
export async function dispatch(argv: string[], handlers: {
  setup: () => Promise<void>;
  serve: () => Promise<void>;
}): Promise<void> {
  const subcommand = argv[2];
  if (subcommand === "setup") {
    await handlers.setup();
    return;
  }
  await handlers.serve();
}

async function main(): Promise<void> {
  try {
    checkEnv();
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error(`[m365-graph-mcp-server] ${message}`);
    process.exit(1);
  }

  await dispatch(process.argv, {
    setup: runSetup,
    serve: runMcpServer,
  });
}

/**
 * Detect whether this module is the entry point (vs imported by tests
 * or other code). When invoked via the `bin` symlink (npm puts the
 * package's bin into `node_modules/.bin/<name>` as a symlink to
 * `dist/index.js`), `process.argv[1]` is the symlink path, not the
 * underlying file. Resolve symlinks via `realpathSync` so the
 * comparison against `import.meta.url` matches in both cases:
 *   - Direct invocation: `node dist/index.js`
 *   - Bin invocation:    `npx @juvantlabs/m365-graph-mcp-server`
 *   - tsx (dev):         `tsx src/index.ts`
 *
 * Also resolves macOS's `/tmp` → `/private/tmp` symlink.
 */
function isInvokedAsMain(): boolean {
  if (!process.argv[1]) return false;
  try {
    const realPath = realpathSync(process.argv[1]);
    return import.meta.url === pathToFileURL(realPath).href;
  } catch {
    return false;
  }
}

if (isInvokedAsMain()) {
  main().catch((err) => {
    console.error("[m365-graph-mcp-server] fatal:", err);
    process.exit(1);
  });
}
