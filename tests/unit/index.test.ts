import { createRequire } from "node:module";

import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import {
  PACKAGE_VERSION,
  TENANT_ID_RE,
  checkEnv,
  dispatch,
  dispatchToolCall,
  readPackageVersion,
} from "../../src/index.js";

const require = createRequire(import.meta.url);
const pkgJson = require("../../package.json") as { version: string };

describe("checkEnv", () => {
  const valid = {
    M365_CLIENT_ID: "00000000-0000-0000-0000-000000000001",
    M365_CLIENT_SECRET: "secret-value",
    M365_TENANT_ID: "c557607d-995c-4eb7-967b-50c6361fbad9",
  };

  it("passes with all 3 required vars + valid tenant UUID", () => {
    expect(() => checkEnv(valid)).not.toThrow();
  });

  it.each(["common", "organizations", "consumers"])(
    "accepts well-known tenant alias '%s'",
    (alias) => {
      expect(() => checkEnv({ ...valid, M365_TENANT_ID: alias })).not.toThrow();
    },
  );

  it("throws with a list of missing vars", () => {
    expect(() => checkEnv({ M365_TENANT_ID: valid.M365_TENANT_ID })).toThrow(
      /M365_CLIENT_ID, M365_CLIENT_SECRET/,
    );
  });

  it("throws on invalid tenant ID shape", () => {
    expect(() => checkEnv({ ...valid, M365_TENANT_ID: "not-a-uuid" })).toThrow(
      /invalid shape/,
    );
  });

  it("throws on empty string tenant ID (treated as missing)", () => {
    expect(() => checkEnv({ ...valid, M365_TENANT_ID: "" })).toThrow(/missing/);
  });
});

describe("TENANT_ID_RE", () => {
  it("matches a UUID with all hex chars", () => {
    expect(TENANT_ID_RE.test("12345678-1234-1234-1234-123456789abc")).toBe(true);
  });

  it("matches well-known aliases", () => {
    for (const alias of ["common", "organizations", "consumers"]) {
      expect(TENANT_ID_RE.test(alias)).toBe(true);
    }
  });

  it("rejects uppercase hex (Graph normalizes to lowercase)", () => {
    expect(TENANT_ID_RE.test("12345678-1234-1234-1234-123456789ABC")).toBe(false);
  });

  it("rejects bare strings + obvious garbage", () => {
    for (const bad of ["", "not-a-tenant", "12345", "https://example.com"]) {
      expect(TENANT_ID_RE.test(bad)).toBe(false);
    }
  });
});

describe("dispatchToolCall", () => {
  const fakeGraph = {} as Client;
  const okResponse = { content: [{ type: "text" as const, text: "ok" }] };

  it("invokes the registered handler for a known tool", async () => {
    const handler = vi.fn().mockResolvedValue(okResponse);
    const handlers = new Map([["m365-graph:foo", handler]]);
    const result = await dispatchToolCall(fakeGraph, handlers, {
      params: { name: "m365-graph:foo", arguments: { x: 1 } },
    });
    expect(result).toEqual(okResponse);
    expect(handler).toHaveBeenCalledWith(fakeGraph, { x: 1 });
  });

  it("throws on unknown tool name", async () => {
    const handlers = new Map();
    await expect(
      dispatchToolCall(fakeGraph, handlers, {
        params: { name: "m365-graph:nope", arguments: {} },
      }),
    ).rejects.toThrow("Unknown tool: m365-graph:nope");
  });

  it("wraps handler exceptions as isError MCP responses", async () => {
    const handler = vi.fn().mockRejectedValue(new Error("graph API exploded"));
    const handlers = new Map([["m365-graph:foo", handler]]);
    const result = await dispatchToolCall(fakeGraph, handlers, {
      params: { name: "m365-graph:foo" },
    });
    expect(result.isError).toBe(true);
    expect((result.content[0] as { text: string }).text).toContain("graph API exploded");
  });

  it("defaults missing arguments to an empty object", async () => {
    const handler = vi.fn().mockResolvedValue(okResponse);
    const handlers = new Map([["m365-graph:foo", handler]]);
    await dispatchToolCall(fakeGraph, handlers, { params: { name: "m365-graph:foo" } });
    expect(handler).toHaveBeenCalledWith(fakeGraph, {});
  });

  it("stringifies non-Error throws into the wrapped response", async () => {
    const handler = vi.fn().mockRejectedValue("plain string thrown");
    const handlers = new Map([["m365-graph:foo", handler]]);
    const result = await dispatchToolCall(fakeGraph, handlers, {
      params: { name: "m365-graph:foo" },
    });
    expect(result.isError).toBe(true);
    expect((result.content[0] as { text: string }).text).toContain("plain string thrown");
  });
});

describe("dispatch", () => {
  it("routes 'setup' subcommand to handlers.setup", async () => {
    const setup = vi.fn().mockResolvedValue(undefined);
    const serve = vi.fn().mockResolvedValue(undefined);
    await dispatch(["node", "index.ts", "setup"], { setup, serve });
    expect(setup).toHaveBeenCalledTimes(1);
    expect(serve).not.toHaveBeenCalled();
  });

  it("routes default (no subcommand) to handlers.serve", async () => {
    const setup = vi.fn().mockResolvedValue(undefined);
    const serve = vi.fn().mockResolvedValue(undefined);
    await dispatch(["node", "index.ts"], { setup, serve });
    expect(setup).not.toHaveBeenCalled();
    expect(serve).toHaveBeenCalledTimes(1);
  });

  it("routes unknown subcommand to handlers.serve (graceful fallback)", async () => {
    const setup = vi.fn().mockResolvedValue(undefined);
    const serve = vi.fn().mockResolvedValue(undefined);
    await dispatch(["node", "index.ts", "unknown"], { setup, serve });
    expect(serve).toHaveBeenCalledTimes(1);
  });

  it("propagates errors from the chosen handler", async () => {
    const setup = vi.fn().mockRejectedValue(new Error("boom"));
    const serve = vi.fn();
    await expect(dispatch(["node", "x", "setup"], { setup, serve })).rejects.toThrow("boom");
  });
});

describe("package version", () => {
  it("PACKAGE_VERSION strictly equals package.json version (no drift)", () => {
    expect(pkgJson.version).toMatch(/^\d+\.\d+\.\d+(?:-[\w.+-]+)?$/);
    expect(PACKAGE_VERSION).toBe(pkgJson.version);
  });

  it("readPackageVersion() resolves the same value as the cached constant", () => {
    expect(readPackageVersion()).toBe(pkgJson.version);
    expect(readPackageVersion()).toBe(PACKAGE_VERSION);
  });
});
