import type { TokenCacheContext } from "@azure/msal-node";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const { entryConstructor, setPasswordMock, getPasswordMock } = vi.hoisted(() => ({
  setPasswordMock: vi.fn(),
  getPasswordMock: vi.fn(),
  entryConstructor: vi.fn(),
}));

vi.mock("@napi-rs/keyring", () => ({
  Entry: entryConstructor,
}));

import { makeCachePlugin } from "../../src/auth/msal.js";

beforeEach(() => {
  entryConstructor.mockReset();
  entryConstructor.mockImplementation(() => ({
    setPassword: setPasswordMock,
    getPassword: getPasswordMock,
    deletePassword: vi.fn(),
  }));
  setPasswordMock.mockReset();
  getPasswordMock.mockReset();
});

afterEach(() => {
  vi.clearAllMocks();
});

function makeContext(opts: { hasChanged: boolean }): {
  ctx: TokenCacheContext;
  deserialize: ReturnType<typeof vi.fn>;
  serialize: ReturnType<typeof vi.fn>;
} {
  const deserialize = vi.fn();
  const serialize = vi.fn().mockReturnValue("serialized-cache-blob");
  const ctx = {
    cacheHasChanged: opts.hasChanged,
    tokenCache: { deserialize, serialize },
  } as unknown as TokenCacheContext;
  return { ctx, deserialize, serialize };
}

describe("makeCachePlugin — beforeCacheAccess (load)", () => {
  it("deserializes the cache from the keychain when entry exists", async () => {
    getPasswordMock.mockReturnValue("cached-blob-from-keychain");
    const plugin = makeCachePlugin("tenant-1");
    const { ctx, deserialize } = makeContext({ hasChanged: false });
    await plugin.beforeCacheAccess!(ctx);
    expect(deserialize).toHaveBeenCalledWith("cached-blob-from-keychain");
  });

  it("does not deserialize when the keychain has no entry yet", async () => {
    getPasswordMock.mockImplementation(() => {
      throw new Error("no entry");
    });
    const plugin = makeCachePlugin("tenant-1");
    const { ctx, deserialize } = makeContext({ hasChanged: false });
    await plugin.beforeCacheAccess!(ctx);
    expect(deserialize).not.toHaveBeenCalled();
  });
});

describe("makeCachePlugin — afterCacheAccess (save)", () => {
  it("serializes + saves the cache when MSAL marks it changed", async () => {
    const plugin = makeCachePlugin("tenant-1");
    const { ctx, serialize } = makeContext({ hasChanged: true });
    await plugin.afterCacheAccess!(ctx);
    expect(serialize).toHaveBeenCalled();
    expect(setPasswordMock).toHaveBeenCalledWith("serialized-cache-blob");
  });

  it("does NOT save when the cache hasn't changed", async () => {
    const plugin = makeCachePlugin("tenant-1");
    const { ctx, serialize } = makeContext({ hasChanged: false });
    await plugin.afterCacheAccess!(ctx);
    expect(serialize).not.toHaveBeenCalled();
    expect(setPasswordMock).not.toHaveBeenCalled();
  });
});
