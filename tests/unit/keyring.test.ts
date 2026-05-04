import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

// `vi.mock()` is hoisted above all imports — the factory cannot
// reference outer variables. Use `vi.hoisted()` to declare the mocks
// in a hoisted block too, so they're available when the factory runs.
const { entryConstructor, setPasswordMock, getPasswordMock, deletePasswordMock } = vi.hoisted(() => ({
  setPasswordMock: vi.fn(),
  getPasswordMock: vi.fn(),
  deletePasswordMock: vi.fn(),
  entryConstructor: vi.fn(),
}));

vi.mock("@napi-rs/keyring", () => ({
  Entry: entryConstructor,
}));

import { getTokenStore } from "../../src/auth/keyring.js";

beforeEach(() => {
  entryConstructor.mockReset();
  entryConstructor.mockImplementation((service: string, account: string) => ({
    service,
    account,
    setPassword: setPasswordMock,
    getPassword: getPasswordMock,
    deletePassword: deletePasswordMock,
  }));
  setPasswordMock.mockReset();
  getPasswordMock.mockReset();
  deletePasswordMock.mockReset();
});

afterEach(() => {
  vi.clearAllMocks();
});

describe("getTokenStore", () => {
  it("returns null on load() when getPassword throws (no entry yet)", () => {
    getPasswordMock.mockImplementation(() => {
      throw new Error("entry not found");
    });
    const store = getTokenStore("tenant-1");
    expect(store.load()).toBeNull();
  });

  it("returns the cached string on load() when getPassword succeeds", () => {
    getPasswordMock.mockReturnValue("cached-msal-blob");
    const store = getTokenStore("tenant-1");
    expect(store.load()).toBe("cached-msal-blob");
  });

  it("calls setPassword with the serialized cache on save()", () => {
    const store = getTokenStore("tenant-1");
    store.save("new-cache-blob");
    expect(setPasswordMock).toHaveBeenCalledWith("new-cache-blob");
  });

  it("calls deletePassword on clear()", () => {
    const store = getTokenStore("tenant-1");
    store.clear();
    expect(deletePasswordMock).toHaveBeenCalled();
  });

  it("clear() ignores deletePassword errors (idempotent)", () => {
    deletePasswordMock.mockImplementation(() => {
      throw new Error("nothing to delete");
    });
    const store = getTokenStore("tenant-1");
    expect(() => store.clear()).not.toThrow();
  });
});

describe("getTokenStore — per-tenant scoping", () => {
  it("creates separate Entry instances per tenant id", () => {
    // Per-tenant scoping prevents cross-tenant token leakage.
    entryConstructor.mockClear();

    getTokenStore("tenant-A");
    getTokenStore("tenant-B");

    expect(entryConstructor).toHaveBeenCalledTimes(2);
    expect(entryConstructor).toHaveBeenNthCalledWith(1, expect.any(String), "tenant:tenant-A");
    expect(entryConstructor).toHaveBeenNthCalledWith(2, expect.any(String), "tenant:tenant-B");
  });
});
