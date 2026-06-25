import { afterEach, describe, expect, it, vi } from "vitest";

const { entryConstructor } = vi.hoisted(() => ({
  entryConstructor: vi.fn().mockImplementation(() => ({
    setPassword: vi.fn(),
    getPassword: vi.fn().mockImplementation(() => {
      throw new Error("no entry");
    }),
    deletePassword: vi.fn(),
  })),
}));

vi.mock("@napi-rs/keyring", () => ({
  Entry: entryConstructor,
}));

import { DELEGATED_SCOPES, REDIRECT_URI, getAccessToken, makeMsalClient } from "../../src/auth/msal.js";

afterEach(() => {
  vi.clearAllMocks();
});

describe("makeMsalClient", () => {
  it("constructs a ConfidentialClientApplication with env-derived config", () => {
    const original = { ...process.env };
    process.env.M365_TENANT_ID = "c557607d-995c-4eb7-967b-50c6361fbad9";
    process.env.M365_CLIENT_ID = "00000000-0000-0000-0000-000000000001";
    process.env.M365_CLIENT_SECRET = "test-secret";
    try {
      const client = makeMsalClient();
      expect(client).toBeDefined();
      expect(typeof client.acquireTokenSilent).toBe("function");
      expect(typeof client.getAuthCodeUrl).toBe("function");
    } finally {
      process.env = original;
    }
  });
});

describe("REDIRECT_URI", () => {
  it("matches the localhost callback registered in the Entra app", () => {
    expect(REDIRECT_URI).toBe("http://localhost:3000/auth/callback");
  });
});

describe("DELEGATED_SCOPES", () => {
  it("requests Files.ReadWrite (subsumes .Read)", () => {
    expect(DELEGATED_SCOPES).toContain("Files.ReadWrite");
    expect(DELEGATED_SCOPES).not.toContain("Files.Read"); // narrow scope no longer requested
  });

  it("requests Calendars.ReadWrite (subsumes .Read)", () => {
    expect(DELEGATED_SCOPES).toContain("Calendars.ReadWrite");
    expect(DELEGATED_SCOPES).not.toContain("Calendars.Read");
  });

  it("includes offline_access for refresh tokens", () => {
    expect(DELEGATED_SCOPES).toContain("offline_access");
  });

  it("requests Sites.ReadWrite.All for SharePoint document libraries (v0.4.0)", () => {
    expect(DELEGATED_SCOPES).toContain("Sites.ReadWrite.All");
  });

  it("does NOT request permission-mutation-class scopes (decisions#210)", () => {
    // These scopes carry sharing-link / invite / permission-grant /
    // ownership-transfer privileges and are deliberately excluded from
    // the v0.4.0 baseline. Introducing any of them requires a
    // permission_mutating-classified tool AND a CI Layer A allowlist
    // entry — see .github/workflows/ci.yml.
    expect(DELEGATED_SCOPES).not.toContain("Sites.Manage.All");
    expect(DELEGATED_SCOPES).not.toContain("Sites.FullControl.All");
    expect(DELEGATED_SCOPES).not.toContain("Application.ReadWrite.All");
  });
});

describe("getAccessToken", () => {
  function makeMsalMock(opts: {
    accounts: Array<{ username: string }>;
    silentResult: { accessToken: string } | null;
    silentThrows?: Error;
  }): {
    getTokenCache: () => { getAllAccounts: () => Promise<unknown[]> };
    acquireTokenSilent: ReturnType<typeof vi.fn>;
  } {
    const acquireTokenSilent = vi.fn();
    if (opts.silentThrows) {
      acquireTokenSilent.mockRejectedValue(opts.silentThrows);
    } else {
      acquireTokenSilent.mockResolvedValue(opts.silentResult);
    }
    return {
      getTokenCache: () => ({
        getAllAccounts: vi.fn().mockResolvedValue(opts.accounts),
      }),
      acquireTokenSilent,
    };
  }

  it("throws when no cached account exists (setup not run)", async () => {
    const msal = makeMsalMock({ accounts: [], silentResult: null });
    await expect(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      getAccessToken(msal as any),
    ).rejects.toThrow(/No cached account/);
  });

  it("returns the access token from acquireTokenSilent", async () => {
    const msal = makeMsalMock({
      accounts: [{ username: "alice@x.com" }],
      silentResult: { accessToken: "the-token" },
    });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const tok = await getAccessToken(msal as any);
    expect(tok).toBe("the-token");
  });

  it("throws if silent result has no accessToken (refresh failed)", async () => {
    const msal = makeMsalMock({
      accounts: [{ username: "alice@x.com" }],
      silentResult: null,
    });
    await expect(
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      getAccessToken(msal as any),
    ).rejects.toThrow(/Silent token acquisition/);
  });

  it("requests the configured DELEGATED_SCOPES", async () => {
    const msal = makeMsalMock({
      accounts: [{ username: "alice@x.com" }],
      silentResult: { accessToken: "tok" },
    });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    await getAccessToken(msal as any);
    expect(msal.acquireTokenSilent).toHaveBeenCalledWith(
      expect.objectContaining({ scopes: DELEGATED_SCOPES }),
    );
  });
});
