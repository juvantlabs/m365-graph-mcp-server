import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("@napi-rs/keyring", () => ({
  // Vitest 4: constructor mocks need a `function`/class implementation.
  Entry: vi.fn().mockImplementation(function () {
    return {
      getPassword: (): string => {
        throw new Error("no entry");
      },
      setPassword: vi.fn(),
      deletePassword: vi.fn(),
    };
  }),
}));

import type { PublicClientApplication } from "@azure/msal-node";

import { DELEGATED_SCOPES } from "../../src/auth/msal.js";
import { browserCommand, parseSetupArgs, runSetup } from "../../src/auth/setup.js";

const account = { username: "user@example.test", homeAccountId: "oid.tid" };

function fakeClient(): {
  client: PublicClientApplication;
  acquireTokenInteractive: ReturnType<typeof vi.fn>;
  acquireTokenByDeviceCode: ReturnType<typeof vi.fn>;
} {
  const acquireTokenInteractive = vi.fn().mockResolvedValue({ account });
  const acquireTokenByDeviceCode = vi.fn().mockImplementation(async (req) => {
    req.deviceCodeCallback({ message: "Enter code FAKE-CODE" });
    return { account };
  });
  return {
    client: { acquireTokenInteractive, acquireTokenByDeviceCode } as unknown as PublicClientApplication,
    acquireTokenInteractive,
    acquireTokenByDeviceCode,
  };
}

let stderr: ReturnType<typeof vi.spyOn>;

beforeEach(() => {
  stderr = vi.spyOn(console, "error").mockImplementation(() => {});
});

afterEach(() => {
  stderr.mockRestore();
});

describe("parseSetupArgs", () => {
  it("defaults to the interactive browser flow", () => {
    expect(parseSetupArgs([])).toEqual({ deviceCode: false });
  });

  it("selects device code with --device-code", () => {
    expect(parseSetupArgs(["--device-code"])).toEqual({ deviceCode: true });
  });

  it("rejects unknown options instead of silently ignoring them", () => {
    expect(() => parseSetupArgs(["--client-secret", "x"])).toThrow(/unknown setup option/);
  });
});

describe("browserCommand", () => {
  const url = "https://login.microsoftonline.com/t/oauth2/v2.0/authorize?a=1&b=2";

  it.each([
    ["darwin", "open"],
    ["linux", "xdg-open"],
    ["win32", "rundll32"],
  ] as const)("on %s uses %s with the URL as a single argument (no shell)", (platform, file) => {
    const cmd = browserCommand(url, platform);
    expect(cmd.file).toBe(file);
    expect(cmd.args[cmd.args.length - 1]).toBe(url);
  });
});

describe("runSetup", () => {
  it("runs the interactive PKCE flow by default with the delegated scopes", async () => {
    const fake = fakeClient();
    const openBrowser = vi.fn();
    const result = await runSetup({ deviceCode: false }, { client: fake.client, openBrowser });

    expect(result.account).toEqual(account);
    expect(fake.acquireTokenInteractive).toHaveBeenCalledTimes(1);
    expect(fake.acquireTokenByDeviceCode).not.toHaveBeenCalled();
    const req = fake.acquireTokenInteractive.mock.calls[0][0];
    expect(req.scopes).toEqual(DELEGATED_SCOPES);
    expect(req.openBrowser).toBe(openBrowser);
  });

  it("runs the device code flow with --device-code and prints the user message on stderr", async () => {
    const fake = fakeClient();
    await runSetup({ deviceCode: true }, { client: fake.client });

    expect(fake.acquireTokenByDeviceCode).toHaveBeenCalledTimes(1);
    expect(fake.acquireTokenInteractive).not.toHaveBeenCalled();
    expect(fake.acquireTokenByDeviceCode.mock.calls[0][0].scopes).toEqual(DELEGATED_SCOPES);
    expect(stderr).toHaveBeenCalledWith(expect.stringContaining("FAKE-CODE"));
  });

  it("throws when MSAL returns no account", async () => {
    const fake = fakeClient();
    fake.acquireTokenByDeviceCode.mockResolvedValue(null);
    await expect(runSetup({ deviceCode: true }, { client: fake.client })).rejects.toThrow(
      /no account/,
    );
  });

  it("writes nothing to stdout (MCP stdio framing)", async () => {
    const stdout = vi.spyOn(process.stdout, "write");
    const log = vi.spyOn(console, "log");
    try {
      await runSetup({ deviceCode: true }, { client: fakeClient().client });
      expect(stdout).not.toHaveBeenCalled();
      expect(log).not.toHaveBeenCalled();
    } finally {
      stdout.mockRestore();
      log.mockRestore();
    }
  });
});
