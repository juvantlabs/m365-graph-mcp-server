import { describe, expect, it, vi } from "vitest";

vi.mock("../../src/auth/msal.js", () => ({
  getAccessToken: vi.fn().mockResolvedValue("test-bearer-token"),
}));

import { getAccessToken } from "../../src/auth/msal.js";
import { MsalAuthProvider, makeGraphClient } from "../../src/client/graph.js";

describe("MsalAuthProvider", () => {
  it("delegates getAccessToken to the MSAL helper", async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const fakeMsal = { _kind: "fake-msal" } as any;
    const provider = new MsalAuthProvider(fakeMsal);
    const token = await provider.getAccessToken();
    expect(token).toBe("test-bearer-token");
    expect(getAccessToken).toHaveBeenCalledWith(fakeMsal);
  });

  it("ignores the AuthenticationProviderOptions arg (we never need scope hints)", async () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const provider = new MsalAuthProvider({} as any);
    const token = await provider.getAccessToken({ scopes: ["irrelevant"] });
    expect(token).toBe("test-bearer-token");
  });
});

describe("makeGraphClient", () => {
  it("returns a Client instance with an MSAL-backed auth provider", () => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const fakeMsal = { _kind: "fake-msal" } as any;
    const client = makeGraphClient(fakeMsal);
    expect(typeof client.api).toBe("function");
  });
});
