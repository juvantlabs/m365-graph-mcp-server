/**
 * Contract tests against the REAL @azure/msal-node (no MSAL mocks).
 *
 * Only the edges are faked: the OS keychain (in-memory map behind the
 * real cache plugin) and the network (offline authority metadata + a
 * fake token endpoint). Everything in between — request building, the
 * auth-code exchange, ID-token handling, cache (de)serialization, silent
 * acquisition and refresh — is the shipped MSAL code.
 *
 * These pin the behaviours an msal-node upgrade can silently break:
 *   1. setup's interactive flow as a PUBLIC client: authorization URL
 *      with a loopback redirect on a free port, PKCE S256,
 *      response_mode=form_post and no nonce; code redemption with the
 *      PKCE verifier and NO client secret.
 *   2. setup → keychain write → silent token on restart.
 *   3. msal-node ≥ 5.6 strict ID-token nonce validation.
 *   3b. the device-code fallback redeems without a client secret.
 *   4. a keychain cache written by an OLDER msal-node (fixture generated
 *      with 5.1.5, see tests/fixtures/msal-legacy-cache/) still returns
 *      the cached token and refreshes with its refresh token.
 */
import { readFileSync } from "node:fs";
import path from "node:path";

import type { INetworkModule, NetworkRequestOptions } from "@azure/msal-node";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const { keychain } = vi.hoisted(() => ({ keychain: new Map<string, string>() }));

vi.mock("@napi-rs/keyring", () => ({
  // Vitest 4: constructor mocks need a `function`/class implementation.
  Entry: vi.fn().mockImplementation(function (service: string, account: string) {
    const key = `${service}/${account}`;
    return {
      getPassword: (): string => {
        const v = keychain.get(key);
        if (v === undefined) throw new Error("no entry");
        return v;
      },
      setPassword: (v: string): void => {
        keychain.set(key, v);
      },
      deletePassword: (): void => {
        keychain.delete(key);
      },
    };
  }),
}));

import {
  DELEGATED_SCOPES,
  deviceCodeRequest,
  getAccessToken,
  interactiveRequest,
  makeMsalClient,
} from "../../src/auth/msal.js";

// Keep in sync with tests/fixtures/msal-legacy-cache/generate.mjs.
const TENANT = "11111111-1111-1111-1111-111111111111";
const CLIENT_ID = "22222222-2222-2222-2222-222222222222";
const OID = "33333333-3333-3333-3333-333333333333";
const HOST = "login.microsoftonline.com";
const KEYCHAIN_KEY = `juvantlabs-m365-graph-mcp-server/tenant:${TENANT}`;

const OFFLINE_METADATA = {
  authorityMetadata: JSON.stringify({
    token_endpoint: `https://${HOST}/${TENANT}/oauth2/v2.0/token`,
    authorization_endpoint: `https://${HOST}/${TENANT}/oauth2/v2.0/authorize`,
    end_session_endpoint: `https://${HOST}/${TENANT}/oauth2/v2.0/logout`,
    issuer: `https://${HOST}/${TENANT}/v2.0`,
    jwks_uri: `https://${HOST}/${TENANT}/discovery/v2.0/keys`,
  }),
  cloudDiscoveryMetadata: JSON.stringify({
    tenant_discovery_endpoint: `https://${HOST}/${TENANT}/v2.0/.well-known/openid-configuration`,
    "api-version": "1.1",
    metadata: [
      {
        preferred_network: HOST,
        preferred_cache: "login.windows.net",
        aliases: [HOST, "login.windows.net", "login.microsoft.com", "sts.windows.net"],
      },
    ],
  }),
};

const b64url = (o: unknown): string => Buffer.from(JSON.stringify(o)).toString("base64url");

function idToken(extraClaims: Record<string, unknown> = {}): string {
  const now = Math.floor(Date.now() / 1000);
  return [
    b64url({ alg: "none", typ: "JWT" }),
    b64url({
      aud: CLIENT_ID,
      iss: `https://${HOST}/${TENANT}/v2.0`,
      tid: TENANT,
      oid: OID,
      sub: "subject",
      preferred_username: "user@example.test",
      name: "Example User",
      iat: now,
      nbf: now,
      exp: now + 3600,
      ver: "2.0",
      ...extraClaims,
    }),
    "sig",
  ].join(".");
}

interface FakeIdp {
  networkClient: INetworkModule;
  posts: Array<{ url: string; body: URLSearchParams }>;
}

function fakeIdp(opts: { accessToken: string; idTokenClaims?: Record<string, unknown> }): FakeIdp {
  const posts: FakeIdp["posts"] = [];
  const networkClient: INetworkModule = {
    async sendGetRequestAsync(url: string) {
      throw new Error(`unexpected network GET ${url}`);
    },
    async sendPostRequestAsync(url: string, options?: NetworkRequestOptions) {
      posts.push({ url, body: new URLSearchParams(String(options?.body ?? "")) });
      if (url.includes("/devicecode")) {
        return {
          status: 200,
          headers: {},
          body: {
            user_code: "FAKE-CODE",
            device_code: "fake-device-code",
            verification_uri: "https://microsoft.com/devicelogin",
            expires_in: 900,
            interval: 1,
            message: "To sign in, open https://microsoft.com/devicelogin and enter FAKE-CODE.",
          },
        } as never;
      }
      return {
        status: 200,
        headers: {},
        body: {
          token_type: "Bearer",
          scope: DELEGATED_SCOPES.filter((s) => s !== "offline_access").join(" "),
          expires_in: 3600,
          ext_expires_in: 3600,
          access_token: opts.accessToken,
          refresh_token: "rotated-refresh-token",
          id_token: idToken(opts.idTokenClaims),
          client_info: b64url({ uid: OID, utid: TENANT }),
        },
      } as never;
    },
  };
  return { networkClient, posts };
}

function client(idp: FakeIdp): ReturnType<typeof makeMsalClient> {
  return makeMsalClient({ ...OFFLINE_METADATA, system: { networkClient: idp.networkClient } });
}

const originalEnv = { ...process.env };

beforeEach(() => {
  keychain.clear();
  process.env.M365_TENANT_ID = TENANT;
  process.env.M365_CLIENT_ID = CLIENT_ID;
  delete process.env.M365_CLIENT_SECRET;
});

afterEach(() => {
  process.env = { ...originalEnv };
});

/**
 * Play the browser + identity provider for acquireTokenInteractive: read
 * the authorization URL MSAL built, then deliver an auth code to MSAL's
 * loopback listener the way Entra does (form_post to the redirect URI).
 */
function fakeBrowser(code: string): { openBrowser: (url: string) => Promise<void>; urls: URL[] } {
  const urls: URL[] = [];
  return {
    urls,
    async openBrowser(raw: string): Promise<void> {
      const url = new URL(raw);
      urls.push(url);
      const redirect = url.searchParams.get("redirect_uri") ?? "";
      // MSAL listens on 127.0.0.1; "localhost" may resolve to ::1 first.
      const target = redirect.replace("//localhost:", "//127.0.0.1:");
      const res = await fetch(target, {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({ code, state: url.searchParams.get("state") ?? "" }),
      });
      await res.text();
    },
  };
}

describe("setup interactive flow: public client, loopback + PKCE (real MSAL)", () => {
  it("builds the authorization URL with a free loopback port, PKCE S256, form_post and no nonce", async () => {
    const idp = fakeIdp({ accessToken: "fresh-access-token" });
    const browser = fakeBrowser("auth-code-123");
    await client(idp).acquireTokenInteractive(interactiveRequest(browser.openBrowser));

    expect(browser.urls).toHaveLength(1);
    const url = browser.urls[0];
    expect(`${url.origin}${url.pathname}`).toBe(`https://${HOST}/${TENANT}/oauth2/v2.0/authorize`);
    expect(url.searchParams.get("client_id")).toBe(CLIENT_ID);
    expect(url.searchParams.get("redirect_uri")).toMatch(/^http:\/\/localhost:\d+$/);
    expect(url.searchParams.get("redirect_uri")).not.toBe("http://localhost:3000");
    expect(url.searchParams.get("code_challenge_method")).toBe("S256");
    expect(url.searchParams.get("code_challenge")).toMatch(/^[A-Za-z0-9_-]{43}$/);
    expect(url.searchParams.get("response_mode")).toBe("form_post");
    expect(url.searchParams.has("nonce")).toBe(false);
    expect(url.searchParams.get("prompt")).toBe("select_account");
    const scopes = (url.searchParams.get("scope") ?? "").split(" ");
    for (const s of DELEGATED_SCOPES) expect(scopes).toContain(s);
  });

  it("redeems the code with the PKCE verifier and no client secret, then serves the token silently after a restart", async () => {
    const idp = fakeIdp({ accessToken: "fresh-access-token" });
    const browser = fakeBrowser("auth-code-123");
    const result = await client(idp).acquireTokenInteractive(
      interactiveRequest(browser.openBrowser),
    );

    expect(result.account?.username).toBe("user@example.test");
    expect(idp.posts).toHaveLength(1);
    const tokenUrl = new URL(idp.posts[0].url);
    expect(`${tokenUrl.origin}${tokenUrl.pathname}`).toBe(
      `https://${HOST}/${TENANT}/oauth2/v2.0/token`,
    );
    const body = idp.posts[0].body;
    expect(body.get("grant_type")).toBe("authorization_code");
    expect(body.get("code")).toBe("auth-code-123");
    expect(body.get("redirect_uri")).toBe(browser.urls[0].searchParams.get("redirect_uri"));
    expect(body.get("code_verifier")).toMatch(/^[A-Za-z0-9_-]{43,128}$/);
    expect(body.has("client_secret")).toBe(false);
    expect(body.has("client_assertion")).toBe(false);
    expect(keychain.has(KEYCHAIN_KEY)).toBe(true);

    // New client instance = server restart reading the keychain.
    const restarted = fakeIdp({ accessToken: "should-not-be-requested" });
    await expect(getAccessToken(client(restarted))).resolves.toBe("fresh-access-token");
    expect(restarted.posts).toHaveLength(0);
  });

  it("ignores a leftover M365_CLIENT_SECRET: nothing secret reaches the token endpoint", async () => {
    process.env.M365_CLIENT_SECRET = "leftover-secret";
    const idp = fakeIdp({ accessToken: "fresh-access-token" });
    await client(idp).acquireTokenInteractive(
      interactiveRequest(fakeBrowser("auth-code-123").openBrowser),
    );
    expect(idp.posts).toHaveLength(1);
    expect(idp.posts[0].body.has("client_secret")).toBe(false);
    expect(idp.posts[0].body.toString()).not.toContain("leftover-secret");
  });

  it("rejects an ID token carrying a nonce the request never sent (strict nonce validation)", async () => {
    const idp = fakeIdp({ accessToken: "x", idTokenClaims: { nonce: "unexpected" } });
    await expect(
      client(idp).acquireTokenInteractive(
        interactiveRequest(fakeBrowser("auth-code-123").openBrowser),
      ),
    ).rejects.toMatchObject({ errorCode: "nonce_mismatch" });
    expect(keychain.has(KEYCHAIN_KEY)).toBe(false);
  });

  it("surfaces an error returned to the loopback redirect and caches nothing", async () => {
    const idp = fakeIdp({ accessToken: "x" });
    const openBrowser = async (raw: string): Promise<void> => {
      const redirect = (new URL(raw).searchParams.get("redirect_uri") ?? "").replace(
        "//localhost:",
        "//127.0.0.1:",
      );
      await (
        await fetch(redirect, {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: new URLSearchParams({ error: "access_denied", error_description: "user cancelled" }),
        })
      ).text();
    };
    await expect(
      client(idp).acquireTokenInteractive(interactiveRequest(openBrowser)),
    ).rejects.toMatchObject({ errorCode: "access_denied" });
    expect(idp.posts).toHaveLength(0);
    expect(keychain.has(KEYCHAIN_KEY)).toBe(false);
  });
});

describe("setup --device-code (real MSAL)", () => {
  it("shows the user message, redeems the device code without a client secret, and caches the token", async () => {
    const idp = fakeIdp({ accessToken: "device-access-token" });
    const messages: string[] = [];
    const result = await client(idp).acquireTokenByDeviceCode(
      deviceCodeRequest((m) => messages.push(m)),
    );

    expect(result?.account?.username).toBe("user@example.test");
    expect(messages).toEqual([expect.stringContaining("FAKE-CODE")]);
    expect(idp.posts.map((p) => new URL(p.url).pathname)).toEqual([
      `/${TENANT}/oauth2/v2.0/devicecode`,
      `/${TENANT}/oauth2/v2.0/token`,
    ]);
    const token = idp.posts[1].body;
    expect(token.get("grant_type")).toBe("device_code");
    expect(token.get("device_code")).toBe("fake-device-code");
    expect(token.has("client_secret")).toBe(false);
    expect(keychain.has(KEYCHAIN_KEY)).toBe(true);
  });
});

describe("silent acquisition (real MSAL)", () => {
  it("throws the setup hint when the keychain holds no account", async () => {
    await expect(getAccessToken(client(fakeIdp({ accessToken: "x" })))).rejects.toThrow(
      "@juvantlabs/m365-graph-mcp-server setup",
    );
  });
});

describe("keychain cache written by an older msal-node (fixture: 5.1.5)", () => {
  const fixture = JSON.parse(
    readFileSync(
      path.resolve(__dirname, "../fixtures/msal-legacy-cache/cache-msal-node-5.1.5.json"),
      "utf8",
    ),
  ) as { generated_by: string; cache: Record<string, Record<string, Record<string, unknown>>> };

  /**
   * The fixture stores a placeholder instead of the JWT-shaped synthetic
   * ID token (secret scanners treat it as a credential). Rebuild it here;
   * every key and other field is exactly as msal-node 5.1.5 wrote it.
   */
  function legacyCache(): Record<string, Record<string, Record<string, unknown>>> {
    const cache = structuredClone(fixture.cache);
    for (const entry of Object.values(cache.IdToken)) {
      expect(entry.secret).toBe("<synthetic-id-token>");
      entry.secret = idToken();
    }
    return cache;
  }

  it("was generated by msal-node 5.1.5", () => {
    expect(fixture.generated_by).toBe("@azure/msal-node 5.1.5");
  });

  it("returns the cached access token with no network call", async () => {
    keychain.set(KEYCHAIN_KEY, JSON.stringify(legacyCache()));
    const idp = fakeIdp({ accessToken: "should-not-be-requested" });
    await expect(getAccessToken(client(idp))).resolves.toBe("legacy-access-token");
    expect(idp.posts).toHaveLength(0);
  });

  it("refreshes an expired legacy access token with the legacy refresh token", async () => {
    const cache = legacyCache();
    for (const at of Object.values(cache.AccessToken)) {
      at.expires_on = "1000000000";
      at.extended_expires_on = "1000000000";
    }
    keychain.set(KEYCHAIN_KEY, JSON.stringify(cache));

    const idp = fakeIdp({ accessToken: "refreshed-access-token" });
    await expect(getAccessToken(client(idp))).resolves.toBe("refreshed-access-token");
    expect(idp.posts).toHaveLength(1);
    expect(idp.posts[0].body.get("grant_type")).toBe("refresh_token");
    expect(idp.posts[0].body.get("refresh_token")).toBe("legacy-refresh-token");
    // Public client: the refresh is redeemed without a client secret.
    expect(idp.posts[0].body.has("client_secret")).toBe(false);
    // The refreshed cache is written back to the keychain.
    expect(keychain.get(KEYCHAIN_KEY)).toContain("refreshed-access-token");
  });
});
