/**
 * MSAL Node client factory + cache plugin wiring.
 *
 * Uses ConfidentialClientApplication (we have a client_secret) with the
 * Authorization Code flow for delegated permissions. Per the handbook
 * spec § Auth, we never roll our own OAuth — MSAL Node is Microsoft's
 * official library and handles refresh, revocation, and edge cases.
 *
 * Tokens persist in the OS keychain via `src/auth/keyring.ts`. The MSAL
 * cache plugin pattern is: load on first access, save when it changes.
 */

import {
  ConfidentialClientApplication,
  type AuthorizationCodeRequest,
  type AuthorizationUrlRequest,
  type Configuration,
  type ICachePlugin,
  type TokenCacheContext,
} from "@azure/msal-node";

import { getTokenStore } from "./keyring.js";

/**
 * Delegated scopes the MCP server requests. Order is irrelevant; MSAL
 * normalizes. `offline_access` is required to get a refresh token.
 *
 * Add scopes here as new tools land — Files.ReadWrite for upload tools,
 * Calendars.ReadWrite for calendar write tools, etc. (per handbook spec
 * § Auth › Scopes: per-tool minimum, justified in ARCHITECTURE.md).
 */
// Files.ReadWrite subsumes Files.Read; Calendars.ReadWrite subsumes
// Calendars.Read. The Entra app permissions list still includes the
// narrower scopes (granted earlier) — they're harmless, just not
// requested at token acquisition time.
//
// `Sites.ReadWrite.All` (added v0.4.0): delegated scope for SharePoint
// document-library RW. Deliberately *not* `Sites.Manage.All` /
// `Sites.FullControl.All` / `Application.ReadWrite.All` — those carry
// permission-mutation privileges (sharing-link creation, guest invites,
// permission grants, ownership transfer) that exceed the v0.4.0 threat
// model. CI Layer A in `.github/workflows/ci.yml` fails the build if
// any of those broader scopes are introduced without a matching
// `permission_mutating` tool classification + allowlist entry.
export const DELEGATED_SCOPES = [
  "User.Read",
  "Files.ReadWrite",
  "Sites.ReadWrite.All",
  "Calendars.ReadWrite",
  "OnlineMeetings.Read",
  "OnlineMeetingTranscript.Read.All",
  "offline_access",
];

/**
 * The redirect URI registered in the Entra app for the OAuth callback.
 * Must exactly match one of the redirect URIs configured in the Entra
 * app registration. See README § Local development.
 */
export const REDIRECT_URI = "http://localhost:3000/auth/callback";

/**
 * Authorization-URL request used by `setup`. Deliberately no `nonce`
 * and no `responseMode`:
 *   - the confidential-client `getAuthCodeUrl` defaults to
 *     `response_mode=query`, which is what the setup localhost listener
 *     reads (GET `?code=`);
 *   - `acquireTokenByCode` (msal-node ≥ 5.6) validates the ID-token
 *     nonce strictly — if a nonce were sent here it would also have to be
 *     passed to the code exchange, otherwise the exchange fails with
 *     `nonce_mismatch`. Sending none keeps the two requests consistent.
 * Pinned by tests/unit/msal_contract.test.ts.
 */
export function authCodeUrlRequest(): AuthorizationUrlRequest {
  return {
    scopes: DELEGATED_SCOPES,
    redirectUri: REDIRECT_URI,
    prompt: "select_account",
  };
}

/** Code-exchange request used by `setup`; mirrors authCodeUrlRequest(). */
export function authCodeTokenRequest(code: string): AuthorizationCodeRequest {
  return {
    code,
    scopes: DELEGATED_SCOPES,
    redirectUri: REDIRECT_URI,
  };
}

/**
 * Build the MSAL cache plugin for a given tenant. Exported so tests
 * can drive the load/save lifecycle directly with a fake
 * TokenCacheContext + spied keychain store.
 */
export function makeCachePlugin(tenantId: string): ICachePlugin {
  const store = getTokenStore(tenantId);
  return {
    async beforeCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
      const data = store.load();
      if (data) {
        cacheContext.tokenCache.deserialize(data);
      }
    },
    async afterCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
      if (cacheContext.cacheHasChanged) {
        store.save(cacheContext.tokenCache.serialize());
      }
    },
  };
}

/**
 * Optional, additive overrides. Production calls `makeMsalClient()` with
 * none; tests pass offline authority/cloud-discovery metadata and a fake
 * network module so the real MSAL code paths run without a tenant.
 * Credentials, authority and cache plugin are never overridable here.
 */
export interface MsalClientOverrides {
  authorityMetadata?: string;
  cloudDiscoveryMetadata?: string;
  system?: Configuration["system"];
}

export function makeMsalClient(overrides: MsalClientOverrides = {}): ConfidentialClientApplication {
  const tenantId = process.env.M365_TENANT_ID ?? "";
  const config: Configuration = {
    auth: {
      clientId: process.env.M365_CLIENT_ID ?? "",
      clientSecret: process.env.M365_CLIENT_SECRET ?? "",
      authority: `https://login.microsoftonline.com/${tenantId}`,
      ...(overrides.authorityMetadata !== undefined
        ? { authorityMetadata: overrides.authorityMetadata }
        : {}),
      ...(overrides.cloudDiscoveryMetadata !== undefined
        ? { cloudDiscoveryMetadata: overrides.cloudDiscoveryMetadata }
        : {}),
    },
    cache: {
      cachePlugin: makeCachePlugin(tenantId),
    },
    ...(overrides.system !== undefined ? { system: overrides.system } : {}),
  };
  return new ConfidentialClientApplication(config);
}

/**
 * Acquire an access token silently from the cache. Refreshes via the
 * cached refresh token if the access token is expired. Throws if no
 * cached account exists — the caller should run `npm run setup` to
 * complete the initial OAuth flow.
 */
export async function getAccessToken(
  client: ConfidentialClientApplication,
): Promise<string> {
  const cache = client.getTokenCache();
  const accounts = await cache.getAllAccounts();
  if (accounts.length === 0) {
    throw new Error(
      "No cached account found in the keychain. Run `npm run setup` (or " +
        "`m365-graph-mcp-server setup`) once to complete the OAuth flow.",
    );
  }
  const result = await client.acquireTokenSilent({
    account: accounts[0],
    scopes: DELEGATED_SCOPES,
  });
  if (!result?.accessToken) {
    throw new Error(
      "Silent token acquisition returned no access token. The refresh " +
        "token may have been revoked or expired. Re-run `npm run setup`.",
    );
  }
  return result.accessToken;
}
