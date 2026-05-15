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
export const DELEGATED_SCOPES = [
  "User.Read",
  "Files.ReadWrite",
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

export function makeMsalClient(): ConfidentialClientApplication {
  const tenantId = process.env.M365_TENANT_ID ?? "";
  const config: Configuration = {
    auth: {
      clientId: process.env.M365_CLIENT_ID ?? "",
      clientSecret: process.env.M365_CLIENT_SECRET ?? "",
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
    cache: {
      cachePlugin: makeCachePlugin(tenantId),
    },
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
