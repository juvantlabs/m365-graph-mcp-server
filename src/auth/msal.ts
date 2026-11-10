/**
 * MSAL Node client factory + cache plugin wiring.
 *
 * The server runs on the user's own machine with delegated permissions,
 * so it is a PUBLIC client: `PublicClientApplication`, no client secret,
 * no certificate. Nothing secret lives on disk except the MSAL token
 * cache, which is kept in the OS keychain (`src/auth/keyring.ts`).
 *
 * Initial sign-in (`setup`):
 *   - default: interactive authorization code + PKCE on a loopback
 *     redirect (`http://localhost:<random free port>`), handled by
 *     MSAL Node's `acquireTokenInteractive`;
 *   - fallback for headless machines: device code flow
 *     (`setup --device-code`).
 * Afterwards the server acquires tokens silently from the cached
 * refresh token. We never roll our own OAuth: MSAL Node handles PKCE,
 * refresh, revocation and edge cases.
 *
 * NOTE: this server intentionally uses its OWN app registration (env
 * vars `M365_TENANT_ID`, `M365_CLIENT_ID`) distinct from the
 * sibling `m365-mail-mcp-server`. Isolation is deliberate: a scope
 * misconfiguration in one server MUST NOT reuse the other's app and
 * silently escalate permissions.
 */

import {
  PublicClientApplication,
  type Configuration,
  type DeviceCodeRequest,
  type ICachePlugin,
  type InteractiveRequest,
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

const SUCCESS_HTML =
  "<html><body><h2>Authentication successful</h2>" +
  "<p>You can close this tab and return to the terminal.</p></body></html>";
const ERROR_HTML =
  "<html><body><h2>Authentication failed</h2>" +
  "<p>Return to the terminal for details.</p></body></html>";

/**
 * Interactive request used by `setup`. MSAL Node starts a one-shot
 * loopback listener on a free port of 127.0.0.1, builds the
 * authorization URL with redirect `http://localhost:<port>`, a PKCE
 * S256 challenge and `response_mode=form_post`, calls `openBrowser`,
 * then redeems the code with the PKCE verifier. No redirect URI is
 * passed: Entra accepts any port on `http://localhost` for the
 * "Mobile and desktop applications" platform. No nonce is sent, so the
 * strict ID-token nonce check (msal-node >= 5.6) has nothing to compare.
 * Pinned by tests/unit/msal_contract.test.ts.
 */
export function interactiveRequest(
  openBrowser: (url: string) => Promise<void>,
): InteractiveRequest {
  return {
    scopes: DELEGATED_SCOPES,
    openBrowser,
    prompt: "select_account",
    successTemplate: SUCCESS_HTML,
    errorTemplate: ERROR_HTML,
  };
}

/**
 * Device-code request used by `setup --device-code`, for machines with
 * no local browser. Requires "Allow public client flows" on the Entra
 * app registration.
 */
export function deviceCodeRequest(onMessage: (message: string) => void): DeviceCodeRequest {
  return {
    scopes: DELEGATED_SCOPES,
    deviceCodeCallback: (response) => {
      const message =
        response.message ||
        (response.userCode
          ? `Open ${response.verificationUri} and enter the code ${response.userCode}.`
          : "");
      if (message) onMessage(message);
    },
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
 * Client ID, authority and cache plugin are never overridable here.
 */
export interface MsalClientOverrides {
  authorityMetadata?: string;
  cloudDiscoveryMetadata?: string;
  system?: Configuration["system"];
}

export function makeMsalClient(overrides: MsalClientOverrides = {}): PublicClientApplication {
  const tenantId = process.env.M365_TENANT_ID ?? "";
  const config: Configuration = {
    auth: {
      clientId: process.env.M365_CLIENT_ID ?? "",
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
  return new PublicClientApplication(config);
}

/**
 * Acquire an access token silently from the cache. Refreshes via the
 * cached refresh token if the access token is expired. Throws if no
 * cached account exists: the user should run `setup` once to complete
 * the initial sign-in.
 */
export async function getAccessToken(
  client: PublicClientApplication,
): Promise<string> {
  const cache = client.getTokenCache();
  const accounts = await cache.getAllAccounts();
  if (accounts.length === 0) {
    throw new Error(
      "No cached account found in the keychain. Run `npx -y " +
        "@juvantlabs/m365-graph-mcp-server setup` once (with the same " +
        "M365_* environment variables) to sign in.",
    );
  }
  let result;
  try {
    result = await client.acquireTokenSilent({
      account: accounts[0],
      scopes: DELEGATED_SCOPES,
    });
  } catch (err) {
    // Typical causes: refresh token expired or revoked, consent changed,
    // or a token cached by a pre-0.6.0 (client-secret) version that the
    // public client cannot redeem. All are fixed by signing in again.
    const reason = err instanceof Error ? err.message : String(err);
    throw new Error(
      `Silent token acquisition failed: ${reason}. Re-run \`npx -y ` +
        "@juvantlabs/m365-graph-mcp-server setup` to sign in again (required " +
        "once after upgrading from a version older than 0.6.0).",
    );
  }
  if (!result?.accessToken) {
    throw new Error(
      "Silent token acquisition returned no access token. The refresh " +
        "token may have been revoked or expired. Re-run `setup`.",
    );
  }
  return result.accessToken;
}
