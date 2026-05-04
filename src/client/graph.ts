/**
 * Microsoft Graph client factory.
 *
 * Wraps the official `@microsoft/microsoft-graph-client` with an auth
 * provider that pulls cached tokens from MSAL on every request. MSAL
 * handles refresh transparently; if the refresh fails (revoked
 * token), the auth provider surfaces the error to the caller.
 *
 * Uses isomorphic-fetch (peer dep of the Graph client per its docs).
 * Imported once at module load.
 */

import "isomorphic-fetch";

import {
  Client,
  type AuthenticationProvider,
  type AuthenticationProviderOptions,
} from "@microsoft/microsoft-graph-client";
import type { ConfidentialClientApplication } from "@azure/msal-node";

import { getAccessToken } from "../auth/msal.js";

/**
 * Authentication provider that bridges MSAL's token cache → the
 * Microsoft Graph client. Each Graph request triggers
 * `getAccessToken()`, which lets MSAL refresh transparently if the
 * cached token is expired.
 *
 * Exported so tests can verify the bridge without instantiating the
 * full Graph client.
 */
export class MsalAuthProvider implements AuthenticationProvider {
  constructor(private readonly msal: ConfidentialClientApplication) {}

  async getAccessToken(_options?: AuthenticationProviderOptions): Promise<string> {
    return getAccessToken(this.msal);
  }
}

export function makeGraphClient(msal: ConfidentialClientApplication): Client {
  return Client.initWithMiddleware({
    authProvider: new MsalAuthProvider(msal),
  });
}
