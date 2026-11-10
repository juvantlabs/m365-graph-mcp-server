/**
 * One-time sign-in: `npx -y @juvantlabs/m365-graph-mcp-server setup`
 * (or `npm run setup` in a checkout). Populates the OS keychain with the
 * initial MSAL token cache; afterwards the MCP server uses the cached
 * refresh token silently for the lifetime of the refresh grant
 * (Microsoft default ~90 days, rolling).
 *
 * Reads only process environment variables (`M365_TENANT_ID`,
 * `M365_CLIENT_ID`); no file in the current directory is needed.
 *
 * Flows (public client, no secret):
 *   - default: interactive authorization code + PKCE. MSAL Node opens a
 *     one-shot listener on a free 127.0.0.1 port, we open the browser at
 *     the authorization URL, MSAL redeems the code with the PKCE
 *     verifier and writes the cache (→ keychain).
 *   - `--device-code`: device code flow for headless machines. The user
 *     enters a code at https://microsoft.com/devicelogin on any device.
 */

import { execFile } from "node:child_process";

import type { AuthenticationResult, PublicClientApplication } from "@azure/msal-node";

import { deviceCodeRequest, interactiveRequest, makeMsalClient } from "./msal.js";

const PACKAGE_NAME = "@juvantlabs/m365-graph-mcp-server";

export interface SetupOptions {
  deviceCode: boolean;
}

/** Parse the arguments that follow `setup` on the command line. */
export function parseSetupArgs(args: string[]): SetupOptions {
  const unknown = args.filter((a) => a !== "--device-code");
  if (unknown.length > 0) {
    throw new Error(
      `unknown setup option(s): ${unknown.join(" ")}. Usage: setup [--device-code]`,
    );
  }
  return { deviceCode: args.includes("--device-code") };
}

/** The platform command that opens `url` in the default browser. */
export function browserCommand(
  url: string,
  platform: NodeJS.Platform = process.platform,
): { file: string; args: string[] } {
  if (platform === "darwin") return { file: "open", args: [url] };
  if (platform === "win32") {
    return { file: "rundll32", args: ["url.dll,FileProtocolHandler", url] };
  }
  return { file: "xdg-open", args: [url] };
}

/**
 * Print the authorization URL, then try to open it. The URL is passed as
 * an argument (no shell), and a failure to open is not fatal: the user
 * can paste the printed URL.
 */
export async function openInBrowser(url: string): Promise<void> {
  console.error("[setup] If your browser does not open automatically, visit:");
  console.error("");
  console.error(`  ${url}`);
  console.error("");
  const { file, args } = browserCommand(url);
  execFile(file, args, (err) => {
    if (err) {
      console.error(
        "[setup] Could not open a browser automatically. Visit the URL above, " +
          "or re-run with --device-code on a machine without a browser.",
      );
    }
  });
}

export interface SetupDeps {
  client?: PublicClientApplication;
  openBrowser?: (url: string) => Promise<void>;
}

export async function runSetup(
  options: SetupOptions = { deviceCode: false },
  deps: SetupDeps = {},
): Promise<AuthenticationResult> {
  const client = deps.client ?? makeMsalClient();
  console.error(
    `[setup] Signing in against tenant ${process.env.M365_TENANT_ID} ` +
      `(${options.deviceCode ? "device code" : "browser, PKCE"} flow).`,
  );

  const result = options.deviceCode
    ? await client.acquireTokenByDeviceCode(
        deviceCodeRequest((message) => console.error(`[setup] ${message}`)),
      )
    : await client.acquireTokenInteractive(
        interactiveRequest(deps.openBrowser ?? openInBrowser),
      );

  if (!result?.account) {
    throw new Error("Token acquisition succeeded but no account info was returned.");
  }

  console.error("[setup] Tokens cached in the OS keychain for:");
  console.error(`        username: ${result.account.username}`);
  console.error(`        homeAccountId: ${result.account.homeAccountId}`);
  console.error("");
  console.error(
    `[setup] Done. Register the server with your MCP client, e.g. ` +
      `\`claude mcp add m365-graph -e M365_TENANT_ID=... -e M365_CLIENT_ID=... -- npx -y ${PACKAGE_NAME}\`.`,
  );
  return result;
}
