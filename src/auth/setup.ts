/**
 * Interactive OAuth setup — run once via `npm run setup` (or
 * `m365-graph-mcp-server setup`) to populate the OS keychain with the
 * initial token cache. After this, `npm run dev` (or `npx ...`) uses
 * the cached refresh token silently for the lifetime of the refresh
 * grant (Microsoft default ~90 days; rolling).
 *
 * Flow:
 *   1. Build the authorization URL via MSAL.
 *   2. Open the user's default browser at that URL.
 *   3. Listen on http://localhost:3000/auth/callback for the
 *      ?code=... redirect.
 *   4. Exchange the code for tokens (MSAL writes to the cache plugin
 *      → keychain).
 *   5. Close the localhost listener and exit.
 */

import { exec } from "node:child_process";
import http from "node:http";
import { URL } from "node:url";

import { DELEGATED_SCOPES, REDIRECT_URI, makeMsalClient } from "./msal.js";

const CALLBACK_PORT = 3000;
const CALLBACK_TIMEOUT_MS = 5 * 60 * 1000; // 5 minutes

function openInBrowser(url: string): void {
  const cmd =
    process.platform === "darwin"
      ? `open "${url}"`
      : process.platform === "win32"
        ? `start "" "${url}"`
        : `xdg-open "${url}"`;
  exec(cmd, (err) => {
    if (err) {
      console.error(
        "[setup] Could not open browser automatically. Visit the URL above manually.",
      );
    }
  });
}

function waitForAuthCode(): Promise<string> {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      server.close();
      reject(new Error(`Timed out waiting for OAuth callback after ${CALLBACK_TIMEOUT_MS / 1000}s`));
    }, CALLBACK_TIMEOUT_MS);

    const server = http.createServer((req, res) => {
      const requestUrl = new URL(req.url ?? "/", `http://localhost:${CALLBACK_PORT}`);
      if (requestUrl.pathname !== "/auth/callback") {
        res.writeHead(404, { "Content-Type": "text/plain" });
        res.end("Not Found");
        return;
      }
      const code = requestUrl.searchParams.get("code");
      const error = requestUrl.searchParams.get("error");
      const errorDescription = requestUrl.searchParams.get("error_description");

      if (error) {
        res.writeHead(400, { "Content-Type": "text/html; charset=utf-8" });
        res.end(
          `<html><body><h2>OAuth error</h2><p><b>${error}</b></p>` +
            `<pre>${errorDescription ?? ""}</pre></body></html>`,
        );
        clearTimeout(timer);
        server.close();
        reject(new Error(`${error}: ${errorDescription ?? "(no description)"}`));
        return;
      }

      if (!code) {
        res.writeHead(400, { "Content-Type": "text/plain" });
        res.end("Missing authorization code in callback URL.");
        return;
      }

      res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
      res.end(
        "<html><body><h2>Authentication successful</h2>" +
          "<p>You can close this tab and return to the terminal.</p></body></html>",
      );
      clearTimeout(timer);
      server.close();
      resolve(code);
    });

    server.on("error", (err) => {
      clearTimeout(timer);
      reject(err);
    });

    server.listen(CALLBACK_PORT, "127.0.0.1");
  });
}

export async function runSetup(): Promise<void> {
  const client = makeMsalClient();

  const authUrl = await client.getAuthCodeUrl({
    scopes: DELEGATED_SCOPES,
    redirectUri: REDIRECT_URI,
    prompt: "select_account",
  });

  console.error("[setup] Starting OAuth flow against tenant:", process.env.M365_TENANT_ID);
  console.error("[setup] If your browser does not open automatically, visit:");
  console.error("");
  console.error(`  ${authUrl}`);
  console.error("");
  console.error(`[setup] Listening for callback on ${REDIRECT_URI}`);
  console.error("");

  openInBrowser(authUrl);
  const code = await waitForAuthCode();

  console.error("[setup] Authorization code received. Exchanging for tokens…");
  const tokenResponse = await client.acquireTokenByCode({
    code,
    scopes: DELEGATED_SCOPES,
    redirectUri: REDIRECT_URI,
  });

  if (!tokenResponse?.account) {
    throw new Error("Token acquisition succeeded but no account info was returned.");
  }

  console.error("[setup] ✓ Tokens cached in OS keychain for:");
  console.error(`        username: ${tokenResponse.account.username}`);
  console.error(`        homeAccountId: ${tokenResponse.account.homeAccountId}`);
  console.error("");
  console.error("[setup] You can now run `npm run dev` (or `npx @juvantlabs/m365-graph-mcp-server`).");
}
