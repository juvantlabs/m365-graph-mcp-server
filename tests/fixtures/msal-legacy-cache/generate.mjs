#!/usr/bin/env node
/**
 * Regenerates `cache-msal-node-5.1.5.json`: a token cache serialized by
 * an OLDER @azure/msal-node, used by tests/unit/msal_contract.test.ts to
 * prove that a keychain cache written before an msal-node upgrade still
 * yields tokens (and refreshes) after it.
 *
 * Fully offline: authority + cloud-discovery metadata are supplied in
 * config and the network module is a fake that answers the token
 * endpoint with synthetic, unsigned tokens. No real tenant, no secrets.
 *
 * Usage (install the old version somewhere isolated first):
 *   node tests/fixtures/msal-legacy-cache/generate.mjs \
 *     /path/to/node_modules/@azure/msal-node > cache-msal-node-5.1.5.json
 */
import { createRequire } from "node:module";
import path from "node:path";

const msalDir = process.argv[2];
if (!msalDir) {
  console.error("usage: generate.mjs <path-to-@azure/msal-node>");
  process.exit(2);
}
const require = createRequire(path.join(path.resolve(msalDir), "package.json"));
const msal = require(path.resolve(msalDir));

// Keep in sync with tests/unit/msal_contract.test.ts.
const TENANT = "11111111-1111-1111-1111-111111111111";
const CLIENT_ID = "22222222-2222-2222-2222-222222222222";
const OID = "33333333-3333-3333-3333-333333333333";
const HOST = "login.microsoftonline.com";
const SCOPES = [
  "User.Read",
  "Files.ReadWrite",
  "Sites.ReadWrite.All",
  "Calendars.ReadWrite",
  "OnlineMeetings.Read",
  "OnlineMeetingTranscript.Read.All",
  "offline_access",
];

const ID_TOKEN_PLACEHOLDER = "<synthetic-id-token>";

const b64url = (o) => Buffer.from(JSON.stringify(o)).toString("base64url");
const now = Math.floor(Date.now() / 1000);
const idToken = [
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
  }),
  "sig",
].join(".");

const networkClient = {
  async sendGetRequestAsync(url) {
    throw new Error(`unexpected GET ${url}`);
  },
  async sendPostRequestAsync() {
    return {
      status: 200,
      headers: {},
      body: {
        token_type: "Bearer",
        scope: SCOPES.filter((s) => s !== "offline_access").join(" "),
        // ~70 years: the legacy access token stays valid in the test.
        expires_in: 2_200_000_000,
        ext_expires_in: 2_200_000_000,
        access_token: "legacy-access-token",
        refresh_token: "legacy-refresh-token",
        id_token: idToken,
        client_info: b64url({ uid: OID, utid: TENANT }),
      },
    };
  },
};

let serialized = "";
const app = new msal.ConfidentialClientApplication({
  auth: {
    clientId: CLIENT_ID,
    clientSecret: "fixture-secret",
    authority: `https://${HOST}/${TENANT}`,
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
  },
  system: { networkClient },
  cache: {
    cachePlugin: {
      async beforeCacheAccess() {},
      async afterCacheAccess(ctx) {
        if (ctx.cacheHasChanged) serialized = ctx.tokenCache.serialize();
      },
    },
  },
});

await app.acquireTokenByCode({
  code: "fixture-code",
  scopes: SCOPES,
  redirectUri: "http://localhost:3000/auth/callback",
});

const parsed = JSON.parse(serialized);
// The synthetic ID token is JWT-shaped, which secret scanners (rightly)
// treat as a credential. Store a placeholder; the test substitutes a
// freshly built synthetic token. Keys and every other field stay exactly
// as the old msal-node serialized them.
for (const entry of Object.values(parsed.IdToken ?? {})) {
  entry.secret = ID_TOKEN_PLACEHOLDER;
}
process.stdout.write(
  JSON.stringify(
    { generated_by: `@azure/msal-node ${msal.version ?? "unknown"}`, cache: parsed },
    null,
    2,
  ) + "\n",
);
