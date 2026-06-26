# ADR 0001 — SharePoint Lists support + multi-principal token storage

**Status:** Proposed — awaiting CEO ratification via [decision issue #6](https://github.com/juvantlabs/m365-graph-mcp-server/issues/6).
**Date:** 2026-06-23
**Component:** `@juvantlabs/m365-graph-mcp-server`
**Target release:** v0.4.0 (additive)

---

## Context

`@juvantlabs/m365-graph-mcp-server` v0.3.0 wraps the Microsoft Graph API
for OneDrive / SharePoint **files** (drives) and Outlook **calendar +
Teams transcripts**. It does **not** expose SharePoint **Lists**.

A new Juvant OS use case requires a structured CEO-upfront approval gate
for social-media publishing: Mira (CMO agent) writes proposed post
batches to a SharePoint List, the CEO reviews and flips a `Status`
choice column to `Approved` / `Rejected` / `Scheduled` in the
SharePoint UI, and Mira reads the approval state back before any
external publish action. A SharePoint List fits this better than an
Excel workbook (native typed choice columns, filterable views,
versioning / audit history, no cell-write friction).

There is already a Juvant precedent for SharePoint List access on
Microsoft Graph — the Vant lead-capture CRM — but it runs **outside
this MCP server**, in a server-to-server Azure Function under
**application-only client-credentials** authentication (see
`juvant-web/src/functions/retention/sharePointGraphClient.ts`). That
list holds customer-lead PII under a 24-month GDPR retention window
and intentionally operates under a different threat model from this
delegated-auth MCP server. The new Social Approvals list is a
deliberately separate concern from the lead-capture list.

This ADR records the design choices needed to add Lists CRUD to the
MCP server cleanly, including a non-trivial change to the token
persistence model that the existing single-principal cache does not
support.

## Decisions

### D1 — OAuth scope: `Sites.ReadWrite.All` delegated (not `Sites.Selected` app-only)

Add `Sites.ReadWrite.All` to `DELEGATED_SCOPES` in `src/auth/msal.ts`.
Read-only Lists tools rely on the read subset (`Sites.Read.All` is
subsumed by `Sites.ReadWrite.All`, mirroring the
`Files.Read`-subsumed-by-`Files.ReadWrite` pattern already established
in the file tools).

**Rejected alternatives:**

- **`Sites.Selected` (Application permission, per-site grants).**
  Strictly tighter blast radius at the OAuth-grant level — but
  available **only as an Application permission**, which would force
  the Lists tools onto **client-credentials / daemon flow** while
  every other tool in this server is **delegated / user-context**.
  This inverts the auth model inside a single MCP server: two code
  paths, two audit profiles (per-user sign-in logs vs per-app
  sign-in logs), and a de-facto violation of the "single threat
  model boundary per server" principle from
  [handbook ADR 0003](https://github.com/juvantlabs/handbook/blob/main/docs/adr/0003-mcp-server-scope-boundaries.md).
  We achieve the equivalent blast-radius outcome via D2 below
  without paying that complexity cost.

- **Separate `juvantlabs/m365-sharepoint-lists-mcp-server` package.**
  Cleanest ADR-0003 compliance in theory, but the threat-model delta
  vs the existing file-write tools (`upload_file`, `delete_file`,
  `copy_file`, `move_file`) is small — both write to the
  delegated user's M365 surface within the tenant — much smaller
  than the delta that justified carving out the hypothetical
  `m365-mail-mcp-server` (outbound email, irreversible broadcast).
  Splitting here would double the ops surface (two npm packages,
  two OAuth setups, two keychain entries) without a commensurate
  threat-model gain.

### D2 — Principal model: dedicated service account per functional role

The blast radius of the delegated `Sites.ReadWrite.All` token is
bounded by **the signed-in principal's effective SharePoint
permissions**, not by the OAuth grant. We exploit this:

- **CEO / design-lead** continue to use the existing principal for
  files + calendar + (read-only) lists in the existing site set.
- **Mira (CMO)** authenticates as a **dedicated service account** in
  the Juvant tenant. That service account is a member of **only** the
  new Social Approvals site collection. It has no access to the Vant
  lead-capture site, no access to executive OneDrive, no access to any
  other SharePoint site.
- Future functional roles get their own service accounts on the same
  pattern, added only to the sites they need.

The canonical blast-radius boundary is therefore **SharePoint site
membership**, managed in the tenant admin UI — not an env var inside
this MCP server.

**Hard precondition — must be ratified by the CEO:** Mira's service
account is **excluded from interactive-MFA Conditional Access
policies**, or configured for certificate-based non-interactive MFA.
Without this exclusion, `acquireTokenSilent` will fail when a CA policy
triggers an MFA prompt against the non-human principal — and the
entire approach degrades. If the CEO cannot grant this exclusion, this
ADR is revisited (likely landing on `Sites.Selected` app-only with the
auth-model inversion cost; strictly worse than the proposal here).

### D3 — Topology: separate site collection for Social Approvals

The new Social Approvals list is placed in a **new site collection**
(working name "GTM / Agent Ops"), **not** co-located with the Vant
lead-capture list:

| List | Site collection | Auth model | Threat profile |
|---|---|---|---|
| Vant Lead Capture (existing) | `juvantio.sharepoint.com,80d68ee6-…,f8d50dee-…` | App-only client-credentials, Azure Function | Customer PII, 24-month GDPR retention, server-to-server, no human in loop |
| Social Approvals (new, v0.4.0) | new site, e.g. "GTM / Agent Ops" | Delegated, MCP server, Mira service-account principal | Agent-proposed content, CEO-reviewed, no PII, audit via item version history |

Mira's service-account principal is added **only** to the Social
Approvals site. This reinforces the existing
lead-capture(app-only-Function) vs social-check(delegated-MCP)
separation by also separating them at the SharePoint site boundary.
A single compromise of either auth path cannot bridge to the other.

### D4 — Multi-principal token storage in v0.4.0

The current keychain key is `("juvantlabs-m365-graph-mcp-server",
"tenant:<id>")` — **one principal per tenant**
(see `src/auth/keyring.ts:25-26` and `src/auth/msal.ts:96-115`,
which selects `accounts[0]` unconditionally). The current usage
(single CEO principal) does not exercise this limit; D2 immediately
does (CEO principal + Mira service-account principal in the same
tenant). Without a fix, running `setup` for Mira after the CEO would
cache two accounts under the same key, and `accounts[0]` would
non-deterministically return whichever MSAL deserialized first — a
silent identity swap.

Therefore, in v0.4.0:

1. Keychain account key becomes `tenant:<id>:principal:<label>`
   (`src/auth/keyring.ts`).
2. New env var `M365_PRINCIPAL_LABEL`, default `"default"`.
3. Backwards compatibility: when `M365_PRINCIPAL_LABEL` is unset (or
   `"default"`) and no entry exists under the new key, attempt a
   one-time migration read from the old `tenant:<id>` key and rewrite
   under the new key.
4. `getAccessToken` (`src/auth/msal.ts`) tightens the account lookup —
   with per-principal keychain isolation each cache holds exactly one
   account, but we'll defensively match by `homeAccountId` rather than
   relying on `accounts[0]`.

This change is small (~60 LOC + tests) and **inseparable** from D1/D2.
Shipping Lists tools without it forces operators into ad-hoc
workarounds (separate working directories, separate XDG cache dirs,
separate processes) that defeat the clean default. In-scope for v0.4.0.

### D5 — Drop `M365_LIST_SITE_ALLOWLIST`; ship `m365-graph:whoami` instead

An earlier proposal added an MCP-layer env var listing allowed `site_id`s
on which write tools could operate. With D2 in place, that env var is
both redundant **and** misleading:

- **Redundant:** the canonical boundary is already the service
  account's site memberships in the tenant.
- **Misleading:** if memberships and the env var drift (and they
  will — memberships change in the tenant UI, the env var lives in
  deployment config), a reader of the README mistakes hygiene for
  security.

Replace with a diagnostic tool `m365-graph:whoami` that returns the
signed-in principal's identity (`GET /me`) and the SharePoint sites
accessible to it (`GET /sites?search=*`). Operators run `whoami` to
verify the real boundary before each deployment, and the README points
to it as the canonical check.

### D6 — `delete_list_item` deferred to v0.5.0

The v0.4.0 Lists tools cover read + create + update. Destructive
delete is **deferred** for two reasons:

1. The Buffer-approval flow uses `Status='Rejected'` as semantic
   delete — versions are preserved (audit), nothing is destroyed.
2. Shipping delete requires the existing two-phase spec/approval token
   pattern (`src/auth/confirmation_tokens.ts`, per
   [handbook ADR 0002](https://github.com/juvantlabs/handbook/blob/main/docs/adr/0002-mcp-abstract-roles.md)
   § Delete-class operations); not large, but additional surface,
   tests, and one more thing for the agent prompts to learn.

When v0.5.0 adds `delete_list_item`, it mirrors `delete_file`:
two-phase, single-use confirmation token, spec hashed on
`{site_id, list_id, item_id}`, 5-minute expiry.

### D7 — Semver: v0.3.0 → v0.4.0, additive

Pure additive tool surface (8 new tools) + one new env var
(`M365_PRINCIPAL_LABEL`, optional, defaulted) + one new delegated
scope (`Sites.ReadWrite.All`). No changes to existing tool signatures,
no removals.

Existing v0.3.0 deployments **re-consent on upgrade** because the
scope set changes — flagged prominently in `CHANGELOG.md` for v0.4.0.
The existing `m365-graph-mcp-server setup` flow handles re-consent
idempotently.

## Tools introduced in v0.4.0

| Tool | Category | Graph endpoint | Required scope (delegated) |
|---|---|---|---|
| `m365-graph:resolve_site` | read | `GET /sites/{hostname}:/{server-relative-path}` | `Sites.Read.All` |
| `m365-graph:list_site_lists` | read | `GET /sites/{site-id}/lists` | `Sites.Read.All` |
| `m365-graph:get_list` | read | `GET /sites/{site-id}/lists/{list-id}?$expand=columns` | `Sites.Read.All` |
| `m365-graph:list_list_items` | read | `GET .../lists/{id}/items?$expand=fields(...)&$filter=…&$orderby=…` | `Sites.Read.All` |
| `m365-graph:get_list_item` | read | `GET .../items/{id}?$expand=fields` | `Sites.Read.All` |
| `m365-graph:create_list_item` | write_idempotent | `POST .../items` | `Sites.ReadWrite.All` |
| `m365-graph:update_list_item` | write_idempotent | `PATCH .../items/{id}/fields` | `Sites.ReadWrite.All` |
| `m365-graph:whoami` | read | `GET /me` + `GET /sites?search=*` | `User.Read` + `Sites.Read.All` |

`delete_list_item` is **deferred** to v0.5.0 per D6.

## Column-type handling

SharePoint List columns are typed and dynamic. The MCP must not
hardcode any schema. Pattern:

- `get_list` returns each column's `name` (**internal** API name),
  `display_name` (UI name), `type`, `required`, and type-specific
  metadata (e.g. `choice.choices`).
- `create_list_item` / `update_list_item` validate the supplied
  `fields` object against the cached column schema **before** the
  Graph call:
  - **Choice / multiChoice:** value ∈ `choices`. Friendly error
    surfaces allowed values.
  - **DateTime:** ISO-8601 parseable.
  - **Boolean:** strict true / false.
  - **Number:** numeric.
  - **PersonOrGroup / Lookup:** object form `{ "LookupId": <id> }`,
    not bare email or display name (common Graph pitfall — documented
    explicitly in tool description).
  - **Text:** length-capped at 255 chars; **Note** capped at 8000
    chars (defense-in-depth).
  - **Required (create only):** all required columns present.
- The MCP accepts the **internal name** in `fields`, not the display
  name. The tool description for `create_list_item` / `update_list_item`
  states this with an example, because the SharePoint internal-vs-display
  name mismatch is the single most common Graph-Lists bug.

## Pagination, filtering, and audit

- **Pagination:** collection responses include `@odata.nextLink`. v0.4.0
  returns up to `limit` items and stops; client-side page-walking via a
  `skip_token` parameter is a clean additive follow-on if a list grows
  large.
- **Filtering on non-indexed columns:** Graph requires the
  `Prefer: HonorNonIndexedQueriesWarningMayFailRandomly` header for
  filters/orderby against unindexed columns on lists >5000 items.
  `list_list_items` sets this header unconditionally on filtered reads
  and surfaces a `filter_was_best_effort: true` flag in the response so
  callers know the filter may have been advisory.
- **Optimistic concurrency on update:** `update_list_item` accepts an
  optional `if_match_etag` and passes it as the `If-Match` header on
  the PATCH. Without it, last-write-wins. Useful once Mira proposing
  and CEO updating happen concurrently.
- **Audit / version history:** `GET /items/{id}/versions` is intentionally
  **not** in v0.4.0 — it's a clean ~100-LOC follow-on for v0.4.x once
  the flow needs explicit audit retrieval. SharePoint already retains
  the version history server-side; we just don't expose a tool to read
  it yet.

## Threat model deltas vs v0.3.0

| Concern | v0.3.0 baseline | v0.4.0 delta | Mitigation |
|---|---|---|---|
| Delegated token surface | `Files.ReadWrite` + `Calendars.ReadWrite` + transcript scopes | adds `Sites.ReadWrite.All` | Bounded by service-account site memberships per D2; not by OAuth grant. |
| Multi-principal token cache | N/A (one principal per tenant) | Two principals per tenant in practice (CEO + Mira) | D4: per-principal keychain isolation; defensive account lookup. |
| Cross-tenant token leakage | Not possible (per-tenant subprocess invariant) | Unchanged | Per-tenant subprocess invariant preserved by D4 (`tenant:<id>:principal:<label>` keeps the tenant in the prefix). |
| OData injection on field names / filter strings | N/A (no list filters yet) | New surface | Field names validated against `get_list` schema; filter strings caller-supplied but Graph SDK escapes through `.filter()` — additional unit tests with quote / backslash payloads. |
| Service-account credential storage | N/A | New surface (Mira's service-account password) | Tenant-side concern (Juvant secret store), not this package's responsibility. Documented in the v0.4.0 setup runbook. |

Universal Boundaries (per `ARCHITECTURE.md` § Threat model) unaffected:
no general-purpose URL forwarder, per-tenant subprocess preserved,
stdout discipline preserved.

## Implementation sketch (informational)

- New module `src/lists/` (or distribute across `src/tools/`)
  containing the 8 new tools, mirroring the file-tool layout.
- `src/auth/keyring.ts`: extend `getTokenStore(tenantId, principalLabel)`
  with default `"default"` and one-time migration read.
- `src/auth/msal.ts`: thread `principalLabel` through `makeCachePlugin`
  and `makeMsalClient`; tighten `getAccessToken` account lookup.
- `src/index.ts`: read `M365_PRINCIPAL_LABEL` from env in `checkEnv`,
  surface it in the startup log line.
- `src/tools/index.ts`: register the 8 new tools.
- README.md § Environment variables: document `M365_PRINCIPAL_LABEL`.
- README.md § Tools: 8 new rows.
- ARCHITECTURE.md § Tool catalog: 8 new rows; new § Multi-principal
  token storage; threat-model table extended.
- CHANGELOG.md v0.4.0 entry calls out re-consent requirement on upgrade.
- Tests: 8 unit tests under `tests/unit/` (one per new tool, mocked
  Graph client); 1 integration test under `tests/integration/` against
  a real Juvant tenant test list; keyring migration test; column-schema
  validation tests covering all column types.

Effort estimate end-to-end (implementation + tests + docs + ADR merge
+ release): **3–5 working days** after CEO ratification.

## Open / follow-on items (out of v0.4.0 scope)

- `delete_list_item` (v0.5.0, two-phase token pattern).
- `list_item_versions` (v0.4.x, additive, ~100 LOC).
- `skip_token` paging for very large lists (v0.4.x or later if
  proven needed).
- App-only client-credentials add-on for daemon use cases (deferred;
  current Vant lead-capture Function continues to live outside this
  server).

## Consequences

**Positive:**

- Buffer / social-approval workflow unblocked with a clean primitive
  surface inside the canonical MCP.
- Threat-model delta minimised: everything stays delegated, single
  server, single auth flow, audit lives with the tenant.
- Multi-principal storage solves a latent bug (`accounts[0]`
  non-determinism) that the current single-principal usage hides.
- Pattern (service account per functional role) scales cleanly to
  future agents without touching the OAuth grant.

**Negative:**

- One more env var (`M365_PRINCIPAL_LABEL`) on the server surface —
  defaulted to preserve existing UX.
- v0.4.0 upgrade is a re-consent event for existing adopters (single
  re-run of `setup`).
- Operational burden on the tenant admin: service-account creation,
  MFA exclusion, SharePoint-site membership management. None of these
  are this package's responsibility, but they are documented in the
  setup runbook as required pre-flight steps.

**Hard precondition for the entire ADR:** MFA exclusion on Mira's
service account. If the CEO cannot grant this, the ADR is revisited.

## References

- [Decision issue #6](https://github.com/juvantlabs/m365-graph-mcp-server/issues/6) — the `juvant:decision` ratification vehicle for this ADR.
- [Handbook ADR 0002 — MCP abstract roles](https://github.com/juvantlabs/handbook/blob/main/docs/adr/0002-mcp-abstract-roles.md) — spec/approval pattern for destructive ops.
- [Handbook ADR 0003 — MCP server scope boundaries](https://github.com/juvantlabs/handbook/blob/main/docs/adr/0003-mcp-server-scope-boundaries.md) — single-threat-model invariant.
- [Handbook ADR 0004 — Agent action guardrails](https://github.com/juvantlabs/handbook/blob/main/docs/adr/0004-agent-action-guardrails.md) — ToolCategory taxonomy.
- [Handbook MCP server spec](https://github.com/juvantlabs/handbook/blob/main/docs/repo-types/mcp-server.md) — anti-pattern checklist this server inherits.
- [`ARCHITECTURE.md`](../../ARCHITECTURE.md) — current server architecture, threat model, tool catalog.
- [`src/auth/keyring.ts`](../../src/auth/keyring.ts), [`src/auth/msal.ts`](../../src/auth/msal.ts) — token persistence model targeted by D4.
- `juvant-web/src/functions/retention/sharePointGraphClient.ts` (separate repo, informational) — existing Vant lead-capture app-only Graph access; the precedent we deliberately keep distinct from this server.
