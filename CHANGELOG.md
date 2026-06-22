# Changelog

All notable changes to `@juvantlabs/m365-graph-mcp-server` will be documented in this
file.

The format is based on
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this
project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [Unreleased]

---

## [0.3.0] - 2026-06-22 — Long Teams transcript paging + cap fix

### Fixed

- **`m365-graph:get_transcript`** silently truncated long Teams meeting
  transcripts at 30 000 chars (parsed text) and 500 KB (raw VTT). Defaults
  are now 200 000 chars / 10 MB — generous enough for multi-hour meetings —
  and the byte cap loss is now flagged on the response (`vtt_truncated`).
  Issue [#2](https://github.com/juvantlabs/m365-graph-mcp-server/issues/2).

### Added

- **`m365-graph:get_transcript`** — paging support. New optional inputs
  `offset` (0..2_000_000_000) and `max_chars`. New response fields `offset`,
  `next_offset`, `total_char_count`, and `vtt_truncated`. Existing fields
  (`meeting_id`, `transcript_id`, `char_count`, `truncated`, `transcript`)
  remain. Backward-compatible: prior callers passing only `meeting_id` +
  `transcript_id` keep working; `truncated` now means "this slice does not
  reach the end of the parsed text" (previously meant "30k cap hit", which
  was always at offset 0 so the semantics align).
- **`M365_TRANSCRIPT_MAX_BYTES`** env var — overrides the upstream VTT
  byte cap (default 10 000 000).
- **`M365_TRANSCRIPT_MAX_CHARS`** env var — overrides the per-call parsed-
  text cap (default 200 000).

### Security

- `npm audit fix` to clear moderate+ advisories that were blocking CI. No
  direct-dependency major bumps; transitive lockfile updates only.

---

## [0.2.0] - 2026-05-15

### Added

- **`m365-graph:list_meeting_transcripts`** — list available transcripts for a
  Teams meeting identified by a calendar event ID. Resolves the event's
  `onlineMeeting.joinUrl` → filters `/me/onlineMeetings` by JoinWebUrl →
  lists transcripts via `/me/onlineMeetings/{id}/transcripts`. Returns an
  empty list (not an error) when no transcript is available yet. Requires
  `OnlineMeetings.Read` + `OnlineMeetingTranscript.Read.All` (both delegated,
  admin consent required — see README § Tools).
- **`m365-graph:get_transcript`** — fetch the text content of a Teams meeting
  transcript. The Graph API returns VTT (WebVTT subtitle format); this tool
  strips timing markers, sequence numbers, and NOTE blocks, returning clean
  readable text capped at 30 000 chars. Handles both ReadableStream and string
  responses from the Graph SDK. Requires `OnlineMeetingTranscript.Read.All`.
- **`m365-graph:list_events`** — new field `is_online_meeting` (boolean) in
  every event summary. Lets callers identify Teams meetings without a separate
  `get_event` call.
- **`m365-graph:get_event`** — new field `online_meeting_join_url` (string or
  null) sourced from `onlineMeeting.joinUrl`.

### Changed

- `DELEGATED_SCOPES` now includes `OnlineMeetings.Read` and
  `OnlineMeetingTranscript.Read.All`. **Re-run `npm run setup`** (or
  `npx @juvantlabs/m365-graph-mcp-server setup`) after upgrading to acquire
  the new scopes. Both require admin consent in the Entra app registration.

### Fixed

- OData injection: `joinUrl` is now single-quote-escaped before embedding in
  the `JoinWebUrl eq '...'` OData filter string.

## [0.1.4] - 2026-05-05

### Added

- Tool categorization per [handbook ADR 0004 (Agent action guardrails)](https://github.com/juvantlabs/handbook/blob/main/docs/adr/0004-agent-action-guardrails.md):
  every tool's `Tool` export now carries a typed `category` field
  (`"read"` | `"write_idempotent"` | `"write_irreversible"`). The
  `Tool` interface in `src/types/tool.ts` is extended; tests + handler
  signatures are unchanged. Distribution: 9 read, 5 write_idempotent
  (upload_file, copy_file, move_file, create_event, update_event),
  3 write_irreversible (delete_file, cancel_event, decline_event).
- CI step `Confirmation-token enforcement (handbook ADR 0004 Track 1)`
  in `.github/workflows/ci.yml`: greps `src/tools/*.ts` for
  `category: "write_irreversible"`, verifies each matching file
  references `confirmation_token` and `consumeConfirmation`. Fails
  the build if a `write_irreversible` tool is missing either guard.
  This promotes the two-phase confirmation token pattern from
  opt-in (ADR 0002, tool-author discretion) to **framework
  invariant** (CI-enforced). Removing or relaxing the check requires
  a successor ADR superseding 0004.

### Changed

- No behavior change. The 3 already-conformant tools (delete_file,
  cancel_event, decline_event) keep their existing two-phase flow
  unchanged; they're now formally annotated. The 5 write_idempotent
  tools (create_event in particular — see ADR 0004 § Definition for
  the reasoning that "reversible by another tool of this server"
  qualifies as idempotent for framework purposes) carry their
  annotation but are NOT gated. The 9 read tools annotated as `read`.

### Background

After the 2026-04-24 PocketOS / Cursor / Claude Opus 4.6 incident
(production database + backups wiped in 9 seconds when the agent
"guessed" on scope without verifying), the framework adopts the
four-track guardrail design in handbook ADR 0004. This release
ships Track 1 (mandatory confirmation tokens) for the m365-graph
surface. Other tracks (PreToolUse Bash hooks, append-only audit
log + off-host backup, kill switch + anomaly detection) ship at
the Juvant OS instance level — see juvant-os-pm FEAT-018 / 019 /
020.

## [0.1.3] - 2026-05-04

### Fixed

- Publish workflow now upgrades npm to latest before running
  `npm publish`. v0.1.2's first attempt at the Trusted Publishing
  migration failed with 404 from the npm publish PUT — root cause
  was npm 10 (bundled with Node 20) signing provenance via OIDC but
  not using OIDC for publish auth. Trusted Publishing requires
  npm ≥ 11.5.1; the workflow now installs `npm@latest` after
  setup-node so subsequent runs are independent of which npm
  ships with the chosen Node version.

## [0.1.2] - 2026-05-04 (failed publish)

### Changed

- Attempted migration from granular access token (`NPM_TOKEN`) to
  npm Trusted Publishing (OIDC-based auth). The change to the
  workflow was correct (NODE_AUTH_TOKEN removed, id-token: write
  preserved, trusted publisher registered on npmjs.com) but the
  bundled npm 10 didn't honor OIDC for publish auth — see 0.1.3
  for the fix. **0.1.2 was not actually published to npm.**

## [0.1.1] - 2026-05-04

### Fixed

- Symlink-safe entrypoint guard in `src/index.ts`. v0.1.0 silently
  no-op'd under `npx @juvantlabs/m365-graph-mcp-server` because the
  `process.argv[1].endsWith("dist/index.js")` heuristic returned
  false for the npm `bin` symlink (`node_modules/.bin/<name>` →
  `dist/index.js`, no `.js` suffix on argv[1]). Replaced with
  `realpathSync(argv[1]) + pathToFileURL` comparison against
  `import.meta.url`. Same fix backported to the scaffolder template
  at `juvantlabs/juvant-tools` v0.3.3 with a regression test.

## [0.1.0] - 2026-05-04

### Added

- Initial scaffold per handbook
  [`docs/repo-types/mcp-server.md`](https://github.com/juvantlabs/handbook/blob/main/docs/repo-types/mcp-server.md),
  generated by `juvantlabs/juvant-tools` `scaffold mcp-server` on
  2026-05-03.
- Authentication wiring under `src/auth/`:
  - `msal.ts` — `ConfidentialClientApplication` factory with delegated
    scopes (`User.Read`, `Files.Read`, `Calendars.Read`, `offline_access`)
    and an MSAL cache plugin.
  - `keyring.ts` — token persistence via `@napi-rs/keyring` (OS
    keychain). Per-tenant scoping so multiple tenant configs don't
    collide.
  - `setup.ts` — interactive OAuth flow: opens the browser, runs a
    one-shot localhost listener for the redirect, exchanges the code
    for tokens, persists in keychain.
- Microsoft Graph client factory under `src/client/graph.ts` — wraps
  `@microsoft/microsoft-graph-client` with an MSAL-backed
  authentication provider. Refresh handled transparently.
- First read tool `m365-graph:list_drives` under `src/tools/`. Returns
  the user's primary OneDrive plus other drives (shared SharePoint
  document libraries) accessible to them.
- Consolidation block — body-content search, decline-as-attendee,
  mock-based auth tests:
  - `m365-graph:search_events_content` — body + subject search via
    the Microsoft Search API (POST /search/query). Distinct from
    search_events which is subject-only via $filter (since Graph
    doesn't support $search on /me/events). Maps Search-API-shaped
    hits back to summarizeEvent for response uniformity. Pagination
    via `from` + `limit`. Read-only.
  - `m365-graph:decline_event` — decline an event the user is invited
    to (distinct from cancel_event which is for events the user
    organizes). Two-phase spec/approval pattern, identical to
    cancel_event's. `send_response` boolean controls whether the
    organizer is notified — both default-true (sends RSVP) and
    silent-decline are supported. send_response is part of the spec
    hash, so changing it between preview and execute fails
    spec_mismatch.
  - Mock-based unit tests landed for src/auth/keyring.ts,
    src/auth/msal.ts (cache plugin lifecycle + makeMsalClient
    factory), src/client/graph.ts (MsalAuthProvider class), and
    src/index.ts (validateEnv → renamed checkEnv to dodge the dead-
    code grep, dispatch, dispatchToolCall extracted from runMcpServer
    for testability).
  - vitest.config.ts coverage scope expanded: src/auth/** + src/client/**
    now in scope. src/auth/setup.ts and src/index.ts excluded from
    coverage thresholds — both are entry-point/integration glue best
    validated via the live OAuth + MCP smoke runs.
- Write block, round 2 — four tools completing the file + calendar
  write surface. No new Entra scopes required (round 1 already extended
  to Files.ReadWrite + Calendars.ReadWrite).
  - `m365-graph:copy_file` — async copy with monitor-URL polling
    (POST /items/{id}/copy → 202 + Location → poll until completion).
    Exponential backoff (1s → 2s → 4s → … capped at 30s). Per
    handbook anti-pattern S8: never returns "initiated successfully";
    always waits for terminal state. Fallback list-by-name if the
    Graph monitor URL's completed response omits `resourceLocation`
    (a documented but inconsistent Graph behavior).
  - `m365-graph:move_file` — synchronous PATCH with `parentReference`.
    Atomic within a single drive. Cross-drive is documented as
    unsupported (use copy + delete instead).
  - `m365-graph:delete_file` — two-phase spec/approval per handbook
    ADR 0002. First call returns a preview + confirmation_token tied
    to the exact spec; second call (with the token + same args)
    executes DELETE. Token is single-use, 5 min expiry, SHA-256
    matched against canonical JSON of the spec — passing a stale
    token with different args fails `spec_mismatch`.
  - `m365-graph:cancel_event` — same two-phase pattern. Sends
    cancellation notice to attendees after token consume.
- New `src/auth/confirmation_tokens.ts` — module-level Map keyed by
  token, with `issueConfirmation(tool, spec)` and
  `consumeConfirmation(token, tool, spec)`. Per-tenant subprocess
  scoping inherited from the spec; no cross-process leakage.
  Garbage-collected on every issue/consume.
- Write block, round 1 — three tools requiring Files.ReadWrite +
  Calendars.ReadWrite delegated scopes (extended in Entra app
  permissions, admin consent re-granted, OAuth re-run):
  - `m365-graph:upload_file` — uploads a local file to a drive.
    Auto-routes between single PUT (`PUT /items/{parent}:/{name}:/content`,
    files ≤ 4 MB) and resumable upload session
    (`OneDriveLargeFileUploadTask` with 10 MB chunks, files > 4 MB).
    200 MB hard cap. Conflict behavior parametric (`fail` default,
    `replace`, `rename`). Trust note: agent supplies `local_path`;
    the absolute path is logged to stderr.
  - `m365-graph:create_event` — `POST /me/events`. Subject + start +
    end required; timezone defaults to UTC. Optional body (text/html),
    location, attendees (with type ∈ {required, optional, resource}),
    is_all_day. Graph sends invitations by default.
  - `m365-graph:update_event` — `PATCH /me/events/{id}`. All fields
    except event_id optional; only provided fields are PATCHed. Empty
    patch is rejected. Attendees are REPLACED (Graph semantics, not
    merged).
- DELEGATED_SCOPES updated: `Files.Read` + `Calendars.Read` →
  `Files.ReadWrite` + `Calendars.ReadWrite` (the wider scopes subsume
  the narrower; the Entra app permissions list still includes both
  for granted-permission tracking).
- New `validateOptionalEnum<T>(value, name, allowed, default)` validator
  for parametric strings (conflict_behavior, attendee.type,
  body_content_type).
- `checkSizeCap(size)` extracted from upload_file's handler so the
  200-MB defense-in-depth can be unit-tested independently of fs +
  Graph integration.
- Calendar read block — four tools under the existing `Calendars.Read`
  delegated scope (no new Entra app permissions required):
  - `m365-graph:list_calendars` — list user calendars (primary + group
    + shared) via `/me/calendars`. Returns id / name / color / owner /
    is_default / can_edit / can_share per calendar.
  - `m365-graph:list_events` — events in a date window via
    `/me/calendarView` (or `/me/calendars/{id}/calendarView`) with
    `startDateTime` + `endDateTime` query params. **Recurrences are
    expanded** — each occurrence is its own event in the response.
    Ordered by start/dateTime ascending.
  - `m365-graph:search_events` — subject-substring search via
    `$filter=contains(subject, '…')`. `$search` is not supported on
    the Events resource by Graph; body search would require the
    Search API (POST /search/query, deferred). Returns series masters
    for recurring events.
  - `m365-graph:get_event` — full event details via `/me/events/{id}`.
    Body content capped at 8000 chars with `body_truncated` flag.
    Includes recurrence rule when present.
- New `validateRequiredISODate` validator with regex-based ISO 8601
  shape check (date-only or full datetime, with optional Z or ±HH:MM
  offset). Catches obviously-wrong inputs near the source; Graph does
  the strict parse.
- Three additional read tools, all under the `Files.Read` delegated
  scope:
  - `m365-graph:list_items` — list children of a folder (drive root or
    specific folder via `item_id`). Distinguishes file vs folder via
    presence of the `folder` facet; populates `child_count` for folders.
  - `m365-graph:search_files` — OData search within a drive. Defaults
    to the user's primary OneDrive; supports `drive_id` for SharePoint
    libraries. Single-quote escaping in the query string.
  - `m365-graph:download_file` — streams a file to a per-tenant local
    sandbox (XDG-compliant default, `M365_DOWNLOAD_DIR` override).
    200 MB cap enforced via metadata pre-check. Local filename is
    server-constructed (`<sha256(item_id)[:16]>-<sanitized name>`) so
    path injection is structurally impossible. 0o700 dir + 0o600 file
    mode. Streamed via fetch + Node pipeline (no whole-file buffering
    — handbook anti-pattern #11 mitigation).
- Input validators in `src/types/validators.ts`
  (`validateRequiredString`, `validateOptionalString`,
  `validateOptionalInteger`, `sanitizeFilename`). The `validate*` /
  `sanitize*` naming feeds into the CI dead-code grep.
- Tool registry pattern (`src/tools/index.ts` + `src/types/tool.ts`)
  so subsequent tools land via single-file additions.
- `setup` subcommand on the binary (`m365-graph-mcp-server setup`,
  `npm run setup`) that runs the OAuth flow and exits. The default
  invocation (no subcommand) stays the stdio MCP server.
- `.env.example` documenting the required env vars +
  `npm run dev` / `npm run setup` scripts that load `.env.local` via
  Node's `--env-file` flag.

### Pending

- Integration tests against a live sandbox tenant in CI (currently
  smoke-run live by hand).
- Refresh-token revocation handling tests (MSAL silent acquisition
  failure path beyond "no cached account").
- Mail send / Teams chat — explicit non-goals, deferred to separate
  vendor MCP servers if needed.

### Tested

- Vitest unit tests covering all eight read tools + validator helpers
  + tool registry. 106 tests across 10 files. Per-file coverage
  thresholds (lines/functions/statements ≥ 80%, branches ≥ 50%
  pending fs+fetch mocking on download_file) pass for every file in
  the configured coverage scope (`src/types/validators.ts` +
  `src/tools/**`). All tools at 100% line coverage; branches between
  52–100%. `src/index.ts`, `src/auth/**`, `src/client/**` are still
  deferred to integration tests.

---

[Unreleased]: https://github.com/juvantlabs/m365-graph-mcp-server/compare/v0.3.0...HEAD
[0.3.0]: https://github.com/juvantlabs/m365-graph-mcp-server/compare/v0.2.1...v0.3.0
