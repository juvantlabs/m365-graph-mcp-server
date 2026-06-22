# Architecture — M365 Graph MCP Server

Design rationale for `@juvantlabs/m365-graph-mcp-server`. Read alongside
the [handbook MCP server spec](https://github.com/juvantlabs/handbook/blob/main/docs/repo-types/mcp-server.md)
for the cross-cutting conventions; this doc covers what's specific to
this server.

## Purpose

Wraps the Microsoft Graph API for AI-agent consumption: file operations
on OneDrive + SharePoint Online, calendar reads + writes on Microsoft
365 mailboxes. Fulfills the `m365-graph` abstract role per ADR 0002 in
the handbook; replaces ad-hoc per-instance Graph SDK code with a
single canonical MCP server, keeping the agent prompt surface clean and
the auth path consolidated.

## Scope

### In scope (planned, via incremental ships)

**OneDrive / SharePoint files**

- List drives a user has access to.
- List items in a drive (folders + files), with paging.
- Search files by name/content within a drive or site.
- Download a file (streamed; bounded by max-size guard).
- Upload a small file (single PUT, ≤ 4 MB).
- Upload a large file (resumable upload session, > 4 MB).
- Copy / move items (async; tool polls the monitor URL until completion).
- Delete items (gated by the spec/approval pattern referenced in [`docs/adr/0002-mcp-abstract-roles.md`](https://github.com/juvantlabs/handbook/blob/main/docs/adr/0002-mcp-abstract-roles.md) — destructive ops surface a "spec preview" before executing).

**Calendar**

- List user calendars.
- List events in a date range.
- Create / update / cancel events.
- Search events by subject / body.

**Meeting transcripts** (v0.2.0+)

- List available transcripts for a Teams meeting (from a calendar event ID).
- Fetch transcript content (VTT parsed to clean text). Long transcripts page
  client-side via `offset` + `next_offset` — the Graph content endpoint
  returns the whole VTT in one blob, so paging is implemented in this server
  by slicing the parsed text. Per-call cap (`M365_TRANSCRIPT_MAX_CHARS`,
  default 200 000) and upstream byte cap (`M365_TRANSCRIPT_MAX_BYTES`,
  default 10 MB) are configurable.
- Post-meeting only: live transcription during active calls requires a media bot
  (separate threat model; explicitly out of scope — see below).

### Out of scope

The exclusions below follow the principles formalized in
[handbook ADR 0003 — Scope boundaries for MCP servers](https://github.com/juvantlabs/handbook/blob/main/docs/adr/0003-mcp-server-scope-boundaries.md):
this MCP server's tools share a single threat model (side effects
confined to the user's own drive / calendar, reversible, no external
broadcast). Capabilities with materially different threat models go
into separate MCP servers; outbound-only notifications use webhooks
instead.

- **Mail send** — explicit non-goal for this server. Outbound email is a
  separate concern with its own threat model (SPF / DKIM / DMARC, reply-all
  blast radius, irreversibility). If a Juvant OS use case needs it, ships
  as `juvantlabs/m365-mail-mcp-server` per ADR 0003 § 1. For agent →
  human notifications, prefer a webhook (Adaptive Cards, Slack incoming)
  per ADR 0003 § 2.
- **Teams chat / channel posts** — same reasoning. Outbound Teams
  notifications already ship via the `notification.sh` hook in
  `juvantlabs/juvant-os` using a pre-shared Adaptive Cards webhook URL —
  no OAuth, no MCP server needed (ADR 0003 § 2). A hypothetical
  `juvantlabs/m365-teams-mcp-server` would only be built if read-back
  from Teams chat / interactive posting becomes a real Juvant OS need.
- **Tenant admin operations** — never automated by the agent layer. These
  remain manual / IT-administered.
- **General-purpose URL forwarder** — explicitly forbidden by handbook
  spec § Anti-patterns #2. Each tool is typed and schema-validated; the
  Microsoft Graph URL surface is hardcoded inside the tool, not
  caller-supplied.

## Authentication

OAuth 2.0 via `@azure/msal-node`'s `ConfidentialClientApplication`,
scoped per the application registration in Microsoft Entra (tenant
admin grants delegated and/or application permissions to the
registered app).

| Concern | Choice |
|---|---|
| OAuth library | `@azure/msal-node` (Microsoft's official) — never roll auth |
| Flow | Authorization Code with PKCE for delegated; Client Credentials for daemon ops |
| Scopes | Per-tool minimum: `Files.Read.All` for read-only file tools, `Files.ReadWrite.All` for write tools, `Calendars.Read` / `Calendars.ReadWrite` for calendar tools, `OnlineMeetings.Read` + `OnlineMeetingTranscript.Read.All` (both **admin consent required**) for transcript tools. Documented in [`README.md`](README.md) § Tools. |
| Token storage | `@napi-rs/keyring` (OS keychain). `keytar` is archived (handbook spec anti-pattern #10) — explicitly NOT used. |
| Token lifetime | Refresh token rotation handled by MSAL; refreshes never enter the agent's context. |
| Tenant ID | Validated at startup against the regex `^(common\|organizations\|consumers\|<UUID>)$` (handbook spec § Auth). Prevents arbitrary string interpolation into the authority URL. |

The MCP server process loads tokens at startup, refreshes on demand, and
exits when the client disconnects. **Per-tenant subprocess** — no
shared module-level state across tenants (handbook spec § Per-tenant
subprocess).

## Tenancy model

This server is a **self-hosted single-tenant** MCP server. Each adopter:

1. Registers their own Entra (Azure AD) application **inside their own tenant**.
2. Configures the application as **single-tenant** (`Accounts in this organizational directory only`) at registration time.
3. Generates a client secret (or, post-MVP, a federated credential).
4. Provides `M365_CLIENT_ID` + `M365_CLIENT_SECRET` + `M365_TENANT_ID` to the server at runtime via env vars (sourced from `.env.local` for local dev, from Azure Key Vault + Managed Identity for cloud deploys).

The npm package (`@juvantlabs/m365-graph-mcp-server`) is **code only**. No
central provider tenant exists. No customer onboarding flow. No cross-tenant
token storage. The package consumes credentials at runtime; the adopter owns
the auth surface end-to-end.

### Comparison with Anthropic's hosted M365 connector

| | Anthropic-hosted (`claude.ai Microsoft 365`) | This server (self-hosted, single-tenant) |
|---|---|---|
| Provider tenant | Anthropic, multi-tenant | None — the adopter IS the tenant |
| Token storage | Anthropic-side, central | Adopter's OS keychain (local) or Key Vault (cloud) |
| Consent flow | User-driven via claude.ai UI | Adopter's tenant admin grants to own app |
| Audit trail | Anthropic-side (opaque to adopter) | Adopter's AAD sign-in logs + per-call audit log |
| Failure mode of central infra | All adopters affected | Each adopter independent |
| Operating cost | Bundled in Claude subscription | Adopter pays own Azure resources (≤€1/mo at MVP) |

The two are complementary, not equivalent. The Anthropic-hosted connector is
appropriate for casual one-off M365 access. This server is appropriate when:

- The agent runs in an environment where token sovereignty matters (regulated industries, enterprise IT mandates).
- Audit trail must live with the adopter, not the LLM provider.
- The agent is part of a larger autonomous system (e.g. Juvant OS) where consistent self-hosted MCP boundaries are an architectural invariant ([handbook ADR 0003](https://github.com/juvantlabs/handbook/blob/main/docs/adr/0003-mcp-server-scope-boundaries.md)).

### Multi-tenant operation — technically possible, not the adoption pattern

`M365_TENANT_ID` accepts the multi-tenant authority strings
(`common` / `organizations` / `consumers`) at the regex layer because
those are valid Microsoft authority values. Configuring the server with
one of those values would enable multi-tenant operation against Microsoft Graph,
**but it is not the canonical adopter pattern this server was designed for.**

Operating multi-tenant inverts every row of the comparison table above:

- Token storage becomes the operator's central concern (no longer the adopter's keychain).
- Audit trail moves to whoever runs the multi-tenant deployment (not the consenting tenant).
- An admin-consent workflow becomes necessary — and **is not implemented in this package**.
- The blast-radius assumptions in the threat model below were drawn for single-tenant scope; they do not transfer.

Adopters who genuinely need to operate this code as a multi-tenant SaaS should
fork the package and build the consent + per-tenant token-isolation layer
themselves. That is a different product with a different threat model.

## Threat model

This server inherits the 12-item anti-pattern checklist from the
[ftaricano audit](https://gist.github.com/juvantlabs/a9fe0a76a23b0c1260b1e0ad3194a6da)
that informs the [handbook MCP server spec](https://github.com/juvantlabs/handbook/blob/main/docs/repo-types/mcp-server.md)
§ Anti-patterns. Specific defenses:

| Threat | Defense |
|---|---|
| Arbitrary local-FS write through `localPath` (audit C1–C4) | Every `download_file` / `upload_file` tool sandboxes to a per-tenant root. `path.resolve` + prefix check + symlink guard. Never trust caller-supplied `localPath`. |
| URL forwarder primitive (audit C5) | No such tool. Each tool's Graph URL is hardcoded. |
| Stdout corruption (audit C6) | `console.error` only; `console.log` blocked by ESLint + CI grep. |
| Outdated MCP SDK (audit C7) | `@modelcontextprotocol/sdk ^1.25.2` pinned in `package.json`. |
| Vulnerable axios (audit C8) | We use the Microsoft Graph SDK; no direct axios dep. |
| Vulnerable `jws` transitively (audit C9) | Quarterly `npm audit` (CI step every PR). |
| Defense-in-depth dead code (audit S1) | CI dead-code grep enforces every exported `validate*` / `sanitize*` / `guard*` is imported elsewhere in `src/`. |
| README env-var lies (audit S2) | CI README env-var accuracy check. |
| OData / URL injection (audit S3) | All Graph queries built via the SDK or with explicit `encodeURIComponent`. |
| Token storage (audit S5) | `@napi-rs/keyring`, never `keytar`. |
| Whole-file buffering (audit S7) | Downloads stream; max file size capped at 200 MB (configurable). Transcript reads cap raw VTT at `M365_TRANSCRIPT_MAX_BYTES` (default 10 MB) and parsed-text per response at `M365_TRANSCRIPT_MAX_CHARS` (default 200 000 chars). |
| No async-op polling (audit S8) | `copy` / `move` poll the monitor URL until completion; never return "initiated successfully" as the final result. |

### Universal Boundaries (per `SYSTEM_INVARIANTS.md` §4)

- No general-purpose URL forwarder primitive.
- Per-tenant subprocess (no shared cache state across tenants).
- Stdout discipline: `console.error` only outside protocol path.

### Delete-class operations

Delete tools (e.g. `delete_file`, `cancel_event`) follow the
**spec/approval pattern**: the agent submits a spec describing what to
delete; the tool returns a preview + a `confirmation_token`; a second
call with the token executes the delete. Mirrors the
`m365-delete-spec` pattern referenced in FEAT-014 and codified more
generally in the handbook MCP abstract roles ADR.

## Performance characteristics

- Typical request latency: 100–500 ms for unary Graph calls; multi-second
  for downloads (streamed).
- Max file size (download + upload): **200 MB** hard cap, server-side.
  Configurable via `M365_MAX_FILE_SIZE_BYTES` if the deployment needs a
  smaller cap; never larger.
- Streaming: downloads use the Graph SDK's stream interface; no
  whole-file `arraybuffer` reads (audit S7 mitigation).
- Async polling: `copy` / `move` ops poll the monitor URL with
  exponential backoff (1s, 2s, 4s, …, max 30s) until status is
  `succeeded` or `failed`. Tool returns the final state, never a 202.

## Tool catalog

Each row is also reflected in [`README.md`](README.md) § Tools.

| Tool | Underlying API call | Input | Output | Scope | Notes |
|---|---|---|---|---|---|
| `m365-graph:list_drives` | `GET /me/drive`, `GET /me/drives` | _(none)_ | `{ primary, accessible: [] }` | `Files.Read` | First-ship; validates the auth + Graph plumbing end-to-end. |
| `m365-graph:list_items` | `GET /me/drive/root/children` or `GET /drives/{id}/items/{item}/children` | `drive_id?`, `item_id?`, `limit?` (1–100, default 50) | `{ count, items: [] }` | `Files.Read` | Item type derived from presence of `folder` facet. `child_count` populated only for folders. |
| `m365-graph:search_files` | `GET /drives/{id}/root/search(q='…')` (defaults to `/me/drive`) | `query`, `drive_id?`, `limit?` (1–50, default 20) | `{ count, results: [] }` with virtual `path` joining `parentReference.path` + `name` | `Files.Read` | OData function call; query single quotes are escaped (`'` → `''`). |
| `m365-graph:download_file` | `GET /me/drive/items/{id}` for metadata, then `GET @microsoft.graph.downloadUrl` (Graph CDN) for the bytes | `item_id`, `drive_id?` | `{ local_path, size_bytes, name, content_type }` | `Files.Read` | **Sandboxed**: writes only under `<sandbox>/<tenant>/<sha256(item_id)[:16]>-<sanitized name>`. Streamed via fetch + Node `pipeline`; no whole-file buffering. 200 MB cap, refused pre-fetch via metadata size. Folders rejected. |
| `m365-graph:list_calendars` | `GET /me/calendars` | `limit?` (1–100, default 50) | `{ count, calendars: [] }` | `Calendars.Read` | Owner extracted from `owner.{name,address}`; falls back to null. |
| `m365-graph:list_events` | `GET /me/calendarView` (or `/me/calendars/{id}/calendarView`) with `startDateTime` + `endDateTime` query params | `start`, `end` (ISO 8601), `calendar_id?`, `limit?` (1–200, default 100) | `{ window, count, events: [] }` | `Calendars.Read` | **Recurrences expanded** (calendarView vs /events). Ordered by `start/dateTime` ascending. ISO 8601 input lightly validated client-side (regex), Graph does the strict parse. |
| `m365-graph:search_events` | `GET /me/events?$filter=contains(subject, '…')` | `query`, `limit?` (1–50, default 20) | `{ count, results: [] }` | `Calendars.Read` | Subject-only substring match. `$search` is not supported on the Events resource by Graph; body search would require POST `/search/query` (separate API, deferred). Single-quote escaping (`'` → `''`) on the query. **Series masters**, not expanded occurrences. |
| `m365-graph:get_event` | `GET /me/events/{id}` | `event_id` | event summary + body + body_truncated + recurrence | `Calendars.Read` | Body capped at 8000 chars (defense-in-depth against pathological event bodies); body_truncated flag indicates clipping. Recurrence rule passed through opaquely (Graph schema). |
| `m365-graph:upload_file` | `PUT /items/{parent}:/{name}:/content` (≤ 4 MB) or `OneDriveLargeFileUploadTask` (resumable, 10 MB chunks, > 4 MB) | `local_path`, `drive_id?`, `parent_item_id?`, `name?`, `conflict_behavior?` | `{ uploaded: { id, name, size, webUrl, upload_path } }` | `Files.ReadWrite` | Trust note: `local_path` from the agent; the MCP server reads from the user's filesystem (no sandbox on read — would defeat the upload's purpose). 200 MB cap (`checkSizeCap` defense-in-depth). Absolute path logged to stderr. |
| `m365-graph:create_event` | `POST /me/events` (or `/me/calendars/{id}/events`) | `subject`, `start`, `end` (required); `timezone?`, `body?`, `body_content_type?`, `location?`, `attendees?`, `is_all_day?`, `calendar_id?` | `{ created: <event summary> }` | `Calendars.ReadWrite` | Timezone defaults to UTC (Graph requires explicit TZ). Attendee.type ∈ {required, optional, resource}. Graph sends invites by default. |
| `m365-graph:update_event` | `PATCH /me/events/{id}` | `event_id` (required); any subset of subject/start/end/timezone/body/location/attendees/is_all_day | `{ updated: <event summary> }` | `Calendars.ReadWrite` | Empty patch rejected. **Attendees REPLACE, not merge** (Graph semantics) — pass the full intended list. Timezone required when start or end is updated (Graph rejects without TZ). |
| `m365-graph:copy_file` | `POST /items/{id}/copy` (raw response → 202 + Location) → poll monitor URL → `GET resourceLocation` (or fallback list-by-name) | `item_id`, `target_parent_id`; `source_drive_id?`, `target_drive_id?`, `new_name?`, `wait_max_seconds?` | `{ status: "completed", copied: { id, name, size, webUrl, parent_id } }` | `Files.ReadWrite` | **Async polling** with exponential backoff (1s → 2s → 4s → … capped at 30s). Per handbook anti-pattern S8: never returns "initiated successfully" — always polls to terminal state. Fallback list-by-name handles the Graph quirk where completed monitor responses sometimes omit `resourceLocation`. |
| `m365-graph:move_file` | `PATCH /me/drive/items/{id}` with `{parentReference: {id}, name?}` | `item_id`, `target_parent_id`; `drive_id?`, `new_name?` | `{ moved: { id, name, ... } }` | `Files.ReadWrite` | Synchronous within a single drive. Cross-drive moves are NOT supported by this PATCH (Graph's documented limitation); use copy + delete for those. |
| `m365-graph:delete_file` | Phase 1: `GET /items/{id}` → preview. Phase 2: `DELETE /items/{id}` after token consume. | `item_id` (required), `drive_id?`, `confirmation_token?` | preview or `{ deleted }` | `Files.ReadWrite` | **Spec/approval two-phase** per handbook ADR 0002. Token tied to canonical-JSON SHA-256 of the spec; passing a token issued for `{item_id: A}` together with `{item_id: B}` fails with `spec_mismatch`. Single-use, 5 min expiry. |
| `m365-graph:cancel_event` | Phase 1: `GET /events/{id}` → preview. Phase 2: `POST /events/{id}/cancel { Comment }` after token consume. | `event_id` (required), `comment?`, `confirmation_token?` | preview or `{ cancelled }` | `Calendars.ReadWrite` | Same two-phase pattern as delete_file. The `comment` is part of the spec — changing it between preview and execute fails `spec_mismatch`. |
| `m365-graph:decline_event` | Phase 1: `GET /events/{id}` → preview. Phase 2: `POST /events/{id}/decline { sendResponse, comment? }` after token consume. | `event_id`, `comment?`, `send_response?`, `confirmation_token?` | preview or `{ declined }` | `Calendars.ReadWrite` | For events the user is invited to (attendee). Cancel_event vs decline_event reflect the Graph distinction (organizer vs attendee). `send_response` is part of the spec hash so changing it between preview/execute fails `spec_mismatch`. Default true = organizer notified; false = silent decline. |
| `m365-graph:search_events_content` | `POST /search/query` with `entityTypes: ["event"]` + queryString | `query`, `limit?`, `from?` | `{ count, total, results: [<event summary>] }` | `Calendars.Read` | Subject + body search via the Microsoft Search API (separate from `$filter` on /me/events used by `search_events`). Search API hit shape `{ hitId, summary, resource }` mapped back to summarizeEvent's shape via `resource`. Returns recurrence series masters. |
| `m365-graph:list_meeting_transcripts` | 3-step: `GET /me/events/{id}` (resolve `onlineMeeting.joinUrl`) → `GET /me/onlineMeetings?$filter=JoinWebUrl eq '…'` (resolve onlineMeeting id) → `GET /me/onlineMeetings/{meeting-id}/transcripts` | `event_id` | `{ event_id, meeting_id, count, transcripts: [{ id, meeting_id, created_at, end_at }] }` | `Calendars.Read`, `OnlineMeetings.Read`, `OnlineMeetingTranscript.Read.All` (latter two delegated, **admin consent required**) | Non-online-meeting events return `{ error: "not_an_online_meeting" }`; missing join URL or unresolved onlineMeeting returns `{ error: "meeting_id_unavailable" }` (both as text content, not thrown). joinUrl is single-quote-escaped before OData filter interpolation. Empty `transcripts` list (not an error) when recording was disabled or transcript is still processing. |
| `m365-graph:get_transcript` | `GET /me/onlineMeetings/{meeting-id}/transcripts/{transcript-id}/content?$format=text/vtt` (returns raw WebVTT) | `meeting_id`, `transcript_id`, `offset?` (0..2_000_000_000, default 0), `max_chars?` (1..`M365_TRANSCRIPT_MAX_CHARS`, default = cap) | `{ meeting_id, transcript_id, offset, char_count, next_offset, total_char_count, truncated, vtt_truncated, transcript }` | `OnlineMeetingTranscript.Read.All` (delegated, **admin consent required**) | **VTT timing-marker stripping**: `WEBVTT` header, `NOTE` blocks, `HH:MM:SS.mmm --> …` timestamp lines, numeric cue-sequence lines, and blank lines are stripped; cue text (speaker names + spoken content) is kept and joined with `\n`. **Two caps** (defense-in-depth, audit-S7 whole-file-buffering bound): raw VTT bytes read from Graph capped at `M365_TRANSCRIPT_MAX_BYTES` (default 10 MB ≈ 8h+ of speech) — surplus is dropped and `vtt_truncated: true` is flagged on the response; parsed-text characters returned per call capped at `M365_TRANSCRIPT_MAX_CHARS` (default 200 000 ≈ 2h of speech). **Client-side paging**: long transcripts are paged by slicing the parsed text — caller passes the previous response's `next_offset` (or `null` if done) back as `offset` on the next call. `truncated` is `true` iff `next_offset !== null` (i.e. the returned slice does not reach the end of the parsed text); distinct from `vtt_truncated`, which signals upstream byte-cap loss. |

## Spec/approval confirmation-token pattern

Destructive tools (`delete_file`, `cancel_event`) implement a two-phase
flow per the handbook MCP server spec § Tool design and ADR 0002 to
prevent single-shot agent destruction in long autonomous loops:

| Phase | Required args | Tool action |
|---|---|---|
| 1 — preview | original args, **no** `confirmation_token` | Fetches a preview of what would be destroyed; issues a token tied to (tool name, canonical-JSON SHA-256 of spec, expiry timestamp). Returns preview + token. |
| 2 — execute | original args + correct `confirmation_token` | Verifies token (exists, not expired, tied to this exact tool, spec hash matches). Executes the destructive operation. Consumes token (single-use). |

State lives in `src/auth/confirmation_tokens.ts` as a module-level
`Map<token, {toolName, specHash, expiresAt}>`. Per-tenant subprocess
per the handbook spec means there's no cross-process leakage; tokens
expire 5 minutes after issue and are garbage-collected on every
issue/consume call.

The spec match is by SHA-256 of canonicalized JSON (keys sorted, top
level only). The agent cannot reuse a token across destructive ops:

- Token issued for `delete_file({item_id: "A"})`
- Agent attempts `delete_file({item_id: "B", confirmation_token: <token>})`
- → `spec_mismatch` error, deletion does not occur

Same protection for `cancel_event` if the agent changes the cancellation
comment between preview and execute.

### Download sandboxing

The `download_file` tool writes to a sandbox directory determined by:

1. `M365_DOWNLOAD_DIR` env var (override) → `<override>/<tenant-id>/`
2. else `XDG_CACHE_HOME` → `<XDG_CACHE_HOME>/m365-graph-mcp-server/<tenant-id>/`
3. else default → `~/.cache/m365-graph-mcp-server/<tenant-id>/`

Local filenames are server-constructed: `<sha256(item_id)[:16]>-<sanitized name>`,
NOT derived from caller-supplied paths (handbook anti-pattern #1
mitigation). Path injection via `item_id` is structurally impossible
because the agent never sees raw filesystem paths in the input. As
defense-in-depth, the resolved path is verified to start with the
sandbox root before writing.

Mode: dirs `0o700`, files `0o600` so other users on a shared host
can't read downloaded content.

## Input validation

Every tool validates its inputs via helpers in
[`src/types/validators.ts`](src/types/validators.ts):

- `validateRequiredString(v, name)` — non-empty string or throw
- `validateOptionalString(v, name)` — `string | undefined`
- `validateOptionalInteger(v, name, {min, max, default})` — bounded
- `sanitizeFilename(name)` — strips `/`, `\`, `\0`, leading dots; caps at 200 chars

Naming convention: every helper starts with `validate*` or
`sanitize*` so the [CI dead-code grep](https://github.com/juvantlabs/handbook/blob/main/docs/repo-types/mcp-server.md#ci-requirements)
enforces it's imported in at least one other file. Defense-in-depth
helpers that are never wired into a real handler are flagged as a
security smell (handbook anti-pattern S1).

## Dependencies

| Dependency | Version | Why |
|---|---|---|
| `@modelcontextprotocol/sdk` | `^1.25.2` | MCP framing; ≥1.25.2 required (ReDoS + DNS rebinding fixes — audit C7) |
| `@microsoft/microsoft-graph-client` | `^3.0.7` | Microsoft's official Graph SDK — handles batching, retries, types |
| `@azure/msal-node` | `^5.0.0` | OAuth — never roll auth. v5+ avoids transitive `uuid <14` GHSA-w5hq-g745-h8pq. |
| `@napi-rs/keyring` | `^1.3.0` | OS keychain for token persistence; replaces archived `keytar`. |
| `isomorphic-fetch` | `^3.0.0` | Peer dep of the Graph SDK. |

## References

- [Handbook MCP server spec](https://github.com/juvantlabs/handbook/blob/main/docs/repo-types/mcp-server.md)
- [Handbook MCP abstract roles ADR (0002)](https://github.com/juvantlabs/handbook/blob/main/docs/adr/0002-mcp-abstract-roles.md)
- [Handbook security disclosure process](https://github.com/juvantlabs/handbook/blob/main/docs/security/disclosure-process.md)
- [Juvant OS MCP_INVENTORY.md](https://github.com/juvantlabs/juvant-os/blob/main/docs/MCP_INVENTORY.md)
- [ftaricano audit (2026-05-03)](https://gist.github.com/juvantlabs/a9fe0a76a23b0c1260b1e0ad3194a6da) — origin of the 12-item anti-pattern checklist
