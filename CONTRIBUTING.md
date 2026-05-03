# Contributing to `@juvantlabs/m365-graph-mcp-server`

This repo is an **MCP server** per the
[handbook docs/repo-types/mcp-server.md](https://github.com/juvantlabs/handbook/blob/main/docs/repo-types/mcp-server.md)
spec. The meta contributor guide (where to file what, PR process, commit
style, code-of-conduct, AI-assisted contribution conventions) lives at
[handbook CONTRIBUTING.md](https://github.com/juvantlabs/handbook/blob/main/CONTRIBUTING.md).

## Quick path

1. **Open an issue** describing the change. For substantial changes
   (new tool, auth model change, breaking API), an issue lets us
   discuss direction first.
2. **Branch** from `main`. Branch name: short, hyphenated, descriptive.
3. **Implement**:
   - Tools under `src/tools/<tool-name>.ts` — typed, schema-validated,
     no general-purpose URL forwarders.
   - Auth flow under `src/auth/`.
   - HTTP client under `src/client/`.
   - Tests under `tests/unit/` (mocked) + `tests/integration/`
     (against vendor sandbox).
4. **CI must pass**: lint, type-check, tests, audit, dead-code check,
   stdout discipline check.
5. **PR**: title `<area>: <imperative summary>`. Body cites related
   issue / FEAT / ADR. Include test plan.

## Anti-patterns

The 12-item checklist from the
[2026-05-03 ftaricano audit](https://gist.github.com/juvantlabs/a9fe0a76a23b0c1260b1e0ad3194a6da)
is codified in the spec. Highlights:

- **No `console.log` in `src/`** — corrupts MCP stdio framing. Always
  `console.error` for diagnostics. CI rule blocks regressions.
- **No general-purpose URL forwarder** primitive. Tools are typed,
  schema-validated, named operations.
- **No arbitrary local-filesystem ops** through caller-supplied paths.
  Sandbox to per-tenant root + validate every input path.
- **No dead defense layers**: every validator exported in `src/` must be
  imported from at least one production handler.
- **Modern MCP SDK** (`@modelcontextprotocol/sdk` ≥ 1.25.2).
- **README accuracy**: every documented env var must be wired in
  `process.env.X` reads.

## Code of conduct

All interactions follow the
[handbook Code of Conduct](https://github.com/juvantlabs/handbook/blob/main/docs/contributing/code-of-conduct.md).
Enforcement: `conduct@juvant.io`.

## AI-assisted contributions

AI-assisted commits include the standard co-author tag:

```
Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
```

Human author of the PR remains accountable for the change.
