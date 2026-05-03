# Security

## Reporting a vulnerability

Please report vulnerabilities **privately** via one of these channels:

1. **GitHub Security Advisory** (preferred) — go to this repo's
   `Security` tab → `Report a vulnerability`. Your report stays
   private between you and the maintainer until we publish a
   coordinated advisory.
2. **Email** — `security@juvant.io`. Reports go to the primary
   maintainer.

**Please do NOT** open a public issue or pull request that contains
reproduction details for the vulnerability. Once a public artifact
exposes the issue, the coordinated-disclosure window collapses.

## What we commit to

This repo follows the
[juvantlabs Security Disclosure Process](https://github.com/juvantlabs/handbook/blob/main/docs/security/disclosure-process.md).
SLOs:

| State | Target |
|---|---|
| Acknowledge receipt | ≤ 7 days |
| Initial triage + severity classification | ≤ 14 days |
| Patch prepared (high/critical) | ≤ 30 days |
| Patch prepared (moderate) | ≤ 90 days |
| Public advisory + CVE | Patch + 1–7 days |

## Supported versions

| Version | Supported |
|---|---|
| Latest `0.x` (current) | ✅ |
| Older `0.x` | ❌ End-of-life with each new release until `1.0` |

Once `1.0` ships, the supported-versions matrix expands to formally
back-port security fixes to the `N-1` major.

## Out of scope

- Issues in dependencies — please report those upstream. We track
  upstream advisories via Dependabot and bump promptly.
- Issues in adopter customizations / forks of this server.
- Theoretical vulnerabilities without a reproduction path.

## Security-relevant dependencies

| Dependency | Version | Why it matters |
|---|---|---|
| `@modelcontextprotocol/sdk` | `^1.25.2` | ≥ 1.25.2 required to avoid ReDoS (`GHSA-8r9q-7v3j-jr4g`) and DNS rebinding (`GHSA-w48q-cv73-mx4w`) advisories on earlier versions |
| _(vendor SDK)_ | | _(fill in as the vendor SDK is selected)_ |

## Crediting

Reporters are credited by name in advisories unless they request
anonymity at report time. Past reports + reporter acknowledgments
will be listed in `SECURITY-CREDITS.md` (created on first disclosure).

## Acknowledgments

No disclosures yet.
