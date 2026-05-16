# Safety Rails

Every action either has a preview gate, a config-driven block, or both.

## `Get-ADLifecycleReport.ps1`

The audit report is read-only. The safety story is about correctness (not flagging the wrong account) rather than reversibility (no writes to reverse).

| Rail | Mechanism |
|---|---|
| Keep-note regex on Description | `(?i)\b(keep|retain|do not delete|legal hold|exempt)\b`, configurable via `-KeepNotePattern` |
| `DISABLE AFTER [date]` parser | Future date vetos. Past date scores normally. Unparseable date vetos as a fail-safe. |
| Explicit exclusions | CSV file, AD exclusion group, plain-text whitelist, or VIP regex via `.env` |
| SamAccountName pattern classification | `svc_`, `svc-`, `_svc`, `.svc`, `adm-`, `admin_`, `_adm`, `.adm`, `ea$` push to L3 sheet |
| DisplayName auto-exclude | External pattern file (`autoexclude-patterns.txt`). Never inline in source. |
| Vendor email domain | Regex via `.env` `VENDOR_EXCLUDE_DOMAINS` |
| Builtin/system account filter | `krbtgt`, `MSOL_*`, `HealthMailbox*`, `DiscoverySearchMailbox*` filtered out before scoring |
| EA suffix auto-exclude | SamAccountName ending in `ea` is always L3 |
| Computer-account filter | SamAccountName ending in `$` is auto-excluded |

## `Offboard.ps1`

The offboarding tool writes. The safety story is about reversibility and explicit confirmation.

| Rail | Mechanism |
|---|---|
| `-WhatIf` preview | `[CmdletBinding(SupportsShouldProcess, ConfirmImpact='High')]` native support |
| `-Confirm` prompt | Native, fires automatically at `ConfirmImpact='High'` |
| `whitelist.txt` | Plain-text file, one SamAccountName per line, gitignored. Checked before any AD lookup. |
| Sensitive-surname pattern | `.env` `SENSITIVE_SURNAMES` regex. Aborts with escalation message if matched. |
| svc / adm / EA pattern block | Hardcoded regex. These account types cannot be offboarded by this tool. Period. |
| Target OU uniqueness | OU lookup must return exactly one match. Ambiguous results abort. |
| Smart-resume | Each step (disable / move / set title) checks if already done. Re-runs are safe. |
| Old-title preservation | Original Title is stamped into Description as `TitleWas: <old>` for one-liner restore. |
| Append-only audit | Every run (success or partial failure) appends to `offboard.csv`. Never overwrites. |

## Restore path

If an offboarding needs to be reversed:

1. Read the relevant row from `offboard.csv` to recover the `OldTitle` value.
2. Move the account back out of the disabled OU.
3. Re-enable the account: `Enable-ADAccount -Identity <sam>`.
4. Restore Title: `Set-ADUser -Identity <sam> -Title <OldTitle>`.

The Description-field `TitleWas:` token serves as a fallback if the CSV row is ever lost.

## What is intentionally NOT in this repo

- **No delete step.** Disabled accounts stay in the disabled OU. Deletion (if ever performed) is a separate, manual decision after retention review of mailbox / OneDrive / Teams data.
- **No bulk offboarding.** The single-action design is deliberate. A bulk variant would require dedicated InfoSec sign-off and a separate confirmation pattern.
- **No re-enable script.** Restoration is a 4-line manual operation (see above). Codifying it risks accidental bulk restores from a stale CSV.
