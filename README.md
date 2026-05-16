# Identity Lifecycle Automation

Identity lifecycle automation for hybrid AD / Entra environments. A weekly read-only audit report split by responsibility tier, and a single-action offboarding tool with safety rails and a full audit trail.

## Overview

Two PowerShell tools for identity lifecycle management in a hybrid Microsoft environment (on-prem AD synced to Entra ID). Designed around three principles:

- **Read before write.** The audit script never modifies AD. Every offboarding action runs through `-WhatIf` first.
- **Separation of duties.** The audit report splits accounts by responsibility tier so L2 and L3 work in their own lanes.
- **Full audit trail.** Every offboarding action appends one row to a gitignored CSV. Re-runs are safe and pick up where they left off.

Built around Entra Group-Based Licensing: when an offboarded user's Title is set to the literal string `DISABLED`, GBL EQ and CONTAINS rules stop matching, so the M365 license deassigns automatically on the next sync cycle.

## Tools

### `Get-ADLifecycleReport.ps1`

Read-only AD scan that produces a single `.xlsx` with two sheets for two audiences.

**Tiered audience split.** One report, two audiences, no collisions:

- **`L2_Users`**: real user accounts, sorted by `RiskScore` descending. Columns surface what an L2 analyst needs to triage: top reasons, manager, description, `DisableBy` date, age buckets, OU path.
- **`L3_SpecialAccounts`**: service, admin, elevated-access, vendor, and generic accounts, sorted by inactivity. Includes an `AccountType` classification column so L3 can scan by type instead of sifting through user rows.

**Pipeline.** AD query, classification, risk scoring, two-sheet workbook via `ImportExcel`, SharePoint upload via `PnP.PowerShell` with cert-based app-only auth, timestamped run folder with `Transcript.log`.

**Risk scoring is graduated and transparent.** Every score carries a `TopReasons` string so a reviewer can see exactly which rules fired. Full rubric in [docs/RISK_SCORING.md](docs/RISK_SCORING.md).

**Three independent exclusion layers** sit on top of scoring:

- Description regex for keep-notes (`keep`, `retain`, `do not delete`, `legal hold`, `exempt`)
- `DISABLE AFTER [date]` parser. Future date vetos. Past date scores normally. Unparseable date vetos as a fail-safe.
- Explicit exclusions: CSV file, AD exclusion group, plain-text whitelist, or VIP regex from `.env`

Plus an external DisplayName auto-exclude pattern file (`autoexclude-patterns.txt`, gitignored) for generic mailboxes, service accounts, and vendor names. Patterns live in config, not source.

**Unattended SharePoint delivery.** `Connect-PnPOnline` with `ClientId` / `TenantId` / `Thumbprint` for cert-based auth, with interactive browser fallback if the cert isn't configured. Auto-creates dated subfolders, uploads all run artifacts, returns a SharePoint URL in the result object.

### `Offboard.ps1`

Single-action confirmed offboarding with audit trail.

```powershell
.\Offboard.ps1                                              # interactive prompts
.\Offboard.ps1 -SamAccountName jdoe -Reason "TICKET-12345"  # non-interactive
.\Offboard.ps1 -WhatIf                                      # preview only
.\Offboard.ps1 -SamAccountName jdoe -KeepTitle              # LOA/hold: license retained
```

**Execution sequence.** This order is non-negotiable to avoid unwanted group-policy interactions:

1. Disable the account
2. Stamp Description: `<existing> | Disabled YYYY-MM-DD | TitleWas: <old> | <reason>`
3. Move to the configured disabled OU
4. Set Title to literal `DISABLED` (breaks Entra GBL EQ + CONTAINS, license auto-deassigns)

**The Title-to-DISABLED move.** Setting Title to the literal string `DISABLED` is deliberate. The literal string ensures no edge-case group based licensing rule still matches, so the M365 license deassigns cleanly. This turns license reclamation from a separate workflow step into a reliable side effect of offboarding.

**Layered safety rails.** Full table in [docs/SAFETY_RAILS.md](docs/SAFETY_RAILS.md). Highlights:

- `[CmdletBinding(SupportsShouldProcess, ConfirmImpact='High')]`. `-WhatIf` and `-Confirm` work natively.
- Sensitive-surname pattern block via `.env` `SENSITIVE_SURNAMES` (regex)
- Service / admin / EA SamAccountName regex blocks
- `whitelist.txt` (gitignored) for protected accounts
- Target OU lookup must return exactly one match. Ambiguous results abort.
- Smart-resume: re-running on a partially-completed account is safe. Already-done steps are skipped.

**Audit trail.** Every run appends one row to `offboard.csv` (gitignored): timestamp, SAM, display name, old title, reason, action bools (`Disabled` / `Moved` / `TitleSet`), `FromOU`, `ToOU`, `RunBy`, error. Even failed mid-sequence runs log exactly what completed and what didn't. `Import-Csv .\offboard.csv | Format-Table` gives a one-command run history.

## Why this matters

- **Read-only audit before any write.** No script in this repo modifies AD without an explicit interactive confirmation or `-WhatIf` preview path.
- **Cross-functional reach.** The two-sheet split means HR, L2, and L3 each see exactly what they need to action, and none of what they don't.
- **License reclamation as a side effect.** The Title-to-DISABLED design turns M365 license deassignment into an automatic consequence of standard offboarding rather than a separate workflow step.
- **Production-grade PowerShell.** `[CmdletBinding(SupportsShouldProcess)]`, comment-based help, `.env`-driven config, try/catch with structured error capture, `Start-Transcript` per run, PnP cert-based app-only auth.

Both tools were built solo in production against a hybrid AD / Entra environment. The audit report is the investigative backstop for accounts L1 and HR miss. It runs weekly, lands in SharePoint for HR and IT to review together, and over the first three months in production surfaced 100+ at-risk accounts that hadn't moved through the standard offboarding workflow. Offboard.ps1 is what L2 reaches for when one of those flagged accounts needs action. It does not replace the L1 offboarding runbook. It runs from one command, picks up where it left off if a step fails part-way, and writes every action to a CSV.

## Setup

1. Clone the repo.
2. Copy `.env.example` to `.env` and fill in your environment values.
3. Copy `whitelist.template.txt` to `whitelist.txt` and add any accounts to protect (or leave empty).
4. Copy `autoexclude-patterns.template.txt` to `autoexclude-patterns.txt` and tune the regex list for your environment.
5. For SharePoint upload: create an Entra app registration with `Sites.Selected` (or an appropriate Files scope), grant access to the target site, and add the cert thumbprint to `.env`.

## Requirements

- PowerShell 7 (uses `-UseWindowsPowerShell` compat shim for the AD module)
- RSAT: Active Directory PowerShell module
- `ImportExcel` module (for the two-sheet workbook output)
- `PnP.PowerShell` module (for SharePoint upload, optional)

```powershell
Install-Module ImportExcel, PnP.PowerShell -Scope CurrentUser
```

## Repository layout

```
.
├── Get-ADLifecycleReport.ps1
├── Offboard.ps1
├── .env.example
├── whitelist.template.txt
├── autoexclude-patterns.template.txt
├── .gitignore
└── docs/
    ├── ARCHITECTURE.md
    ├── RISK_SCORING.md
    └── SAFETY_RAILS.md
```
