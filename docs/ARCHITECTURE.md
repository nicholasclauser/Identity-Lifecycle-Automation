# Architecture

Two tools, one identity lifecycle pipeline. Both read from a shared `.env`. Neither talks to the other directly. The audit report informs which accounts to offboard; the offboarding tool handles the action.

## `Get-ADLifecycleReport.ps1` (read-only)

```
.env / whitelist.txt / autoexclude-patterns.txt
                |
                v
        AD query (all enabled + disabled users in BaseDN)
                |
                v
        Classification (user / service / admin / EA / IS admin)
                |
                v
        Risk scoring (graduated, transparent, with TopReasons)
                |
                v
        Filter: exclusions, keep-notes, DISABLE AFTER dates
                |
                v
        Two-sheet workbook (L2_Users + L3_SpecialAccounts)
                |
                v
        SharePoint upload via PnP.PowerShell (cert-based)
```

No writes to AD. Every run produces a timestamped folder with the workbook, summary text, JSON summary, and full transcript.

## `Offboard.ps1` (single-action write)

```
.env / whitelist.txt
                |
                v
        Pre-flight checks:
          - whitelist match     -> abort
          - sensitive surname   -> abort
          - svc/adm/EA pattern  -> abort
          - target OU unique?   -> abort if not
                |
                v
        Preview (printed; if -WhatIf, stops here)
                |
                v
        Steps (smart-resume, already-done steps skipped):
          1. Disable account
          2. Update Description (preserve old title)
          3. Move to disabled OU
          4. Set Title = "DISABLED"  (breaks Entra GBL, license auto-deassigns)
                |
                v
        Append one row to offboard.csv
```

Title-to-DISABLED is the load-bearing piece. Entra GBL EQ and CONTAINS rules both stop matching once Title is the literal string `DISABLED`, so the M365 license deassigns on the next group-based licensing sync cycle.

## Why the split?

Two scripts, not one, because:

- The audit is read-only and runs unattended on a schedule. The offboarding tool is interactive and one-account-at-a-time.
- The audit serves multiple audiences (L2 / L3 / HR). The offboarding tool serves the L2 analyst running it.
- The audit is safe to share broadly. Offboarding is gated behind explicit human confirmation.

## Phases

The current scripts cover the read-only audit and the confirmed offboarding action. Future phases (not in this repo):

- **Phase 2: automated manager notification.** When an account hits 60 days of inactivity, email the manager with a structured A/B/C response (keep / disable / extended leave). Gated on Mail.Send scope and InfoSec approval.
- **Phase 3: ADP-AD sync.** Replace the manual Description-field signals (`Disabled TICKET-###`, `DISABLE AFTER [date]`) with HR-system-driven attributes. Reduces dependency on stopgap text patterns.
