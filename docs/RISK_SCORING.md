# Risk Scoring

Every account in the audit report gets a `RiskScore` and a `TopReasons` string. Scoring is graduated (not on/off) so reviewers see severity, not just a binary flag.

## Scoring rubric

| Vector | Threshold | Score |
|---|---|---|
| Inactivity (enabled) | 365+ days | +5 |
| | 180+ days | +4 |
| | 90+ days | +2 |
| Never used (enabled, no logon ever) | account age 180+ days | +5 |
| | 90+ days | +4 |
| | 60+ days | +3 |
| | 30+ days | +1 |
| Disabled duration | 365+ days | +5 |
| | 180+ days | +4 |
| | 120+ days | +2 |
| Password staleness | 2+ years | +2 |
| | 1+ year | +1 |
| Missing manager (termination blind spot) | | +3 |
| Missing department | | +1 |
| Missing title | | +1 |
| In external stale list (when `-StaleListPath` supplied) | | +2 |
| Service / privileged pattern (don't flag system accounts as users) | | -2 |
| Rare-login title (licensed by design) | | -1 |
| Manual / grounds / property role (no license expected) | | +2 |

## Vetoes

A `RiskScore` of `-100` means the account is vetoed regardless of any other signal.

Vetoes fire on:

- `HasKeepNote`. Description matches the keep-note regex (`keep`, `retain`, `do not delete`, `legal hold`, `exempt`).
- `IsExcluded`. Account is in the whitelist file, exclusions CSV, exclusions group, or VIP pattern.
- `DISABLE AFTER [date]` in Description where the date is in the future. Use this for known returns from leave.
- `DISABLE AFTER [date]` in Description where the date is unparseable. Fail-safe.

## Role-aware adjustments

The role-aware section is the part most likely to need tuning for your environment. Two distinct cases that look similar but score in opposite directions:

1. **Rare-login titles, licensed by design.** Field or shift-worker roles where the company has decided everyone gets a license even though they sign in rarely. Inactivity here is expected, not staleness. The score is adjusted DOWN.
2. **Roles that should not hold licenses at all.** Manual, grounds, property, maintenance roles. If they have a license and rarely use it, that's a reclaim candidate. The score is adjusted UP.

Edit `$rareLoginTitlePatterns` and the manual-roles regex (search the script for `Manual/grounds`) to match your environment.

## TopReasons

Every score carries a semicolon-joined `TopReasons` string showing exactly which rules fired. A reviewer can read a row and know within seconds why it landed where it did. No black-box decisions.

Example: `Inactive (>=365d); No manager (termination blind spot); Pwd old (>=2yr); Role: Rare-login (licensed by design)` produces a score of `+5 + +3 + +2 + -1 = +9`, surfaced near the top of the L2 sheet.
