<#
.SYNOPSIS
    Offboard a single AD user. Disables the account, preserves the old title in
    Description, moves to a configurable disabled OU, and sets Title to "DISABLED"
    so Entra Group-Based Licensing deassigns the M365 seat on the next sync cycle.

.DESCRIPTION
    Safety rails layered on top:
      * whitelist.txt (gitignored) protects specific SamAccountNames
      * Regex pattern blocks service / admin / elevated-access accounts
      * Configurable sensitive-surname pattern via .env SENSITIVE_SURNAMES
      * Target OU lookup must return exactly one match
      * Smart-resume: already-disabled, already-moved, already-titled steps are skipped
      * Every run appends one row to offboard.csv (gitignored) for audit

    Pair with -WhatIf to preview before applying. -KeepTitle for LOA/hold (license retained).
    Review run history: Import-Csv .\offboard.csv | Format-Table

.EXAMPLE
    .\Offboard.ps1                                          # interactive prompts
    .\Offboard.ps1 -SamAccountName jdoe -Reason "TICKET-12345" # non-interactive
    .\Offboard.ps1 -WhatIf                                  # preview only
    .\Offboard.ps1 -SamAccountName jdoe -KeepTitle          # LOA/hold: seat retained

.NOTES
    PowerShell 7 recommended (uses WinPS compat shim for the AD module).
    Reads optional config from a .env file next to this script:
      DISABLED_OU_NAME    = name of the target OU (default: User_Disabled_Master)
      SENSITIVE_SURNAMES  = regex pattern of surnames to block
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact='High')]
param(
    [string]$SamAccountName,
    [string]$Reason,
    [switch]$KeepTitle
)

if ([string]::IsNullOrWhiteSpace($SamAccountName)) {
    $SamAccountName = (Read-Host 'Username').Trim()
    if ([string]::IsNullOrWhiteSpace($SamAccountName)) { Write-Error 'Username required.'; return }
}
if ([string]::IsNullOrWhiteSpace($Reason)) {
    $Reason = (Read-Host 'Reason').Trim()
    if ([string]::IsNullOrWhiteSpace($Reason)) { Write-Error 'Reason required.'; return }
}

$DisabledTitle = 'DISABLED'

$scriptDir = $PSScriptRoot
$whitelist = Join-Path $scriptDir 'whitelist.txt'
$csvPath   = Join-Path $scriptDir 'offboard.csv'

function Get-DotEnv {
    param([string]$Path)
    $map = @{}
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $map }
    foreach ($line in Get-Content -LiteralPath $Path) {
        $l = $line.Trim()
        if (-not $l -or $l.StartsWith('#')) { continue }
        $kv = $l -split '=', 2
        if ($kv.Count -eq 2) {
            $k = $kv[0].Trim()
            $v = $kv[1].Trim()
            if (($v.StartsWith('"') -and $v.EndsWith('"')) -or ($v.StartsWith("'") -and $v.EndsWith("'"))) {
                $v = $v.Substring(1, $v.Length - 2)
            }
            $map[$k] = $v
        }
    }
    return $map
}

$envMap            = Get-DotEnv -Path (Join-Path $scriptDir '.env')
$TargetOuName      = if ($envMap.ContainsKey('DISABLED_OU_NAME'))   { $envMap['DISABLED_OU_NAME'] }   else { 'User_Disabled_Master' }
$sensitiveSurnames = if ($envMap.ContainsKey('SENSITIVE_SURNAMES')) { $envMap['SENSITIVE_SURNAMES'] } else { $null }

# Suppress -WhatIf for the read-only setup phase. Get-ADUser and Get-ADOrganizationalUnit
# must actually run during a preview so we can show what the change would look like.
# $WhatIfPreference is restored before any write operation (search this file for "savedWhatIf").
$savedWhatIf = $WhatIfPreference
$WhatIfPreference = $false

if ($PSVersionTable.PSVersion.Major -ge 7) {
    Import-Module ActiveDirectory -UseWindowsPowerShell -ErrorAction Stop
} else {
    Import-Module ActiveDirectory -ErrorAction Stop
}

try {
    $user = Get-ADUser -Identity $SamAccountName -Properties Surname,DisplayName,Title,Description,Enabled,DistinguishedName -ErrorAction Stop
} catch {
    Write-Error "User '$SamAccountName' not found."
    return
}

# Sensitive-surname block. Pattern (regex) comes from .env SENSITIVE_SURNAMES.
# Use this for families or individuals whose accounts require explicit higher-tier approval.
if ($sensitiveSurnames -and $user.Surname -and $user.Surname.Trim() -match $sensitiveSurnames) {
    Write-Error "BLOCKED: '$($user.SamAccountName)' surname matches the sensitive-surname pattern. Escalate."
    return
}

if (Test-Path $whitelist) {
    $wl = Get-Content $whitelist |
        Where-Object { $_ -and -not $_.TrimStart().StartsWith('#') } |
        ForEach-Object { $_.Trim().ToLower() }
    if ($wl -contains $user.SamAccountName.ToLower()) {
        Write-Error "BLOCKED: '$($user.SamAccountName)' is whitelisted."
        return
    }
}

$blockedPatterns = @(
    '^svc[_-]','^service[_-]','_svc$','\.svc$',
    '^adm[_-]','^admin[_-]','_adm$','\.adm$',
    'ea$'
)
foreach ($p in $blockedPatterns) {
    if ($user.SamAccountName -imatch $p) {
        Write-Error "BLOCKED: '$($user.SamAccountName)' matches protected pattern '$p'."
        return
    }
}

$ouResults = @(Get-ADOrganizationalUnit -Filter "Name -eq '$TargetOuName'" -ErrorAction Stop)
if ($ouResults.Count -eq 0) { Write-Error "Target OU '$TargetOuName' not found."; return }
if ($ouResults.Count -gt 1) { Write-Error "Multiple OUs named '$TargetOuName', can't disambiguate."; return }
$targetOu = $ouResults[0]

$stamp = Get-Date -Format 'yyyy-MM-dd'

$needsDisable      = [bool]$user.Enabled
$alreadyInTargetOu = $user.DistinguishedName -like "*,$($targetOu.DistinguishedName)"
$needsMove         = -not $alreadyInTargetOu
$titleAlreadySet   = ($user.Title -eq $DisabledTitle)
$needsTitleSet     = -not $KeepTitle -and -not $titleAlreadySet

# Description: <existing> | Disabled YYYY-MM-DD | TitleWas: <old> | <reason>
# On fix-up runs (already disabled, only title changing) we just patch in
# TitleWas if it's missing, never re-stamp Disabled or Reason.
$descAlreadyHasTitleWas = ($user.Description -match 'TitleWas:')

$descParts = @()
if ($user.Description) { $descParts += $user.Description }

$descChanged = $false

if ($needsDisable) {
    $descParts += "Disabled $stamp"
    if ($user.Title -and -not $titleAlreadySet) { $descParts += "TitleWas: $($user.Title)" }
    $descParts += $Reason
    $descChanged = $true
} elseif ($needsTitleSet -and $user.Title -and -not $descAlreadyHasTitleWas) {
    $descParts += "TitleWas: $($user.Title)"
    $descChanged = $true
}

$newDesc  = if ($descChanged) { $descParts -join ' | ' } else { $null }
$newTitle = if ($needsTitleSet) { $DisabledTitle } else { $null }

Write-Host ""
Write-Host "Offboarding plan for $($user.SamAccountName):" -ForegroundColor Yellow
Write-Host "  DisplayName     : $($user.DisplayName)"
Write-Host "  Current Enabled : $($user.Enabled)"
Write-Host "  Current Title   : $($user.Title)"
Write-Host "  Current DN      : $($user.DistinguishedName)"
Write-Host ""
Write-Host '  Steps:'

if ($needsDisable) {
    Write-Host '    [1] DISABLE account' -ForegroundColor Cyan
    Write-Host "        Description: '$($user.Description)'"
    Write-Host "                  -> '$newDesc'"
} else {
    Write-Host '    [1] DISABLE: skip (already disabled)' -ForegroundColor DarkGray
    if ($descChanged) {
        Write-Host '        Description fix-up (adding TitleWas):' -ForegroundColor Cyan
        Write-Host "          '$($user.Description)'"
        Write-Host "       -> '$newDesc'"
    }
}

if ($needsMove) {
    Write-Host '    [2] MOVE to User_Disabled_Master' -ForegroundColor Cyan
    Write-Host "        Target: $($targetOu.DistinguishedName)"
} else {
    Write-Host '    [2] MOVE: skip (already in target OU)' -ForegroundColor DarkGray
}

if ($KeepTitle) {
    Write-Host '    [3] SET TITLE: skip (-KeepTitle, seat retained)' -ForegroundColor DarkGray
} elseif ($titleAlreadySet) {
    Write-Host "    [3] SET TITLE: skip (already '$DisabledTitle')" -ForegroundColor DarkGray
} else {
    Write-Host "    [3] SET TITLE = '$DisabledTitle' (breaks GBL EQ + CONTAINS, license deassigns)" -ForegroundColor Cyan
    Write-Host "        '$($user.Title)' -> '$newTitle'"
}
Write-Host ""

if (-not ($needsDisable -or $needsMove -or $needsTitleSet -or $descChanged)) {
    Write-Warning "Nothing to do."
    return
}

$WhatIfPreference = $savedWhatIf

$parts = @()
if ($needsDisable)  { $parts += 'disable' }
if ($needsMove)     { $parts += 'move' }
if ($needsTitleSet) { $parts += 'set-title' }
$action = "Offboard ($($parts -join ' + '))"

if ($PSCmdlet.ShouldProcess($user.SamAccountName, $action)) {

    $didDisable  = $false
    $didMove     = $false
    $didTitleSet = $false
    $errMsg      = $null

    try {
        if ($needsDisable) {
            Disable-ADAccount -Identity $user.DistinguishedName -Confirm:$false -ErrorAction Stop
            Set-ADUser -Identity $user.DistinguishedName -Description $newDesc -Confirm:$false -ErrorAction Stop
            $didDisable = $true
            Write-Host "  [1] Disabled. Description updated." -ForegroundColor Green
        } elseif ($descChanged) {
            Set-ADUser -Identity $user.SamAccountName -Description $newDesc -Confirm:$false -ErrorAction Stop
            Write-Host "  [1] Description patched (TitleWas added)." -ForegroundColor Green
        }

        if ($needsMove) {
            Move-ADObject -Identity $user.DistinguishedName -TargetPath $targetOu.DistinguishedName -Confirm:$false -ErrorAction Stop
            $didMove = $true
            Write-Host "  [2] Moved to $TargetOuName." -ForegroundColor Green
        }

        if ($needsTitleSet) {
            Set-ADUser -Identity $user.SamAccountName -Title $newTitle -Confirm:$false -ErrorAction Stop
            $didTitleSet = $true
            Write-Host "  [3] Title -> '$newTitle'. M365 license deassigns on next GBL cycle." -ForegroundColor Green
        }

        Write-Host ""
        Write-Host "Offboarded: $($user.SamAccountName)" -ForegroundColor Green
    } catch {
        $errMsg = $_.Exception.Message
        Write-Error "Failed mid-sequence: $_"
        Write-Warning "Re-run, completed steps will be skipped."
    } finally {
        if ($didDisable -or $didMove -or $didTitleSet -or $errMsg) {
            [PSCustomObject]@{
                Timestamp      = (Get-Date -Format 's')
                SamAccountName = $user.SamAccountName
                DisplayName    = $user.DisplayName
                OldTitle       = $user.Title
                Reason         = $Reason
                Disabled       = $didDisable
                Moved          = $didMove
                TitleSet       = $didTitleSet
                KeepTitle      = [bool]$KeepTitle
                FromOU         = $user.DistinguishedName
                ToOU           = if ($didMove) { $targetOu.DistinguishedName } else { $user.DistinguishedName }
                RunBy          = $env:USERNAME
                Error          = $errMsg
            } | Export-Csv -Path $csvPath -Append -NoTypeInformation
        }
    }
}
