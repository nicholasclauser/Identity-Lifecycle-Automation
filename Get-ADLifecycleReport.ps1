<#
.SYNOPSIS
  Read-only AD lifecycle report. Identifies inactive, never-used, and long-disabled
  user accounts. Produces a two-sheet Excel workbook (L2 user accounts, L3 service/admin
  accounts) and optionally uploads to SharePoint via PnP.PowerShell. Makes NO CHANGES to AD.

.DESCRIPTION
  - Searches on-prem AD (optionally within a Base DN) using RSAT ActiveDirectory.
  - Excludes nothing by default. Flags:
      * HasKeepNote (AD 'Description' matches keep pattern).
      * IsExcluded (membership in supplied exclusions CSV or AD group).
  - Computes DaysSinceLastLogon (from LastLogonDate with fallback to lastLogonTimestamp), DaysSinceCreated, DaysSincePwdSet.
  - "Disabled ≥120d candidates" uses a conservative heuristic:
      Enabled = $false AND (LastLogonDate ≤ Now-120d OR (LastLogonDate null AND WhenChanged ≤ Now-120d)).
    This is a best-effort proxy, AD does not store a native "DisabledOn" timestamp.
  - Outputs: CSVs + SummaryReport.txt + Summary.json + Transcript.log in a timestamped folder (optional).
  - NO AD CHANGES.

.PARAMETER BaseDN
  Optional search base (DN). If empty, defaults to the domain DN.

.PARAMETER OutputDir
  Root output directory. If omitted, uses .env OUTPUT_DIR or the current directory.

.PARAMETER MaxPreview
  Console preview rows per category (default 5). Can be sourced from .env MAX_PREVIEW.

.PARAMETER Timestamped
  If set, creates a yyyyMMdd-HHmm subfolder in OutputDir. Can be sourced from .env TIMESTAMPED=true/false.

.PARAMETER ExclusionsCsv
  Optional CSV of accounts to exclude from "Actionable" counts. Accepts columns:
  Identity | SamAccountName | DistinguishedName | UserPrincipalName | UPN | Mail | Name
  (any one present is enough).

.PARAMETER ExclusionsGroupDN
  Optional AD group DN whose (recursive) members should be excluded from "Actionable" counts
  (e.g., Service/VIP/Admin exceptions group).

.PARAMETER StaleListPath
  Optional external list (.xlsx or .csv) to compute overlap. If .xlsx is provided, the script uses the
  ImportExcel module if available; otherwise asks you to save as CSV and rerun.

.PARAMETER DaysInactive
  Threshold for "inactive (enabled)" users, default 90 days.

.PARAMETER DaysNeverUsed
  Threshold for "never-logged-in (enabled)" users by account age, default 30 days.

.PARAMETER DaysDisabledForDelete
  Threshold for "disabled" deletion candidates, default 120 days.

.PARAMETER KeepNotePattern
  Regex to detect "keep/exempt/hold" text in Description (default matches keep|retain|do not delete|legal hold|exempt).

.EXAMPLE
  .\Get-ADLifecycleReport.ps1 -BaseDN "DC=corp,DC=example,DC=com" -Timestamped

.NOTES
  Optimized for PowerShell 7 (uses WinPS compat to load AD module). Does not modify AD.
  Company-specific exclusion patterns are loaded from an external file (see .env
  AUTOEXCLUDE_PATTERNS_FILE) and a vendor email regex (.env VENDOR_EXCLUDE_DOMAINS).
#>

[CmdletBinding()]
param(
  [string]$BaseDN,
  [string]$OutputDir,
  [int]$MaxPreview = 5,
  [switch]$Timestamped,
  [string]$ExclusionsCsv,
  [string]$ExclusionsGroupDN,
  [string]$StaleListPath,
  [int]$DaysInactive = 90,
  [int]$DaysNeverUsed = 30,
  [int]$DaysDisabledForDelete = 120,
  [string]$KeepNotePattern = '(?i)\b(keep|retain|do not delete|legal hold|exempt)\b'
)

#region Helpers
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
      # remove balanced quotes
      if (($v.StartsWith('"') -and $v.EndsWith('"')) -or ($v.StartsWith("'") -and $v.EndsWith("'"))) {
        $v = $v.Substring(1, $v.Length - 2)
      }
      $map[$k] = $v
    }
  }
  return $map
}

function ConvertTo-Bool {
  param([string]$s)
  if ($null -eq $s) { return $false }
  $t = $s.Trim().ToLower()
  switch -regex ($t) {
    '^(true|1|yes|y)$' { return $true }
    default            { return $false }
  }
}

function New-OutputFolder {
  param([string]$Root, [switch]$UseTimestamp)
  # Fall back to the script's directory (not cwd) so scheduled tasks don't dump reports
  # into C:\Windows\System32. Override via -OutputDir or .env OUTPUT_DIR.
  if (-not $Root) { $Root = $PSScriptRoot }
  if (-not (Test-Path -LiteralPath $Root)) { New-Item -Path $Root -ItemType Directory -Force | Out-Null }
  if ($UseTimestamp) {
    $sub = (Get-Date).ToString('yyyyMMdd-HHmm')
    $full = Join-Path $Root $sub
    New-Item -Path $full -ItemType Directory -Force | Out-Null
    return $full
  } else {
    return $Root
  }
}

function Get-OUFromDN {
  param([string]$DistinguishedName)
  if (-not $DistinguishedName) { return $null }
  # Respect escaped commas in RDNs
  $parts = $DistinguishedName -split '(?<!\\),'
  $ouParts = $parts | Where-Object { $_ -like 'OU=*' -or $_ -like 'DC=*' }
  [string]::Join(',', $ouParts)
}

function ConvertTo-DateTime {
  param($v)
  if ($null -eq $v) { return $null }
  try { return [datetime]$v } catch { return $null }
}

function Build-IdentityIndex {
  param($objects, [string[]]$KeyProps = @('SamAccountName','UserPrincipalName','DistinguishedName','Mail','Name'))
  $set = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  foreach ($o in $objects) {
    foreach ($k in $KeyProps) {
      $val = $o.$k
      if ($val) { [void]$set.Add([string]$val) }
    }
  }
  return $set
}

function Get-AutoExcludePatterns {
  # Loads DisplayName auto-exclude regex patterns from an external file (one regex per line).
  # Lines starting with '#' are comments. Returns an empty array if the file is missing or unset.
  param([string]$Path)
  if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return @() }
  $patterns = @()
  foreach ($line in Get-Content -LiteralPath $Path) {
    $t = $line.Trim()
    if ($t -and -not $t.StartsWith('#')) { $patterns += $t }
  }
  return $patterns
}

function Import-StaleList {
  param([string]$Path)
  if (-not $Path) { return @() }
  if (-not (Test-Path -LiteralPath $Path)) {
    Write-Warning "StaleListPath not found: $Path"
    return @()
  }
  $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()
  if ($ext -eq '.csv') {
    return Import-Csv -LiteralPath $Path
  } elseif ($ext -eq '.xlsx') {
    $importExcel = Get-Module -ListAvailable -Name ImportExcel
    if ($null -ne $importExcel) {
      Import-Module ImportExcel -ErrorAction Stop
      return Import-Excel -Path $Path
    } else {
      Write-Warning "ImportExcel module not found. Save '$Path' as CSV and pass that to -StaleListPath for overlap."
      return @()
    }
  } else {
    Write-Warning "Unsupported stale list format: $ext. Use .csv or .xlsx."
    return @()
  }
}

function Write-CsvUtf8Bom {
  param(
    [Parameter(Mandatory=$true, ValueFromPipeline=$true)] $InputObject,
    [Parameter(Mandatory=$true)][string] $Path
  )
  begin { $all = [System.Collections.Generic.List[object]]::new() }
  process { $all.Add($InputObject) }
  end {
    $tmp = [IO.Path]::GetTempFileName()
    try {
      $all | Export-Csv -LiteralPath $tmp -NoTypeInformation -Encoding UTF8
      $utf8 = [System.Text.Encoding]::UTF8
      $raw = Get-Content -LiteralPath $tmp -Raw
      $bytes = $utf8.GetPreamble() + $utf8.GetBytes($raw)
      [IO.File]::WriteAllBytes($Path, $bytes)
    } finally {
      if (Test-Path $tmp) { Remove-Item $tmp -Force }
    }
  }
}

function Import-Exclusions {
  param([string]$CsvPath, [string]$GroupDN)
  $ids = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)

  if ($CsvPath -and (Test-Path -LiteralPath $CsvPath)) {
    try {
      $csv = Import-Csv -LiteralPath $CsvPath
      $setCsv = Build-IdentityIndex -objects $csv -KeyProps @('Identity','SamAccountName','UserPrincipalName','UPN','DistinguishedName','Mail','Name')
      foreach ($x in $setCsv) { [void]$ids.Add($x) }
      Write-Host "Loaded exclusions from CSV: $($setCsv.Count) identities"
    } catch {
      Write-Warning "Failed to import exclusions CSV: $($_.Exception.Message)"
    }
  }

  if ($GroupDN) {
    try {
      $members = Get-ADGroupMember -Identity $GroupDN -Recursive -ErrorAction Stop
      $userMembers = $members | Where-Object { $_.objectClass -eq 'user' }
      foreach ($m in $userMembers) {
        try {
          $u = Get-ADUser -Identity $m.DistinguishedName -Properties userPrincipalName,mail,samAccountName,distinguishedName
          foreach ($val in @($u.DistinguishedName, $u.SamAccountName, $u.UserPrincipalName, $u.Mail)) {
            if ($val) { [void]$ids.Add([string]$val) }
          }
        } catch {
          if ($m.DistinguishedName) { [void]$ids.Add([string]$m.DistinguishedName) }
          if ($m.SamAccountName)    { [void]$ids.Add([string]$m.SamAccountName)    }
        }
      }
      Write-Host "Loaded exclusions from AD group: $($members.Count) members"
    } catch {
      Write-Warning "Failed to load exclusions group '$GroupDN': $($_.Exception.Message)"
    }
  }

  return $ids
}
function Publish-SharePointReport {
  <#
  .SYNOPSIS
    Upload a run folder's artifacts to a SharePoint document library via PnP.PowerShell.

  .DESCRIPTION
    Cert-based app-only auth when Thumbprint is supplied, otherwise interactive
    browser login. Auto-creates the target folder structure under
    Shared Documents/<RootFolder>/<run-leaf>. Skips Transcript.log. Returns the
    SharePoint URL on success, $null on failure.
  #>
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][string]$RunFolder,
    [Parameter(Mandatory)][string]$SiteUrl,
    [Parameter(Mandatory)][string]$RootFolder,
    [Parameter(Mandatory)][string]$ClientId,
    [string]$TenantId,
    [string]$Thumbprint
  )

  if (-not (Test-Path -LiteralPath $RunFolder)) {
    Write-Warning "Publish-SharePointReport: run folder not found: $RunFolder"
    return $null
  }
  if ($null -eq (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Warning "PnP.PowerShell module not found. Install with: Install-Module PnP.PowerShell -Scope CurrentUser"
    return $null
  }

  try {
    Import-Module PnP.PowerShell -ErrorAction Stop

    if ($TenantId -and $Thumbprint) {
      Connect-PnPOnline -Url $SiteUrl `
        -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint `
        -ErrorAction Stop
    } else {
      Write-Host "SharePoint: launching interactive browser login (app: $ClientId)..."
      Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive -ErrorAction Stop
    }

    $subFolder = Split-Path $RunFolder -Leaf
    Add-PnPFolder -Name $RootFolder -Folder 'Shared Documents'              -ErrorAction SilentlyContinue | Out-Null
    Add-PnPFolder -Name $subFolder  -Folder "Shared Documents/$RootFolder"  -ErrorAction SilentlyContinue | Out-Null
    $targetFolder = "Shared Documents/$RootFolder/$subFolder"

    $uploadFiles = Get-ChildItem -LiteralPath $RunFolder -File | Where-Object Name -ne 'Transcript.log'
    foreach ($f in $uploadFiles) {
      Add-PnPFile -Path $f.FullName -Folder $targetFolder -ErrorAction Stop | Out-Null
      Write-Host "  Uploaded: $($f.Name)"
    }

    $resultUrl = "$SiteUrl/Shared Documents/$RootFolder/$subFolder"
    Write-Host "SharePoint upload successful ($($uploadFiles.Count) files): $resultUrl"

    Disconnect-PnPOnline -ErrorAction SilentlyContinue
    return $resultUrl
  } catch {
    Write-Warning "SharePoint upload failed: $($_.Exception.Message)"
    return $null
  }
}
#endregion Helpers

# --- ActiveDirectory module import (PS5.1 native; PS7 via WinPS compat) ---
try {
  if ($PSVersionTable.PSVersion.Major -ge 7) {
    Import-Module ActiveDirectory -UseWindowsPowerShell -ErrorAction Stop
  } else {
    Import-Module ActiveDirectory -ErrorAction Stop
  }
} catch {
  Write-Error "ActiveDirectory module not available. On PS7, ensure WinPS compatibility; on PS5.1, install RSAT. $($_.Exception.Message)"
  exit 1
}

# Load .env from the script's directory (NOT cwd, so scheduled tasks still find it).
$envPath = Join-Path -Path $PSScriptRoot -ChildPath ".env"
$envMap = Get-DotEnv -Path $envPath

# Apply .env defaults if not explicitly provided
if (-not $PSBoundParameters.ContainsKey('OutputDir') -and $envMap.ContainsKey('OUTPUT_DIR')) { $OutputDir = $envMap['OUTPUT_DIR'] }
if (-not $PSBoundParameters.ContainsKey('MaxPreview') -and $envMap.ContainsKey('MAX_PREVIEW')) {
  try { [int]$MaxPreview = $envMap['MAX_PREVIEW'] } catch { Write-Warning "Invalid MAX_PREVIEW in .env: '$($envMap['MAX_PREVIEW'])', using default $MaxPreview" }
}
if (-not $PSBoundParameters.ContainsKey('BaseDN') -and $envMap.ContainsKey('BASE_DN')) { $BaseDN = $envMap['BASE_DN'] }

# SharePoint upload settings from .env
$SharePointSiteUrl    = if ($envMap.ContainsKey('SharePointSiteUrl'))    { $envMap['SharePointSiteUrl'] }    else { $null }
$SharePointFolder     = if ($envMap.ContainsKey('SharePointFolder'))     { $envMap['SharePointFolder'] }     else { $null }
$SharePointClientId   = if ($envMap.ContainsKey('SharePointClientId'))   { $envMap['SharePointClientId'] }   else { $null }
$SharePointTenantId   = if ($envMap.ContainsKey('SharePointTenantId'))   { $envMap['SharePointTenantId'] }   else { $null }
$SharePointThumbprint = if ($envMap.ContainsKey('SharePointThumbprint')) { $envMap['SharePointThumbprint'] } else { $null }

# Environment-specific exclusion config. Keep company patterns out of source.
$vendorExcludeDomains   = if ($envMap.ContainsKey('VENDOR_EXCLUDE_DOMAINS'))   { $envMap['VENDOR_EXCLUDE_DOMAINS'] }   else { $null }
$autoExcludePatternsFile = if ($envMap.ContainsKey('AUTOEXCLUDE_PATTERNS_FILE')) { $envMap['AUTOEXCLUDE_PATTERNS_FILE'] } else { $null }
if (-not $autoExcludePatternsFile) {
  $defaultPath = Join-Path $PSScriptRoot 'autoexclude-patterns.txt'
  if (Test-Path -LiteralPath $defaultPath) { $autoExcludePatternsFile = $defaultPath }
}
$autoExcludePatterns = Get-AutoExcludePatterns -Path $autoExcludePatternsFile

# Default BaseDN to the domain DN if still empty
if (-not $BaseDN) {
  try { $BaseDN = (Get-ADDomain).DistinguishedName } catch {}
}

# Normalize -Timestamped (switch) including .env
$useTimestamp = $false
if ($PSBoundParameters.ContainsKey('Timestamped')) {
  $useTimestamp = [bool]$Timestamped
} elseif ($envMap.ContainsKey('TIMESTAMPED')) {
  $useTimestamp = ConvertTo-Bool $envMap['TIMESTAMPED']
}

# Prepare output folder and transcript
$outFolder = New-OutputFolder -Root $OutputDir -UseTimestamp:$useTimestamp
$transcriptPath = Join-Path $outFolder "Transcript.log"
Start-Transcript -Path $transcriptPath -Force | Out-Null
Write-Host "Output Folder: $(Split-Path $outFolder -Leaf)"
Write-Host "Invocation PSBoundParameters: $(( $PSBoundParameters.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" } ) -join '; ')"

# Load exclusions (optional)
$exclusionsIndex = Import-Exclusions -CsvPath $ExclusionsCsv -GroupDN $ExclusionsGroupDN
# Defensive: ensure we have a HashSet to call .Contains on even if loader returned $null
if (-not $exclusionsIndex) {
  $exclusionsIndex = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
}

# Time thresholds
$now = Get-Date
$cutInactive = $now.AddDays(-1 * $DaysInactive)
$cutNeverUsed = $now.AddDays(-1 * $DaysNeverUsed)
$cutDisabled  = $now.AddDays(-1 * $DaysDisabledForDelete)

# Query AD users
$props = @(
  'samAccountName','userPrincipalName','name','givenName','sn','displayName','mail','department',
  'enabled','whenCreated','whenChanged','lastLogonTimestamp','lastLogonDate','pwdLastSet','title',
  'description','distinguishedName','manager'
)
$searchParams = @{ Filter = '*' ; Properties = $props ; ResultPageSize = 2000 ; ResultSetSize = $null }
if ($BaseDN) { $searchParams['SearchBase'] = $BaseDN }

Write-Host "Querying Active Directory users..."
$adUsers = Get-ADUser @searchParams

# Build manager lookup (resolve DNs directly to avoid LDAP filter escaping issues)
$allMgrDns = $adUsers | Where-Object { $_.Manager } | Select-Object -ExpandProperty Manager -Unique
$managerMap = @{}
if ($allMgrDns) {
  foreach ($dn in $allMgrDns) {
    try {
      $mgr = Get-ADUser -Identity $dn -Properties displayName,distinguishedName
      if ($mgr) { $managerMap[$mgr.DistinguishedName] = $mgr.DisplayName }
    } catch {
      $managerMap[$dn] = $dn
    }
  }
}

# Project inventory with computed fields
$inventory = $adUsers | ForEach-Object {
  $dn = $_.DistinguishedName
  $ou = Get-OUFromDN -DistinguishedName $dn

  # Normalize deserialized date fields
  $llDate      = ConvertTo-DateTime $_.LastLogonDate
  $lltRaw      = $_.lastLogonTimestamp
  $llTimestamp = if ($lltRaw -and $lltRaw -ne 0) { [DateTime]::FromFileTime([Int64]$lltRaw) } else { $null }
  if (-not $llDate -and $llTimestamp) { $llDate = $llTimestamp }

  $whenCreated = ConvertTo-DateTime $_.WhenCreated
  $whenChanged = ConvertTo-DateTime $_.WhenChanged

  $daysSinceLogon   = if ($llDate)      { [int][math]::Round( ($now - $llDate).TotalDays ) } else { $null }
  $daysSinceCreated = if ($whenCreated) { [int][math]::Round( ($now - $whenCreated).TotalDays ) } else { $null }
  $pwdSet = if ($_.pwdLastSet -and $_.pwdLastSet -ne 0) { [DateTime]::FromFileTime([Int64]$_.pwdLastSet) } else { $null }
  $daysSincePwd = if ($pwdSet)          { [int][math]::Round( ($now - $pwdSet).TotalDays ) } else { $null }

  $hasKeep = $false
  if ($_.Description -and $KeepNotePattern) {
    $hasKeep = [regex]::IsMatch($_.Description, $KeepNotePattern)
  }
  # "DISABLE AFTER [date]" logic: veto if date is in the future, allow scoring if date has passed
  $disableByDate = $null
  if ($_.Description -match '(?i)disable\s+after\s+(.+?)(\s*[/|]|$)') {
    $disableAfterRaw = $Matches[1].Trim()
    $parsedDate = $null
    if ([datetime]::TryParse($disableAfterRaw, [ref]$parsedDate)) {
      $disableByDate = $parsedDate
      if ($parsedDate -gt $now) { $hasKeep = $true }  # future date, don't touch yet
      # past date, let it get scored normally (account should have been disabled)
    } else {
      $hasKeep = $true  # can't parse date, veto to be safe
    }
  }

  $isExcluded = $false
  foreach ($id in @($_.SamAccountName, $_.UserPrincipalName, $_.DistinguishedName, $_.Mail, $_.Name)) {
    if ($id -and $exclusionsIndex -and $exclusionsIndex.Contains([string]$id)) { $isExcluded = $true ; break }
  }

  $isPotentialService = $false
  if ($_.SamAccountName -match '^(svc_|svc-|service_)' -or $_.SamAccountName -match '(_svc|\.svc)$') { $isPotentialService = $true }
  $isPrivileged = $false
  if ($_.SamAccountName -match '(_adm|\.adm)$' -or $_.SamAccountName -match '^(adm-|admin_)') { $isPrivileged = $true }

  # Account type classification (used for L3 sheet)
  $accountType = 'User'
  if ($_.SamAccountName -match 'ea$')          { $accountType = 'Elevated Access' }
  elseif ($isPotentialService)                  { $accountType = 'Service Account' }
  elseif ($isPrivileged)                        { $accountType = 'Privileged/Admin' }
  elseif ($_.title -and $_.title.Trim() -match '(?i)^ISA$') { $accountType = 'IS Admin' }

  # Auto-exclude: computer accounts (SAM ends in $)
  if ($_.SamAccountName -match '\$$') { $isExcluded = $true }

  # Auto-exclude: EA (elevated access) suffix on SamAccountName
  if ($_.SamAccountName -match 'ea$') { $isExcluded = $true }

  # Auto-exclude: vendor/third-party accounts by email domain.
  # Pattern is loaded from .env VENDOR_EXCLUDE_DOMAINS (regex matching the email suffix).
  if ($vendorExcludeDomains -and $_.Mail -and $_.Mail -match $vendorExcludeDomains) { $isExcluded = $true }

  # Auto-exclude: DisplayName keyword patterns (service accounts, system accounts, shared mailboxes).
  # Patterns are loaded from an external file (one regex per line). See autoexclude-patterns.template.txt
  # for the expected format. Default location: ./autoexclude-patterns.txt (gitignored).
  if ($_.DisplayName -and $autoExcludePatterns -and $autoExcludePatterns.Count -gt 0) {
    foreach ($pat in $autoExcludePatterns) {
      if ($_.DisplayName -match $pat) { $isExcluded = $true; break }
    }
  }

  # Resolve manager display name defensively (avoid inline 'if' expression in hashtable)
  $managerName = $null
  if ($_.Manager -and $managerMap -and $managerMap.ContainsKey($_.Manager)) {
    $managerName = $managerMap[$_.Manager]
  }

  [pscustomobject]@{
    SamAccountName       = $_.SamAccountName
    UserPrincipalName    = $_.UserPrincipalName
    Name                 = $_.Name
    DisplayName          = $_.DisplayName
    GivenName            = $_.GivenName
    Surname              = $_.SN
    Mail                 = $_.Mail
    Title                = $_.title
    Department           = $_.Department
    Enabled              = [bool]$_.Enabled
    WhenCreated          = $whenCreated
    WhenChanged          = $whenChanged
    LastLogonDate        = $llDate
    DaysSinceLastLogon   = $daysSinceLogon
    DaysSinceCreated     = $daysSinceCreated
    PwdLastSetDate       = $pwdSet
    DaysSincePwdLastSet  = $daysSincePwd
    Description          = $_.Description
    DistinguishedName    = $dn
    OUPath               = $ou
    ManagerDN            = $_.Manager
    ManagerName          = $managerName
    HasKeepNote          = $hasKeep
    IsExcluded           = $isExcluded
    IsPotentialService   = $isPotentialService
    IsPrivileged         = $isPrivileged
    AccountType          = $accountType
    DisableBy            = $disableByDate
  }
}

# Optional: filter obvious system/builtin noise from reporting
$inventory = $inventory | Where-Object {
  $_.SamAccountName -ne 'krbtgt' -and
  $_.SamAccountName -notlike 'MSOL_*' -and
  $_.SamAccountName -notlike 'HealthMailbox*' -and
  $_.SamAccountName -notlike 'DiscoverySearchMailbox*'
}

# Split into L2 (real users) and L3 (service/admin/generic accounts)
$l2Inventory = $inventory | Where-Object { -not $_.IsExcluded -and -not $_.IsPotentialService -and -not $_.IsPrivileged }
$l3Inventory = $inventory | Where-Object { $_.IsExcluded -or $_.IsPotentialService -or $_.IsPrivileged }

# Define categories (L2 only)
$inactiveEnabled = $l2Inventory | Where-Object {
  $_.Enabled -eq $true -and $_.LastLogonDate -and $_.LastLogonDate -le $cutInactive
}
$neverUsedEnabled = $l2Inventory | Where-Object {
  $_.Enabled -eq $true -and (-not $_.LastLogonDate) -and $_.WhenCreated -le $cutNeverUsed
}
$disabledCandidates = $l2Inventory | Where-Object {
  $_.Enabled -eq $false -and (
    ($_.LastLogonDate -and $_.LastLogonDate -le $cutDisabled) -or
    ((-not $_.LastLogonDate) -and $_.WhenChanged -le $cutDisabled)
  )
}

# Add proposed action & actionable flags
function Add-Proposed { param($rows, [string]$proposed, [string]$reason)
  $rows | ForEach-Object {
    $_ | Add-Member -NotePropertyName 'ProposedAction' -NotePropertyValue $proposed -Force
    $_ | Add-Member -NotePropertyName 'ProposedReason' -NotePropertyValue $reason -Force
    $_ | Add-Member -NotePropertyName 'Actionable' -NotePropertyValue ((-not $_.HasKeepNote) -and (-not $_.IsExcluded)) -Force
    $_
  }
}
$inactiveEnabled    = Add-Proposed -rows $inactiveEnabled    -proposed "Disable+Move" -reason "Enabled & inactive ≥ $DaysInactive days"
$neverUsedEnabled   = Add-Proposed -rows $neverUsedEnabled   -proposed "Delete"       -reason "Enabled & never used ≥ $DaysNeverUsed days since created"
$disabledCandidates = Add-Proposed -rows $disabledCandidates -proposed "Delete"       -reason "Disabled & meets ≥ $DaysDisabledForDelete day rule (heuristic)"

# Overlap with external stale list (optional)
$overlap = @()
if ($StaleListPath) {
  $stale = Import-StaleList -Path $StaleListPath
  if ($stale.Count -gt 0) {
    $staleIdx = Build-IdentityIndex -objects $stale -KeyProps @('samaccountname','userprincipalname','upn','email','mail','name','distinguishedname')
    $overlap = $inventory | Where-Object {
      foreach ($id in @($_.SamAccountName, $_.UserPrincipalName, $_.DistinguishedName, $_.Mail, $_.Name)) {
        if ($id -and $staleIdx.Contains([string]$id)) { return $true }
      }
      return $false
    }
    $overlap | ForEach-Object {
      $_ | Add-Member -NotePropertyName 'InStaleList' -NotePropertyValue $true -Force
      $_ | Add-Member -NotePropertyName 'Actionable'  -NotePropertyValue ((-not $_.HasKeepNote) -and (-not $_.IsExcluded)) -Force
    } | Out-Null
  }
}

# Write outputs
$files = @{
  SummaryReport = Join-Path $outFolder "SummaryReport.txt"
}


# --- Readability and lightweight scoring (RiskScore) ---
function Get-AgeBucket([int]$days) {
  if ($null -eq $days) { return 'No data' }
  switch ($days) {
    {$_ -le 30}  { '0–30' ; break }
    {$_ -le 60}  { '31–60'; break }
    {$_ -le 90}  { '61–90'; break }
    {$_ -le 180} { '91–180'; break }
    {$_ -le 365} { '181–365'; break }
    default      { '>365' }
  }
}

# Titles licensed by design but that sign in rarely. Set these for your org.
$rareLoginTitlePatterns = @('(?i)^fieldstaff')

$inventory | ForEach-Object {
  $now = Get-Date
  $src = if ($_.LastLogonDate) { 'LastLogonDate/LLT' } else { 'null' }
  $_ | Add-Member -NotePropertyName 'LastLogonSource' -NotePropertyValue $src -Force
  $_ | Add-Member -NotePropertyName 'LastLogon_AgeBucket' -NotePropertyValue (Get-AgeBucket $_.DaysSinceLastLogon) -Force
  $_ | Add-Member -NotePropertyName 'Pwd_AgeBucket' -NotePropertyValue (Get-AgeBucket $_.DaysSincePwdLastSet) -Force

  # scoring per recommended AD-only rubric (simple, transparent)
  $score = 0
  $reasons = @()

  if ($_.HasKeepNote -or $_.IsExcluded) {
    $score = -100
    $reasons += 'Veto: Keep/Excluded'
  } else {
    # --- Inactivity (graduated: max +5) ---
    if ($_.Enabled -and $_.DaysSinceLastLogon -ne $null) {
      if     ($_.DaysSinceLastLogon -ge 365) { $score += 5; $reasons += 'Inactive ≥365d' }
      elseif ($_.DaysSinceLastLogon -ge 180) { $score += 4; $reasons += 'Inactive ≥180d' }
      elseif ($_.DaysSinceLastLogon -ge 90)  { $score += 2; $reasons += 'Inactive ≥90d' }
    }

    # --- Never used (graduated: max +5) ---
    if ($_.Enabled -and (-not $_.LastLogonDate)) {
      if     ($_.DaysSinceCreated -ge 180) { $score += 5; $reasons += 'Never used (≥180d)' }
      elseif ($_.DaysSinceCreated -ge 90)  { $score += 4; $reasons += 'Never used (≥90d)' }
      elseif ($_.DaysSinceCreated -ge 60)  { $score += 3; $reasons += 'Never used (≥60d)' }
      elseif ($_.DaysSinceCreated -ge 30)  { $score += 1; $reasons += 'Never used (≥30d)' }
    }

    # --- Disabled duration (graduated: max +5) ---
    if (-not $_.Enabled -and $_.WhenChanged) {
      $disabledDays = ($now - $_.WhenChanged).Days
      if     ($disabledDays -ge 365) { $score += 5; $reasons += 'Disabled ≥365d' }
      elseif ($disabledDays -ge 180) { $score += 4; $reasons += 'Disabled ≥180d' }
      elseif ($disabledDays -ge 120) { $score += 2; $reasons += 'Disabled ≥120d' }
    }

    # --- Password staleness (graduated: max +2) ---
    if ($_.DaysSincePwdLastSet -ne $null) {
      if     ($_.DaysSincePwdLastSet -ge 730) { $score += 2; $reasons += 'Pwd old (≥2yr)' }
      elseif ($_.DaysSincePwdLastSet -ge 365) { $score += 1; $reasons += 'Pwd old (≥1yr)' }
    }

    # --- Missing attributes ---
    if (-not $_.ManagerDN) { $score += 3; $reasons += 'No manager (termination blind spot)' }
    if (-not $_.Department) { $score += 1; $reasons += 'No department' }
    if (-not $_.Title)      { $score += 1; $reasons += 'No job title' }

    # --- External signals ---
    if ($_.PSObject.Properties.Name -contains 'InStaleList' -and $_.InStaleList) { $score += 2; $reasons += 'In requester stale list' }
    if ($_.IsPotentialService -or $_.IsPrivileged) { $score -= 2; $reasons += 'Service/Privileged pattern' }

    # --- Role-aware adjustments ---
    # Two distinct cases that look similar but should score in opposite directions:
    #   (1) Rare-login titles (e.g. field roles licensed by design): score DOWN.
    #       These users are expected to sign in rarely. Inactivity is not staleness.
    #   (2) Roles your org has decided should not hold a paid seat at all:
    #       score UP. If they have a license and rarely use it, it's reclaim-eligible.
    # Adjust both pattern lists for your environment.
    if ($_.Title) {
      foreach ($pat in $rareLoginTitlePatterns) {
        if ($_.Title -match $pat) { $score -= 1; $reasons += 'Role: Rare-login (licensed by design)'; break }
      }
      # example no-license titles; replace with the ones that apply in your org
      if ($_.Title -match '(?i)(fieldrole|seasonal|contractor)') {
        $score += 2; $reasons += 'Role: no license expected'
      }
    }
  }

  $_ | Add-Member -NotePropertyName 'RiskScore' -NotePropertyValue $score -Force
  $_ | Add-Member -NotePropertyName 'TopReasons'  -NotePropertyValue ($reasons -join '; ') -Force
}

# Build review set (sorted by score), actionable candidates only
$reviewCols = @(
  'RiskScore','TopReasons',
  'SamAccountName','DisplayName','Title','Department','ManagerName',
  'Description','DisableBy',
  'Enabled','LastLogonDate','WhenCreated','WhenChanged',
  'LastLogon_AgeBucket','Pwd_AgeBucket','OUPath'
)

$review = ($inactiveEnabled + $neverUsedEnabled + $disabledCandidates) |
  Where-Object { $_.Actionable -and $_.RiskScore -ge 1 } |
  Sort-Object -Property @{e='RiskScore';Descending=$true}, @{e='Department';Descending=$false}, @{e='ManagerName';Descending=$false} |
  Select-Object $reviewCols

# XLSX output
$files['ReviewXlsx'] = Join-Path $outFolder 'AD_Lifecycle_Report.xlsx'

# Two-sheet workbook: L2_Users (scored user accounts) + L3_SpecialAccounts (service/admin/generic)
$importExcel = Get-Module -ListAvailable -Name ImportExcel
if ($null -ne $importExcel) {
  Import-Module ImportExcel -ErrorAction SilentlyContinue
  $wb = $files['ReviewXlsx']

  # Sheet 1: L2, all scored user candidates
  $review | Export-Excel -Path $wb -WorksheetName 'L2_Users' -AutoSize -ClearSheet -AutoFilter -FreezeTopRow -TableName 'L2Table' -Verbose:$false

  # Sheet 2: L3, special accounts (service/admin/EA/generic), sorted by inactivity
  $l3Cols = @(
    'AccountType','SamAccountName','DisplayName','Title','Department','ManagerName',
    'Description','Enabled','LastLogonDate','DaysSinceLastLogon',
    'WhenCreated','WhenChanged','RiskScore','OUPath'
  )
  $l3Inventory |
    Sort-Object -Property @{e='DaysSinceLastLogon';Descending=$true}, @{e='AccountType';Descending=$false} |
    Select-Object $l3Cols |
    Export-Excel -Path $wb -WorksheetName 'L3_SpecialAccounts' -AutoSize -AutoFilter -FreezeTopRow -TableName 'L3Table' -Verbose:$false
}

# Console previews (Top-N only)
function Show-Preview($title, $data) {
  Write-Host ""
  Write-Host "== $title =="
  Write-Host "Count: $($data.Count)"
  if ($data.Count -gt 0 -and $MaxPreview -gt 0) {
    $data | Select-Object -First $MaxPreview |
      Format-Table SamAccountName,Department,Enabled,LastLogonDate,WhenCreated,DisableBy -AutoSize |
      Out-String | Write-Host
  }
}
Show-Preview -title "Inactive (Enabled) ≥ $DaysInactive days" -data $inactiveEnabled
Show-Preview -title "Never-Logged-In (Enabled) ≥ $DaysNeverUsed days" -data $neverUsedEnabled
Show-Preview -title "Disabled ≥ $DaysDisabledForDelete days (candidates)" -data $disabledCandidates

# Summary TXT
$summary = @()
$summary += "AD Lifecycle Report"
$summary += "Generated: $now"
$summary += ""
$summary += "Inactive (Enabled) >= $DaysInactive d:       $($inactiveEnabled.Count)"
$summary += "Never-Logged-In (Enabled) >= $DaysNeverUsed d: $($neverUsedEnabled.Count)"
$summary += "Disabled >= $DaysDisabledForDelete d (candidates):  $($disabledCandidates.Count)"
$summary += "Flagged for review (RiskScore >= 1):          $($review.Count)"
$summary += ""
$summary += "Report: AD_Lifecycle_Report.xlsx"
$summary += "No changes made to Active Directory."
($summary -join [Environment]::NewLine) | Out-File -LiteralPath $files.SummaryReport -Encoding UTF8 -Force

Write-Host ""
Write-Host "Summary: $([IO.Path]::GetFileName($files.SummaryReport))"
Write-Host "Report:  $([IO.Path]::GetFileName($files.ReviewXlsx))"

# SharePoint upload (cert-based or interactive)
$spUploadResult = $null
if ($SharePointSiteUrl) {
  if (-not $SharePointClientId) {
    Write-Warning "SharePoint: skipping upload, SharePointClientId not set. Add it to .env or pass -SharePointClientId."
  } else {
    $spUploadResult = Publish-SharePointReport `
      -RunFolder  $outFolder `
      -SiteUrl    $SharePointSiteUrl `
      -RootFolder $SharePointFolder `
      -ClientId   $SharePointClientId `
      -TenantId   $SharePointTenantId `
      -Thumbprint $SharePointThumbprint
  }
}

Stop-Transcript | Out-Null

# Return a lightweight object for callers/pipelines
[pscustomobject]@{
  OutputFolder  = $outFolder
  Report        = $files.ReviewXlsx
  SummaryReport = $files.SummaryReport
}