#Requires -Version 5.1
<#
    Compare-Toolkit.ps1
    Prove the shipped code is clean and that nothing has been duplicated.

    Author   : Glenn Armbruster
    Revised  : 2026-09-17

    WHY THIS EXISTS. Two folders that are meant to be identical drift, and they
    drift quietly: a fix lands in one, a comment is reworded in the other, and
    six weeks later a bug is reported in one install that cannot be reproduced
    in the other because they are not the same program. That happened -- a
    theme feature was gated on a config mode one install had and the other did
    not, so the code path was never exercised where it mattered, and it took a
    reported bug to find it.

    THE SHAPE, SINCE 2026-09-17. There is ONE copy of the shared code, in this
    folder. The other install no longer has a mirror of it: its folder holds
    ONLY the scripts that belong to that install alone -- its suppliers, its
    tracing, its bulk jobs -- plus its markdown. A change to shared code is
    made once, here, and staged for both installs (shipEX\ as-is, shipFH\ with
    .txt appended). There is no second copy to drift.

    WHAT IT CHECKS
      1. No .ps1 exists in BOTH folders. A script in the site-only folder that
         also exists here is a copy, and a copy is the drift this exists to
         stop. Shared code lives here and only here.
      2. Every name in -SiteOnly still exists somewhere. The list is the
         record of which scripts do not travel; a name that matches nothing
         is a script that was renamed or removed without updating the record.
      3. None of the words in -Banned appear in the code of EITHER folder.
         This folder ships to the other organisation, so its code has to be
         clean; and the site-only folder is that organisation's own, so it
         must not carry this one's names either.

    WHAT IT DELIBERATELY DOES NOT CHECK
      Config\, Logs\, Reports\, and anything under them. Those are per-install
      by design, and a tool that flagged them would be ignored within a week.
      Projects\ and probes\ too -- see Get-ToolkitScript for why.

    USAGE

        .\Compare-Toolkit.ps1 -Other ..\AdminConsoleFH
        .\Compare-Toolkit.ps1 -Other ..\AdminConsoleFH -Detail

    Exit code 0 when clean, 1 when not, so it can gate a packaging step.
#>

[CmdletBinding()]
param(
    # The other install's SITE-ONLY folder. The folder this script sits in is
    # the reference and holds all shared code.
    [Parameter(Mandatory)][string]$Other,

    # This folder. Defaults to wherever this script lives.
    [string]$Reference = $PSScriptRoot,

    # Scripts that belong to ONE install and are never staged for the other.
    # Matched by leaf name in EITHER folder. Kept as a record even though the
    # site-only folder now makes the split physical: it is what a reader
    # checks when deciding whether a script in THIS folder goes to shipFH.
    [string[]]$SiteOnly = @(
        # This install's own -- live here, never staged for the other.
        'Disable-User.ps1'          # staff termination, this directory only
        'NewUser.ps1'               # staff provisioning, this directory only
        'NotDeadYet.ps1'            # un-terminate, this directory only
        'FranUser_Office.ps1'       # office report for one directory
        'ManageOwnerGroups.ps1'     # owner groups, this directory only
        'Rename-OfficeGroups.ps1'   # office group rename, this directory only
        'EX-ServiceNowQueue.ps1'    # this organisation's ServiceNow queue
        'Compare-Toolkit.ps1'       # this file; lives with the reference

        # The other install's own -- live in ITS folder, never here.
        'AvePoint-Dashboard.ps1', 'AvePointQuery.ps1'
        'ProofpointPodQuery.ps1', 'ProofpointSerQuery.ps1', 'SendGridQuery.ps1'
        'Forget-ApiCredentials.ps1'
        'Connect-Teams.ps1', 'Connect-Security.ps1', 'Grant-ExternalAccess.ps1'
        'Set-PimTimer.ps1', 'Get-Queues.ps1', 'Search_Purge.ps1'
        'HybridMessageTrace.ps1', 'Unified-Trace.ps1', 'Unified-Trace-EndUser.ps1'
        'Unified-Trace-Mockup.ps1', 'Resource-Delegates.ps1', 'RenameUser.ps1'
        'Member_BulkADD_CSV.ps1', 'SMB_Inventory.ps1', 'dl_send_to-as.ps1'
        'new-OOF.ps1', 'onprem-mailtest.ps1'
    ),

    # Words that must not appear in code in either folder.
    [string[]]$Banned = @(
        'firsthorizon', 'franchise', 'fran', 'HQ', 'HQTS',
        'express', 'expresspros', 'expresspersonnel', 'SRG', 'ANZ'
    ),

    # List every finding rather than a count.
    [switch]$Detail
)

function Get-ToolkitScript {
    <#
        Every .ps1 a folder owns, keyed by its path relative to the root.

        Config, Logs and Reports are skipped wholesale -- see the header. So is
        DisabledTools\, which is where a retired script goes to stop being part
        of the toolkit.

        Projects\ AND probes\ ARE SKIPPED FOR THE SAME REASON, and it is worth
        writing down. Projects\ is migration work written against a particular
        cutover at a particular organisation -- it is a record of what was
        done, not part of the shipped toolkit, and it has no business being
        copied to another install. probes\ is throwaway diagnostics. Between
        them they are forty-odd files that would be reported on every run, for
        ever, and forty rows nobody reads is how the one row that matters gets
        skipped.
    #>
    param([Parameter(Mandatory)][string]$Root)
    $skip = @('config', 'logs', 'reports', 'disabledtools', 'one-off', 'markdown',
              'projects', 'probes', 'retryqueue')
    $out = @{}
    foreach ($f in @(Get-ChildItem -LiteralPath $Root -Filter '*.ps1' -File -Recurse -ErrorAction SilentlyContinue)) {
        $rel = $f.FullName.Substring($Root.Length).TrimStart('\')
        $top = ($rel -split '\\')[0].ToLowerInvariant()
        if ($rel -match '\\' -and $skip -contains $top) { continue }
        $out[$rel] = $f.FullName
    }
    return $out
}

# ---- resolve ------------------------------------------------------------------
foreach ($p in @($Reference, $Other)) {
    if (-not (Test-Path -LiteralPath $p)) {
        Write-Host "Folder not found: $p" -ForegroundColor Red
        exit 1
    }
}
$refRoot = (Resolve-Path -LiteralPath $Reference).Path
$othRoot = (Resolve-Path -LiteralPath $Other).Path

Write-Host ''
Write-Host "Reference (shared code) : $refRoot" -ForegroundColor Cyan
Write-Host "Site-only folder        : $othRoot" -ForegroundColor Cyan
Write-Host ''

$ref = Get-ToolkitScript -Root $refRoot
$oth = Get-ToolkitScript -Root $othRoot
$problems = 0

# ---- 1. nothing in both ---------------------------------------------------------
# Compared on LEAF name, not relative path: a copy of Common.ps1 dropped into a
# subfolder of the site-only install is still a second copy of Common.ps1.
$refLeaf = @{}
foreach ($k in $ref.Keys) { $refLeaf[([System.IO.Path]::GetFileName($k)).ToLowerInvariant()] = $k }
$dupes = @($oth.Keys | Where-Object { $refLeaf.ContainsKey(([System.IO.Path]::GetFileName($_)).ToLowerInvariant()) } | Sort-Object)
if ($dupes.Count -eq 0) {
    Write-Host ("  OK    no shared code in the site-only folder ({0} site-only script(s) there, {1} here)" -f $oth.Count, $ref.Count) -ForegroundColor Green
} else {
    Write-Host ("  COPY  {0} script(s) exist in BOTH folders -- shared code lives in the reference only" -f $dupes.Count) -ForegroundColor Red
    foreach ($d in $dupes) { Write-Host "          $d" -ForegroundColor Red }
    $problems += $dupes.Count
}

# ---- 2. the site-only record is current ----------------------------------------
$allLeaf = @{}
foreach ($k in @($ref.Keys) + @($oth.Keys)) { $allLeaf[([System.IO.Path]::GetFileName($k)).ToLowerInvariant()] = $true }
$stale = @($SiteOnly | Where-Object { -not $allLeaf.ContainsKey($_.ToLowerInvariant()) } | Sort-Object)
if ($stale.Count -eq 0) {
    Write-Host ("  OK    every -SiteOnly name still exists ({0} declared)" -f $SiteOnly.Count) -ForegroundColor Green
} else {
    Write-Host ("  STALE {0} -SiteOnly name(s) match no file -- renamed or removed without updating the list" -f $stale.Count) -ForegroundColor Yellow
    foreach ($s in $stale) { Write-Host "          $s" -ForegroundColor Yellow }
    $problems += $stale.Count
}
if ($Detail) {
    $undeclared = @($oth.Keys | Where-Object { $SiteOnly -notcontains ([System.IO.Path]::GetFileName($_)) } | Sort-Object)
    if ($undeclared.Count) {
        Write-Host '        in the site-only folder but not in -SiteOnly (add them so the record stays whole):' -ForegroundColor DarkGray
        foreach ($u in $undeclared) { Write-Host "          $u" -ForegroundColor DarkGray }
    }
}

# ---- 3. no organisation names in code, either folder ----------------------------
#
# WORD BOUNDARIES, NOT SUBSTRINGS. 'express' as a substring matches 'expression'
# and 'expressed', which appear in ordinary English comments throughout the
# toolkit; flagging those trains people to ignore this check.
$pattern = '(?i)\b(' + (($Banned | ForEach-Object { [regex]::Escape($_) }) -join '|') + ')\b'
$hits = @()
foreach ($set in @(@{ Label = 'reference'; Map = $ref }, @{ Label = 'site-only'; Map = $oth })) {
    foreach ($k in ($set.Map.Keys | Sort-Object)) {
        # This install's own tools may name this install; they never travel.
        if ($set.Label -eq 'reference' -and $SiteOnly -contains ([System.IO.Path]::GetFileName($k))) { continue }
        $n = 0
        try { $n = @(Select-String -LiteralPath $set.Map[$k] -Pattern $pattern -AllMatches -ErrorAction Stop).Count } catch { }
        if ($n -gt 0) { $hits += [pscustomobject]@{ Where = $set.Label; File = $k; Count = $n } }
    }
}
if ($hits.Count -eq 0) {
    Write-Host '  OK    no organisation names in shipped code' -ForegroundColor Green
} else {
    Write-Host ("  NAMES {0} file(s) carry a banned word" -f $hits.Count) -ForegroundColor Red
    foreach ($h in $hits) { Write-Host ("          [{0}] {1}  ({2})" -f $h.Where, $h.File, $h.Count) -ForegroundColor Red }
    $problems += $hits.Count
}

Write-Host ''
if ($problems -eq 0) {
    Write-Host 'Clean. Safe to ship.' -ForegroundColor Green
    exit 0
}
Write-Host ("$problems finding(s). Fix these before shipping.") -ForegroundColor Red
exit 1
