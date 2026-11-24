#Requires -Version 5.1
# TOOLKIT-HELP-RUNS: nothing. It reads and rewrites this toolkit's own log files.
# TOOLKIT-HELP-BEGIN
# ONE-TIME. Merges the old two-column audit log into the current format so the
# toolkit has a single history instead of two partial ones.
#
#   .\Convert-LogHistory.ps1
#
# Three tools used to write their own log -- Task and DateTime, nothing else --
# to Logs\CompletedTasks.csv, while everything else wrote the full format to
# Logs\CompletedTasks-v2.csv. Two files, two schemas, and any count taken from
# either one was a statement about only part of the toolkit.
#
# This reads both, converts the old rows, and writes a single CompletedTasks.csv
# in the current format. Both originals are backed up first and neither is
# deleted.
#
# THE OLD ROWS HAVE NO RESULT, so Result is left BLANK for them rather than
# filled in as Success. Those tools only ever logged completions, so the
# information does not exist -- and a blank that reads as "unknown" is honest
# where a Success would be a made-up fact sitting in an audit log.
#
# Run it once, after taking the updated tool scripts. Running it again is safe:
# rows already in the current format are carried across unchanged.
# TOOLKIT-HELP-END
<#
    Convert-LogHistory.ps1
    One-time merge of the v1 and v2 completed-task logs.

    Reasoning: ScriptNotes/Versioning.notes.md (logging section)

    Author   : Glenn Armbruster
    Created  : 2026-08-07
#>

# THE TOOLKIT ROOT IS ONE LEVEL UP. This script lives in 'Diagnostic tools\\'
# as of 2026-09-20; Common.ps1 did not move with it.
#
# THE FALLBACK IS NOT DECORATION. If somebody copies this script back into the
# toolkit root -- which is exactly what a person does when a diagnostic will
# not run -- the parent would then be the folder ABOVE the toolkit and
# Common.ps1 would not be there. So the parent is tried first and, when
# Common.ps1 is not beside it, this script's own folder is used instead. It
# works from either place without being told which.
$here = $(if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path })
$root = Split-Path -Path $here -Parent
if (-not (Test-Path -LiteralPath (Join-Path $root 'Common.ps1'))) { $root = $here }
. (Join-Path $root 'Common.ps1')

# NO Test-ConsoleLaunch GUARD ANY MORE, AND THAT IS A DELIBERATE TRADE.
#
# This script moved out of the console's button list into 'Diagnostic tools\\'
# on 2026-09-20. The guard returns unless the process was started BY the
# console -- so with no console button left to start it, the guard would have
# made the script unreachable from anywhere. It had to go or the move had to.
#
# WHAT THAT COSTS: this is the only one of the four diagnostics that WRITES.
# It rewrites the toolkit's own log files, once, to merge the old two-column
# audit history into the current format. Anyone who opens the subfolder and
# double-clicks it now runs it, where before the console turned them back.
# It is idempotent on an already-converted log, and it is in a folder nobody
# browses by accident -- but it is no longer behind a door.

$folder = Get-LogPath
if (-not (Test-Path -LiteralPath $folder)) {
    Write-Host "No Logs folder at $folder -- nothing to convert." -ForegroundColor Yellow
    return
}

$primary  = Join-Path $folder 'CompletedTasks.csv'
$fallback = Join-Path $folder 'CompletedTasks-v2.csv'
$stamp    = (Get-Date -Format 'yyyy-MM-dd_HHmm')

Write-Host "Log folder: $folder"
Write-Host ''

# ---------------------------------------------------------------------------
# Read what exists
# ---------------------------------------------------------------------------
function Get-HeaderLine {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try { return (Get-Content -LiteralPath $Path -TotalCount 1 -ErrorAction Stop) } catch { return $null }
}

$primaryHeader = Get-HeaderLine -Path $primary
$isPrimaryV2   = ("$primaryHeader".Trim() -eq $script:LogHeaderLine)

$v1Rows = @()
$v2Rows = @()

if ($primaryHeader) {
    if ($isPrimaryV2) {
        $v2Rows += @(Import-Csv -LiteralPath $primary)
        Write-Host "CompletedTasks.csv     : already current format, $($v2Rows.Count) row(s)"
    } else {
        $v1Rows = @(Import-Csv -LiteralPath $primary)
        Write-Host "CompletedTasks.csv     : OLD format, $($v1Rows.Count) row(s)" -ForegroundColor Yellow
    }
} else {
    Write-Host "CompletedTasks.csv     : not present"
}

if (Test-Path -LiteralPath $fallback) {
    $f = @(Import-Csv -LiteralPath $fallback)
    $v2Rows += $f
    Write-Host "CompletedTasks-v2.csv  : current format, $($f.Count) row(s)"
} else {
    Write-Host "CompletedTasks-v2.csv  : not present"
}

if ($v1Rows.Count -eq 0 -and $v2Rows.Count -eq 0) {
    Write-Host ''
    Write-Host 'Nothing to merge.' -ForegroundColor Yellow
    return
}

if ($v1Rows.Count -eq 0 -and -not (Test-Path -LiteralPath $fallback)) {
    Write-Host ''
    Write-Host 'Already a single log in the current format. Nothing to do.' -ForegroundColor Green
    return
}

# ---------------------------------------------------------------------------
# Back up BEFORE writing anything
# ---------------------------------------------------------------------------
foreach ($f in @($primary, $fallback)) {
    if (Test-Path -LiteralPath $f) {
        $bak = "$f.$stamp.bak"
        Copy-Item -LiteralPath $f -Destination $bak -Force
        Write-Host "Backed up -> $(Split-Path -Leaf $bak)"
    }
}

# ---------------------------------------------------------------------------
# Convert the old rows
# ---------------------------------------------------------------------------
$converted = foreach ($r in $v1Rows) {
    $task = "$($r.Task)"
    $parsed = Split-LegacyTask -Task $task

    # v1 wrote "yyyy-MM-dd HHmm". Parsed back to a real date so the dashboard can
    # filter on it. The UTC offset was never recorded, so the machine's current
    # offset is applied -- which is an assumption, and the only one made here.
    $ts = ''
    $dt = [datetime]::MinValue
    if ([datetime]::TryParseExact("$($r.DateTime)", 'yyyy-MM-dd HHmm',
            [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::None, [ref]$dt)) {
        $ts = ([DateTimeOffset]::new($dt, [DateTimeOffset]::Now.Offset)).ToString('yyyy-MM-ddTHH:mm:sszzz')
    } else {
        $ts = "$($r.DateTime)"
    }

    [PSCustomObject]@{
        Timestamp   = $ts
        Operator    = ''
        Machine     = ''
        Script      = ''
        Action      = $parsed.Action
        Target      = $parsed.Target
        # Deliberately blank. See the help block: these tools logged completions
        # only, so "did it succeed" was never recorded. Writing Success here
        # would invent an audit fact.
        Result      = ''
        Detail      = $task
        RunId       = ''
        Environment = $parsed.Environment
    }
}

Write-Host ''
Write-Host "Converted $($converted.Count) old row(s)."
if ($converted.Count -gt 0) {
    $withAction = @($converted | Where-Object { $_.Action }).Count
    $withTarget = @($converted | Where-Object { $_.Target }).Count
    Write-Host "  Action recovered from the sentence : $withAction of $($converted.Count)"
    Write-Host "  Target recovered from the sentence : $withTarget of $($converted.Count)"
    Write-Host "  Result: blank on all of them -- it was never recorded." -ForegroundColor Yellow
}

# ---------------------------------------------------------------------------
# Merge and write
# ---------------------------------------------------------------------------
$all = @($converted) + @($v2Rows) |
       Select-Object $script:LogColumns |
       Sort-Object {
           $d = [datetime]::MinValue
           if ([datetime]::TryParse("$($_.Timestamp)", [ref]$d)) { $d } else { [datetime]::MinValue }
       }

$all | Export-Csv -LiteralPath $primary -NoTypeInformation -Encoding UTF8

Write-Host ''
Write-Host "Wrote $($all.Count) row(s) to $primary" -ForegroundColor Green

if (Test-Path -LiteralPath $fallback) {
    $retired = Join-Path $folder "CompletedTasks-v2.merged-$stamp.csv"
    Move-Item -LiteralPath $fallback -Destination $retired -Force
    Write-Host "Retired the split file -> $(Split-Path -Leaf $retired)"
    Write-Host 'Common.ps1 will now write to CompletedTasks.csv again.' -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# Verify what was just written is readable as the current format
# ---------------------------------------------------------------------------
Write-Host ''
Write-Host 'Verifying...'
$head = Get-HeaderLine -Path $primary
if ("$head".Trim() -ne $script:LogHeaderLine) {
    Write-Host '  HEADER MISMATCH after writing. The backups above are intact.' -ForegroundColor Red
    Write-Host "  expected: $script:LogHeaderLine" -ForegroundColor Red
    Write-Host "  got     : $head" -ForegroundColor Red
    return
}
$check = @(Import-Csv -LiteralPath $primary)
if ($check.Count -ne $all.Count) {
    Write-Host "  ROW COUNT MISMATCH: wrote $($all.Count), read back $($check.Count)." -ForegroundColor Red
    return
}
$blankResult = @($check | Where-Object { -not $_.Result }).Count
Write-Host "  OK -- $($check.Count) rows, header correct." -ForegroundColor Green
Write-Host "  $blankResult row(s) have no Result. Those are the converted ones;"
Write-Host "  anything reporting on results should show them as Unknown, not Success."
