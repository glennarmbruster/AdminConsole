#Requires -Version 5.1
<#
    Find-ResultSizeWarning.ps1

    Finds which line actually emits

        WARNING: There are more results available than are currently
        displayed. To view them, increase the value for the ResultSize
        parameter.

    Run it in the SAME PowerShell window that holds your Exchange Online
    session -- connections are per-process, and a fresh window would prove
    nothing because no Exchange cmdlet would run at all.

    Two stages:

      STAGE 1  Confirms the Common.ps1 sitting beside AdminConsole.ps1 is the
               fixed one. This is the likeliest answer and costs nothing.

      STAGE 2  Runs the console under Set-PSDebug -Trace 1, which echoes every
               statement as it executes, and prints the statements immediately
               BEFORE the warning. The last Exchange cmdlet in that window is
               the culprit.

    The console window will open during stage 2. Close it normally; the trace
    is written either way.

    Author  : Glenn Armbruster
    Created : 2026-08-03
#>

param(
    [string]$ConsolePath = (Join-Path $PSScriptRoot 'AdminConsole.ps1'),
    [string]$TracePath   = (Join-Path $env:TEMP 'AdminConsole-trace.txt')
)

$ErrorActionPreference = 'Continue'

function Write-Head {
    param([string]$Text)
    Write-Host ''
    Write-Host $Text -ForegroundColor Cyan
    Write-Host ('-' * $Text.Length) -ForegroundColor Cyan
}

# ---------------------------------------------------------------------------
# Stage 1 -- is the fixed Common.ps1 actually there?
# ---------------------------------------------------------------------------
Write-Head 'STAGE 1  Which Common.ps1 is in play'

if (-not (Test-Path -LiteralPath $ConsolePath)) {
    Write-Host "Could not find $ConsolePath" -ForegroundColor Red
    Write-Host "Pass -ConsolePath pointing at your AdminConsole.ps1 and run again."
    return
}

$dir    = Split-Path -Parent $ConsolePath
$common = Join-Path $dir 'Common.ps1'
Write-Host "Folder : $dir"

if (-not (Test-Path -LiteralPath $common)) {
    Write-Host 'Common.ps1 is MISSING from that folder.' -ForegroundColor Red
    return
}

$text = Get-Content -LiteralPath $common -Raw
$stamp = (Get-Item -LiteralPath $common).LastWriteTime

# Three markers that only exist in the fixed version.
$markers = [ordered]@{
    'Get-SessionOrganization function'   = 'function Get-SessionOrganization'
    'Test-TenantDomainMismatch function' = 'function Test-TenantDomainMismatch'
    'configured-domain early return'     = '(Get-ToolkitSettings).MailDomain'
}

Write-Host "Common.ps1 last written: $stamp"
Write-Host ''
$allPresent = $true
foreach ($k in $markers.Keys) {
    $present = $text.Contains($markers[$k])
    if (-not $present) { $allPresent = $false }
    $mark  = if ($present) { 'present' } else { 'MISSING' }
    $color = if ($present) { 'Green' }   else { 'Red' }
    Write-Host ("  {0,-36} {1}" -f $k, $mark) -ForegroundColor $color
}

if (-not $allPresent) {
    Write-Host ''
    Write-Host 'This is the OLD Common.ps1. That alone explains the warning --' -ForegroundColor Yellow
    Write-Host 'replace it and re-run this script before going any further.'    -ForegroundColor Yellow
    return
}

# The old unbounded call should be gone entirely.
$unbounded = Select-String -Path $common -SimpleMatch "Get-AcceptedDomain' -Parameters @{ ErrorAction"
if ($unbounded) {
    Write-Host ''
    Write-Host 'Found an UNBOUNDED Get-AcceptedDomain still in this file:' -ForegroundColor Yellow
    $unbounded | ForEach-Object { Write-Host "  line $($_.LineNumber): $($_.Line.Trim())" }
}

Write-Host ''
Write-Host 'Fixed Common.ps1 confirmed. The warning is coming from somewhere else.' -ForegroundColor Green

# ---------------------------------------------------------------------------
# Stage 2 -- trace it
# ---------------------------------------------------------------------------
Write-Head 'STAGE 2  Tracing execution'

Write-Host 'The console will open. Use it normally if you like, then CLOSE IT'
Write-Host 'to finish the trace.'
Write-Host ''
Write-Host "Trace file: $TracePath"
Write-Host ''
Read-Host 'Press Enter to start'

if (Test-Path -LiteralPath $TracePath) {
    Remove-Item -LiteralPath $TracePath -Force -ErrorAction SilentlyContinue
}

# Trace 1 echoes each statement before running it. Redirecting every stream
# (*>) keeps the warning and the trace interleaved in one file, which is the
# whole point -- the statements immediately above the warning are the source.
try {
    Set-PSDebug -Trace 1
    & $ConsolePath *> $TracePath
} finally {
    Set-PSDebug -Off
}

Write-Head 'RESULT'

if (-not (Test-Path -LiteralPath $TracePath)) {
    Write-Host 'No trace file was produced.' -ForegroundColor Red
    return
}

$hits = @(Select-String -Path $TracePath -Pattern 'more results available' -Context 25, 2)

if ($hits.Count -eq 0) {
    Write-Host 'The warning did NOT appear this run.' -ForegroundColor Green
    Write-Host ''
    Write-Host 'If you saw it before but not now, it was almost certainly the old'
    Write-Host 'Common.ps1 still being loaded from a different copy of the folder.'
    return
}

Write-Host ("Warning fired {0} time(s). Statements leading up to the FIRST one:" -f $hits.Count) -ForegroundColor Yellow
Write-Host ''

$first = $hits[0]
foreach ($line in $first.Context.PreContext) {
    # Trace lines look like: DEBUG:   123+     $x = Get-Something
    $clean = $line -replace '^DEBUG:\s*', ''
    if ($clean.Trim()) { Write-Host "    $clean" }
}
Write-Host ''
Write-Host "  >>> $($first.Line.Trim())" -ForegroundColor Yellow
Write-Host ''

# Name the likely cmdlet outright rather than making someone read the trace.
$cmdlets = @()
foreach ($line in $first.Context.PreContext) {
    foreach ($m in [regex]::Matches($line, '\b(?:Get|Set|New|Remove|Add|Search)-[A-Za-z]+\b')) {
        if ($m.Value -notmatch '^(Get-Command|Get-Content|Get-Item|Get-Date|Get-ChildItem|Get-Member|Get-Variable)$') {
            $cmdlets += $m.Value
        }
    }
}
if ($cmdlets.Count -gt 0) {
    Write-Host 'Cmdlets seen just before the warning (last is most likely):' -ForegroundColor Cyan
    $cmdlets | Select-Object -Last 6 | ForEach-Object { Write-Host "    $_" }
    Write-Host ''
    Write-Host 'Fix: add  ResultSize = ''Unlimited''  to that call.'
} else {
    Write-Host 'No cmdlet was visible in the preceding statements, which usually'
    Write-Host 'means the warning came from a MODULE rather than from this script'
    Write-Host '-- check your PowerShell profile and your connect script:'
    Write-Host ''
    Write-Host '    notepad $PROFILE'
}

Write-Host ''
Write-Host "Full trace kept at: $TracePath"
