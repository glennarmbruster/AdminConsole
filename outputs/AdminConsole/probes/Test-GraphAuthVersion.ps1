#Requires -Version 5.1
# ---------------------------------------------------------------------------
<#
    Test-GraphAuthVersion.ps1

    Which installed version of Microsoft.Graph.Authentication actually loads.

    WHY A CHILD PROCESS PER VERSION. A failed assembly load POISONS the
    AppDomain: once Microsoft.Graph.Authentication.Core has been bound at the
    wrong version, every later Import-Module in that same window fails the same
    way, whatever version you ask for. Looping inside one PowerShell would
    therefore report "all versions broken" the moment the first one broke,
    which is a confident wrong answer.

    So each version is tested in a FRESH powershell.exe that imports exactly
    one version and exits. Slower, and the only way to get a trustworthy
    result.

    Changes nothing. Installs nothing. Removes nothing.

    Author   : Glenn Armbruster
    Revised  : 2026-08-12
#>

$ErrorActionPreference = 'Continue'

Write-Host ''
Write-Host 'Microsoft.Graph.Authentication -- which version loads?' -ForegroundColor Cyan
Write-Host ('-' * 72) -ForegroundColor DarkGray

$versions = @(Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication' -ErrorAction SilentlyContinue |
              Sort-Object Version -Descending)

if ($versions.Count -eq 0) {
    Write-Host '  None installed.' -ForegroundColor Red
    return
}

Write-Host "  $($versions.Count) version(s) installed. Testing newest first, one fresh process each."
Write-Host ''

# -NoProfile matters. A profile that imports Graph, or sets
# PSModuleAutoloadingPreference, would contaminate the very thing being
# measured -- the test has to start from nothing.
$psExe = (Get-Process -Id $PID).Path
if (-not $psExe) { $psExe = 'powershell.exe' }

$results = @()
foreach ($v in $versions) {
    $ver  = "$($v.Version)"
    $base = "$($v.ModuleBase)"

    # The child prints one token so the parent never has to interpret an
    # exception object across a process boundary. Anything else on stdout is
    # the error text, which is what you actually want to read.
    $cmd = @"
`$ErrorActionPreference = 'Stop'
try {
    Import-Module Microsoft.Graph.Authentication -RequiredVersion '$ver' -ErrorAction Stop
    if (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue) { 'LOADED' } else { 'NO-CMDLET' }
} catch {
    'FAILED: ' + `$_.Exception.Message
}
"@

    Write-Host ("  {0,-10} " -f $ver) -NoNewline
    $out = & $psExe -NoProfile -NonInteractive -Command $cmd 2>&1
    $txt = ($out | Out-String).Trim()

    if ($txt -match '^LOADED') {
        Write-Host 'LOADS' -ForegroundColor Green
        $results += [PSCustomObject]@{ Version = $ver; Ok = $true; Base = $base }
    } else {
        Write-Host 'fails' -ForegroundColor Red
        $first = ($txt -split "`n" | Where-Object { $_.Trim() } | Select-Object -First 1)
        if ($first) {
            $first = $first.Trim()
            if ($first.Length -gt 110) { $first = $first.Substring(0, 110) + '...' }
            Write-Host "             $first" -ForegroundColor DarkGray
        }
        $results += [PSCustomObject]@{ Version = $ver; Ok = $false; Base = $base }
    }
}

Write-Host ''
Write-Host 'VERDICT' -ForegroundColor Cyan
Write-Host ('-' * 72) -ForegroundColor DarkGray

$good = @($results | Where-Object { $_.Ok })
if ($good.Count -gt 0) {
    $pick = $good[0]
    Write-Host "  $($pick.Version) loads cleanly." -ForegroundColor Green
    Write-Host "    $($pick.Base)"
    Write-Host ''
    Write-Host '  Pin it for now, in a FRESH window, before anything else Graph-related:'
    Write-Host "    Import-Module Microsoft.Graph.Authentication -RequiredVersion $($pick.Version)" -ForegroundColor White
    Write-Host ''
    Write-Host '  That is a workaround, not a fix. The duplicates are still there and the'
    Write-Host '  next tool that imports without -RequiredVersion gets the highest version'
    Write-Host '  and the same failure.'
} else {
    Write-Host '  No installed version loads. The install needs cleaning, not pinning.' -ForegroundColor Red
}

Write-Host ''
Write-Host '  Versions found:'
foreach ($r in $results) {
    Write-Host ("    {0,-10} {1,-7} {2}" -f $r.Version, $(if ($r.Ok) { 'ok' } else { 'broken' }), $r.Base)
}
Write-Host ''
