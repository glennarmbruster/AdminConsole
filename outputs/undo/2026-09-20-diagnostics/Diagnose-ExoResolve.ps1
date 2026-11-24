#Requires -Version 5.1
# TOOLKIT-VARIANT: Dual
# TOOLKIT-HELP-RUNS: Exchange Online and on-premises Exchange
# TOOLKIT-HELP-BEGIN
# Shows how a recipient resolves in each place, for when a tool finds an object the operator cannot.
# Read-only.
# TOOLKIT-HELP-END
<#
    Diagnose-ExoResolve.ps1
    Which module is each Exchange cmdlet actually coming from, and does the
    toolkit's resolver pick the right one?

    RUN THIS IN THE SESSION THAT IS MISBEHAVING. The answer depends entirely on
    what is loaded in that process.

    WHY IT EXISTS: on 2026-08-23 a report returned UNKNOWN for 1,039 mailboxes
    because Get-EXOCommand rejected the genuine Exchange Online copy of
    Get-MailboxPermission -- it matched module names like tmpEXO_*, and the
    current module generates tmp_<random>.<random>. The fix resolves by
    elimination instead. But BOTH Exchange Online and on-premises implicit
    remoting generate tmp_* module names, so "not the on-prem module" has to be
    decided from what a module EXPORTS, never from what it is called.

    This prints the evidence rather than the conclusion.

    It reads. It changes nothing, and it calls no Exchange cmdlet that touches
    a mailbox.
#>

$ErrorActionPreference = 'Continue'
$root = $PSScriptRoot
if (-not $root) { $root = (Get-Location).Path }
. (Join-Path $root 'Common.ps1')

Write-Host ''
Write-Host '  EXCHANGE CMDLET RESOLUTION' -ForegroundColor White
Write-Host "  machine: $env:COMPUTERNAME   user: $env:USERDOMAIN\$env:USERNAME" -ForegroundColor DarkGray
Write-Host ''

# ---- every loaded module that owns Get-Mailbox ----
Write-Host '  MODULES EXPORTING Get-Mailbox' -ForegroundColor Cyan
$exchModules = @()
foreach ($m in @(Get-Module)) {
    $ec = $m.ExportedCommands
    if (-not $ec -or -not $ec.ContainsKey('Get-Mailbox')) { continue }
    $isOnPrem = $ec.ContainsKey('Get-RemoteMailbox')
    $exchModules += [PSCustomObject]@{ Module = $m; OnPrem = $isOnPrem }
    Write-Host ("      {0,-28} {1}  exports={2}" -f $m.Name,
        $(if ($isOnPrem) { 'ON-PREMISES (has Get-RemoteMailbox)' } else { 'cloud (no Get-RemoteMailbox)     ' }),
        $ec.Count)
}
if (@($exchModules).Count -eq 0) { Write-Host '      none -- no Exchange session in this process' -ForegroundColor Red }

Write-Host ''
Write-Host '  WHAT Get-ExchangeModuleMap CONCLUDES' -ForegroundColor Cyan
$map = Get-ExchangeModuleMap -Refresh
Write-Host ("      OnPrem : {0}" -f $(if ($map.OnPrem) { $map.OnPrem.Name } else { '(none identified)' }))
Write-Host ("      Cloud  : {0}" -f $(if ($map.Cloud)  { $map.Cloud.Name }  else { '(none identified)' }))
Write-Host ("      Both   : {0}" -f $map.Both)
if (-not $map.OnPrem -and @($exchModules | Where-Object { $_.OnPrem }).Count -gt 0) {
    Write-Host '      MISMATCH -- a module above looks on-premises but the map did not catch it.' -ForegroundColor Red
}

# ---- the cmdlets the reports depend on ----
Write-Host ''
Write-Host '  PER-CMDLET RESOLUTION' -ForegroundColor Cyan
foreach ($name in @('Get-Mailbox', 'Get-MailboxPermission', 'Get-RecipientPermission', 'Get-RemoteMailbox')) {
    $all = @(Get-Command -Name $name -All -ErrorAction SilentlyContinue)
    Write-Host ("      {0}" -f $name) -ForegroundColor White
    if (@($all).Count -eq 0) {
        Write-Host '          NOT LOADED in this session at all.' -ForegroundColor Red
        continue
    }
    foreach ($c in $all) {
        $src = "$($c.ModuleName)"
        if (-not $src) { $src = "$($c.Source)" }
        $tag = ''
        if ($map.OnPrem -and $src -eq "$($map.OnPrem.Name)") { $tag = '  <- ON-PREMISES' }
        if ($map.Cloud  -and $src -eq "$($map.Cloud.Name)")  { $tag = '  <- cloud' }
        Write-Host ("          copy in {0,-28}{1}" -f $src, $tag)
    }
    $picked = $null
    $err = ''
    try { $picked = Get-EXOCommand -Name $name } catch { $err = "$($_.Exception.Message)" }
    if ($picked) {
        $psrc = "$($picked.ModuleName)"
        if (-not $psrc) { $psrc = "$($picked.Source)" }
        $wrong = ($map.OnPrem -and $psrc -eq "$($map.OnPrem.Name)")
        Write-Host ("          RESOLVER PICKS: {0}{1}" -f $psrc, $(if ($wrong) { '   *** THAT IS THE ON-PREMISES COPY ***' } else { '' })) `
            -ForegroundColor $(if ($wrong) { 'Red' } else { 'Green' })
    } else {
        Write-Host ("          RESOLVER FAILS: {0}" -f $err.Substring(0, [Math]::Min(120, $err.Length))) -ForegroundColor Red
    }
}

Write-Host ''
Write-Host '  Send this whole output back. It says which module each cmdlet lives in' -ForegroundColor DarkGray
Write-Host '  and which one the resolver chose, which is the only thing that settles it.' -ForegroundColor DarkGray
Write-Host ''
