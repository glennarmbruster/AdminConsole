#Requires -Version 5.1
# TOOLKIT-VARIANT: OnPremAD
# TOOLKIT-HELP-RUNS: Active Directory
# TOOLKIT-HELP-BEGIN
# Reports why the AD capability gate is open or closed on this machine: module present, domain reachable, rights held.
# Read-only.
# TOOLKIT-HELP-END
<#
    Diagnose-AdGate.ps1
    Why does the console say Active Directory is not connected?

    RUN THIS ON THE MACHINE THE CONSOLE RUNS ON. That is the whole point: the
    answer depends on what that process can see, and a working Get-ADUser in a
    different window on a different box is not evidence about this one.

    Test-AdAvailable has four gates and returns a single sentence. When the
    sentence is wrong -- and on 2026-08-23 it was, telling an operator there
    were no directories when there was one with an empty Domain -- there is no
    way to tell which gate actually fired. This walks them in order and shows
    the value each one saw.

    It reads. It changes nothing.
#>

$ErrorActionPreference = 'Continue'

$root = $PSScriptRoot
if (-not $root) { $root = (Get-Location).Path }
. (Join-Path $root 'Common.ps1')

function Line { param($n, $label, $value, $ok)
    $mark = if ($ok -eq $true) { 'PASS' } elseif ($ok -eq $false) { 'FAIL' } else { '    ' }
    $colour = if ($ok -eq $true) { 'Green' } elseif ($ok -eq $false) { 'Red' } else { 'Gray' }
    Write-Host ("  {0}  gate {1}  {2,-34} {3}" -f $mark, $n, $label, $value) -ForegroundColor $colour
}

Write-Host ''
Write-Host '  AD GATE DIAGNOSTIC' -ForegroundColor White
Write-Host "  machine : $env:COMPUTERNAME" -ForegroundColor DarkGray
Write-Host "  process : PowerShell $($PSVersionTable.PSVersion), 64-bit=$([Environment]::Is64BitProcess)" -ForegroundColor DarkGray
Write-Host "  user    : $env:USERDOMAIN\$env:USERNAME" -ForegroundColor DarkGray
Write-Host ''

# ---- gate 1 ----
$cap = $false
try { $null = Get-TenantConfig -Refresh } catch { }
try { $cap = Test-TenantCapability -Name 'OnPremAD' } catch { }
Line 1 'Capabilities.OnPremAD' $cap $cap
if (-not $cap) {
    Write-Host ''
    Write-Host '  FIX: Tenant Manager -> Capabilities -> tick OnPremAD.' -ForegroundColor Yellow
    Write-Host ''
    return
}

# ---- gate 2 ----
#
# -Refresh IS NOT OPTIONAL HERE, and leaving it out wasted a round trip on
# 2026-08-23. Get-TenantConfig caches in $global:TenantConfig for the life of
# the process, and this is routinely run from a long-lived Exchange Management
# Shell where the console has already loaded the OLD file. The operator copies
# a corrected Tenant.json, re-runs, and is shown the stale values with total
# confidence.
#
# A diagnostic that reads a cache is not a diagnostic.
$cfgPath = ''
try { $cfgPath = Get-TenantConfigPath } catch { }
Write-Host ("        reading: {0}" -f $(if ($cfgPath) { $cfgPath } else { '(path unknown)' })) -ForegroundColor DarkGray
if ($cfgPath -and (Test-Path -LiteralPath $cfgPath)) {
    Write-Host ("        last written: {0}" -f (Get-Item -LiteralPath $cfgPath).LastWriteTime) -ForegroundColor DarkGray
} else {
    Write-Host '        THAT FILE DOES NOT EXIST -- everything below is defaults.' -ForegroundColor Red
}

$dirs = @()
try { $dirs = @((Get-TenantConfig -Refresh).Directories) } catch { }
Line 2 'directories in Tenant.json' "$(@($dirs).Count)" $(@($dirs).Count -gt 0)
foreach ($d in $dirs) {
    $dcs = @($d.DomainControllers)
    Write-Host ("        {0,-10} Domain='{1}'  DomainControllers={2}" -f "$($d.Id)", "$($d.Domain)", $dcs.Count) -ForegroundColor DarkGray
}
$server = ''
try { $server = "$(Get-TenantDirectoryServer)" } catch { }
Line 2 'server it will bind to' $(if ($server) { "'$server'" } else { '(nothing)' }) ([bool]$server)
if (-not $server) {
    Write-Host ''
    Write-Host '  FIX: Tenant Manager -> Directories -> set Domain on the default directory.' -ForegroundColor Yellow
    Write-Host '       Leave DomainControllers empty -- that is correct, it lets AD find a DC.' -ForegroundColor Yellow
    Write-Host ''
    return
}

# ---- gate 3 ----
$cmd = Get-Command -Name 'Get-ADUser' -ErrorAction SilentlyContinue
Line 3 'Get-ADUser resolves' $(if ($cmd) { "yes -- $($cmd.Source)" } else { 'NO' }) ([bool]$cmd)
$mod = Get-Module ActiveDirectory
if ($mod) { Write-Host ("        module {0}  v{1}  {2}" -f $mod.ModuleType, $mod.Version, $mod.Path) -ForegroundColor DarkGray }
else {
    $avail = @(Get-Module -ListAvailable ActiveDirectory)
    Write-Host ("        not loaded; {0} available version(s) on this machine" -f $avail.Count) -ForegroundColor DarkGray
}
if (-not $cmd) {
    Write-Host ''
    Write-Host '  FIX: the ActiveDirectory module is not present IN THIS PROCESS.' -ForegroundColor Yellow
    Write-Host '       If RSAT is installed here, try:  Import-Module ActiveDirectory' -ForegroundColor Yellow
    Write-Host '       If that works, the console needs the same import -- tell me and I will add it.' -ForegroundColor Yellow
    Write-Host '       A working Get-ADUser in a window on ANOTHER machine does not count.' -ForegroundColor Yellow
    Write-Host ''
    return
}

# ---- gate 4 ----
$ok4 = $false
$detail = ''
try {
    $rootDse = Get-ADRootDSE -Server $server -ErrorAction Stop
    $ok4 = $true
    $detail = "$($rootDse.defaultNamingContext)"
} catch {
    $detail = "$($_.Exception.Message)"
}
Line 4 "Get-ADRootDSE -Server $server" $detail $ok4

Write-Host ''
$final = Test-AdAvailable
if ($final.Available) {
    Write-Host '  Test-AdAvailable says CONNECTED. The console will show it.' -ForegroundColor Green
    Write-Host '  If the console still shows disconnected, it is holding a cached state --' -ForegroundColor DarkGray
    Write-Host '  close every console window and reopen one.' -ForegroundColor DarkGray
} else {
    Write-Host "  Test-AdAvailable says NOT connected:" -ForegroundColor Red
    Write-Host "     $($final.Detail)" -ForegroundColor Red
}
Write-Host ''
