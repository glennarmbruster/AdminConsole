#Requires -Version 5.1
<#
    REVERT-MultiTenant.ps1
    Turns multi-tenant mode off. Nothing is deleted.

    Author   : Glenn Armbruster
    Revised  : 2026-09-13
#>
# TOOLKIT-HELP-RUNS: Renames Config\Tenants so the toolkit falls back to Config\Tenant.json. Touches nothing else.
# TOOLKIT-VARIANT: Any
# TOOLKIT-HELP-BEGIN
# Switches multi-tenant mode OFF and puts the console back to reading
# Config\Tenant.json, the way it worked before.
#
# HOW IT WORKS, and why there is so little to it: multi-tenant mode is ON
# because Config\Tenants exists and has a .json in it. Everything else --
# the -Tenant switch, the per-tenant report and log folders, the tenant in the
# title bar, the confirm on write tools -- asks that one question and does
# nothing when the answer is no.
#
# So reverting is renaming one folder. This script renames it to
# Config\Tenants.disabled-<date> rather than deleting it, because the tenant
# files in it are configuration somebody typed.
#
# TO GO BACK: rename it to Tenants again. Nothing else was changed.
#
# WHAT THIS DOES NOT DO: it does not move reports or logs back out of their
# per-tenant subfolders. Those are your files and the tool will not shuffle
# them; the folders simply stop being used.
# TOOLKIT-HELP-END

. (Join-Path $PSScriptRoot 'Common.ps1')

$dir = Join-Path (Join-Path $PSScriptRoot 'Config') 'Tenants'
if (-not (Test-Path -LiteralPath $dir)) {
    Write-Host 'Multi-tenant mode is already off -- there is no Config\Tenants folder.' -ForegroundColor Yellow
    return
}

$files = @(Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue)
Write-Host ''
Write-Host "Config\Tenants holds $($files.Count) file(s):" -ForegroundColor Cyan
foreach ($f in $files) { Write-Host "    $($f.Name)" }
Write-Host ''
Write-Host 'Renaming it turns multi-tenant mode off. The files are kept.' -ForegroundColor Yellow

# THE SINGLE-TENANT FALLBACK HAS TO EXIST, or reverting swaps one problem for a
# worse one: a console with no tenant file at all, where every capability reads
# as false and nothing says why.
$fallback = Join-Path (Join-Path $PSScriptRoot 'Config') 'Tenant.json'
if (-not (Test-Path -LiteralPath $fallback)) {
    Write-Host ''
    Write-Host 'STOP: there is no Config\Tenant.json to fall back to.' -ForegroundColor Red
    Write-Host 'Copy the tenant you want to be the single tenant there first:' -ForegroundColor Red
    if ($files.Count -gt 0) {
        Write-Host "    Copy-Item '$($files[0].FullName)' '$fallback'" -ForegroundColor Red
    }
    return
}

$ans = Read-Host 'Turn multi-tenant mode off? (y/N)'
if ("$ans".Trim() -notmatch '^(?i)y') { Write-Host '  Left alone.' -ForegroundColor Yellow; return }

$dest = "$dir.disabled-$(Get-Date -Format 'yyyy-MM-dd_HHmm')"
Rename-Item -LiteralPath $dir -NewName (Split-Path -Leaf $dest)
Write-Host ''
Write-Host "Done. Renamed to $(Split-Path -Leaf $dest)." -ForegroundColor Green
Write-Host 'The console now reads Config\Tenant.json. -Tenant is ignored.' -ForegroundColor Green
Write-Host 'To turn it back on, rename the folder back to Tenants.' -ForegroundColor Green
