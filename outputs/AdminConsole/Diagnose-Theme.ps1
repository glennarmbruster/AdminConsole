#Requires -Version 5.1
<#
    Diagnose-Theme.ps1
    Says which theme the console WILL use, and why.

    Author   : Glenn Armbruster
    Revised  : 2026-09-16

    2026-09-16  Reports the PER-USER theme, which since the profile layer went
                in is the first rule and was missing here entirely -- so the
                one tool whose job is to answer "why is my theme not applying"
                confidently named the wrong winner.
#>
# TOOLKIT-HELP-RUNS: Reads Themes.json and Tenant.json and prints how the theme resolves. Changes nothing.
# TOOLKIT-VARIANT: Any
# TOOLKIT-HELP-BEGIN
# Prints every step of theme resolution so "I picked a theme and nothing
# changed" has an answer instead of a guess.
#
# There are five ways a theme can be chosen and four ways the choice can be
# silently discarded -- a disabled theme, a tenant binding overriding it, a
# name that is not in Themes.json at all, or a console that cached the profile
# before the save. Each of those looks identical from the console: it just
# opens in the wrong colours.
#
# Run it from a prompt. It changes nothing.
# TOOLKIT-HELP-END

. (Join-Path $PSScriptRoot 'Common.ps1')

# ---- WHICH BUILD IS IN THIS FOLDER ------------------------------------------
#
# ASKED FIRST, BECAUSE IT IS THE ANSWER MOST OFTEN. The theme fix of 2026-09-16
# spans TWO files -- the resolver in Common.ps1 and the save logic in
# ThemeManager.ps1 -- and copying one without the other reproduces the original
# symptom exactly: the label names a theme you are not using, and the one theme
# whose name matches ActiveTheme cannot be selected. Nothing on screen
# distinguishes 'the fix does not work' from 'the fix is not here', so this
# line exists to distinguish them.
Write-Host ''
Write-Host 'Build check:' -ForegroundColor Cyan
$hasResolver = [bool](Get-Command -Name 'Resolve-ConsoleThemeName' -CommandType Function -ErrorAction SilentlyContinue)
Write-Host ("  Common.ps1        : {0}" -f $(if ($hasResolver) { 'has the shared resolver  (2026-09-16 or later)' } else { 'OLD -- no Resolve-ConsoleThemeName' })) `
    -ForegroundColor $(if ($hasResolver) { 'Green' } else { 'Red' })

$tmPath = Join-Path $PSScriptRoot 'ThemeManager.ps1'
$tmNew  = $false
try { if (Test-Path -LiteralPath $tmPath) { $tmNew = [bool](Select-String -LiteralPath $tmPath -Pattern 'EffectiveThemeOnDisk' -SimpleMatch -Quiet) } } catch { }
Write-Host ("  ThemeManager.ps1  : {0}" -f $(if ($tmNew) { 'compares against YOUR theme  (2026-09-16 or later)' } else { 'OLD -- compares against the shared ActiveTheme' })) `
    -ForegroundColor $(if ($tmNew) { 'Green' } else { 'Red' })
if (-not $hasResolver -or -not $tmNew) {
    Write-Host ''
    Write-Host '  >> THE FILES ARE NOT BOTH UPDATED. Copy Common.ps1 AND ThemeManager.ps1' -ForegroundColor Red
    Write-Host '     from the ships folder before reading anything below as a real fault.' -ForegroundColor Red
}

$tf = Get-ThemeFilePath
Write-Host ''
Write-Host "Themes file : $tf" -ForegroundColor Cyan
if (-not (Test-Path -LiteralPath $tf)) { Write-Host '  NOT FOUND -- the console runs unthemed.' -ForegroundColor Red; return }

$cfg = Get-Content -LiteralPath $tf -Raw | ConvertFrom-Json
Write-Host "ActiveTheme : $($cfg.ActiveTheme)"
Write-Host "FollowTenant: $($cfg.FollowTenant)"

# WHOSE PROFILE, AND WHETHER IT IS BEING READ FRESH. Printed before the
# resolution because a stale cache explains more reported cases than any of
# the rules do -- this runs in its own process, so what it reads here is the
# file as it stands now, which is the point of comparison.
$pp = Get-UserProfilePath
Write-Host ''
Write-Host "Your profile: $pp" -ForegroundColor Cyan
if (Test-Path -LiteralPath $pp) {
    Write-Host "  written     : $((Get-Item -LiteralPath $pp).LastWriteTime)"
} else {
    Write-Host '  none yet -- nothing is customised for this Windows account.'
}
$tenantKey = ''
try { $tenantKey = "$(Get-ActiveTenantKey)" } catch { }
Write-Host "  tenant key  : $(if ($tenantKey) { $tenantKey } else { '(none -- the Default block is what applies)' })"

Write-Host ''
Write-Host 'Resolution, in order:' -ForegroundColor Cyan
$name = "$($cfg.ActiveTheme)"
Write-Host "  1. ActiveTheme from Themes.json      -> '$name'"

$multi = $false
try { $multi = [bool](Test-MultiTenantMode) } catch { }
Write-Host "  2. Multi-tenant mode                 -> $multi"

# THE ORDER HERE IS Get-ConsoleTheme's ORDER. If one changes, change both --
# a diagnostic that resolves differently from the thing it diagnoses is worse
# than no diagnostic, because it is believed.
$ut = ''
try { $ut = "$(Get-UserLayoutProperty -Name 'Theme')".Trim() } catch { }
if ($ut) {
    $name = $ut
    Write-Host "  3. YOUR profile names a theme        -> '$ut'  (this wins)" -ForegroundColor Yellow
} else {
    Write-Host "  3. Your profile names no theme       -> the tenant decides"
}

if ($multi) {
    $tt = ''
    try { $tt = "$((Get-TenantConfig).Theme)".Trim() } catch { }
    if ($tt -and -not $ut) { $name = $tt; Write-Host "     Tenant.json Theme WINS            -> '$tt'" -ForegroundColor Yellow }
    elseif ($tt)           { Write-Host "     Tenant.json says '$tt' -- outranked by yours" }
    else                   { Write-Host "     Tenant.json names no theme        -> ActiveTheme stands" }
} else {
    Write-Host "     Single tenant: Tenant.json's Theme is IGNORED by design."
}

if (-not $multi -and $cfg.FollowTenant) {
    $d = ''
    try { $d = Get-ConnectedTenantDomain } catch { }
    Write-Host "  3. FollowTenant is ON, domain        -> '$d'"
    if ($d -and $cfg.TenantThemes.PSObject.Properties.Name -contains $d) {
        $name = $cfg.TenantThemes.$d
        Write-Host "     Tenant binding WINS               -> '$name'" -ForegroundColor Yellow
    }
}

# ---- AND WHAT THE TOOLKIT ITSELF SAYS ---------------------------------------
#
# The steps above are narration -- they explain. This is the ANSWER, from the
# same function the console and Theme Manager call. A diagnostic that resolves
# differently from the thing it diagnoses is worse than no diagnostic, because
# it is believed; so the two are printed together and disagreement is reported
# rather than hidden.
if ($hasResolver) {
    $r = $null
    try { $r = Resolve-ConsoleThemeName -Config $cfg } catch { }
    if ($r) {
        Write-Host ''
        Write-Host "Resolve-ConsoleThemeName says: '$($r.Name)'  (from $($r.Detail))" -ForegroundColor Cyan
        if ("$($r.Name)" -ne "$name") {
            Write-Host "  DISAGREES with the walk-through above ('$name'). Trust this line --" -ForegroundColor Red
            Write-Host '  it is the code the console actually runs -- and report the mismatch.' -ForegroundColor Red
        }
        $name = "$($r.Name)"
    }
}

# ---- CAN THEME MANAGER ACTUALLY SWITCH YOU? ---------------------------------
#
# The question behind the original report. On the OLD build, selecting the
# theme whose name equals ActiveTheme read as 'nothing changed', so nothing was
# written and that one theme could never be chosen.
if ("$($cfg.ActiveTheme)" -and "$name" -eq "$($cfg.ActiveTheme)") {
    Write-Host ''
    Write-Host "You are currently on '$name', which is ALSO what ActiveTheme says." -ForegroundColor Cyan
    Write-Host '  Nothing to prove here -- pick a different theme to test switching.'
} elseif ("$($cfg.ActiveTheme)") {
    Write-Host ''
    Write-Host "Selecting '$($cfg.ActiveTheme)' in Theme Manager is the case that used to fail:" -ForegroundColor Cyan
    Write-Host ("  it matches ActiveTheme but not your theme ('$name'), so the old build" )
    Write-Host '  concluded nothing had changed and wrote nothing.'
    Write-Host ("  This build: {0}" -f $(if ($tmNew) { 'asks you, then writes it to your profile.' } else { 'STILL BROKEN -- ThemeManager.ps1 is the old one.' })) `
        -ForegroundColor $(if ($tmNew) { 'Green' } else { 'Red' })
}

Write-Host ''
Write-Host "Resolved name: '$name'" -ForegroundColor Cyan
if (-not $name) { Write-Host '  Empty -- console runs unthemed.' -ForegroundColor Red; return }
if ($cfg.Themes.PSObject.Properties.Name -notcontains $name) {
    Write-Host "  '$name' IS NOT IN Themes.json -- console runs unthemed." -ForegroundColor Red; return
}
if (-not $cfg.Themes.$name.Enabled) {
    Write-Host "  '$name' EXISTS BUT IS DISABLED -- console runs unthemed." -ForegroundColor Red
    Write-Host "  That is the silent one: the save succeeds and nothing says why." -ForegroundColor Red
    return
}
Write-Host "  '$name' is present and enabled. The console will use it." -ForegroundColor Green
Write-Host ''
Write-Host 'If the console still looks wrong:' -ForegroundColor Yellow
Write-Host '  - it was opened before the last save, or' -ForegroundColor Yellow
Write-Host '  - it is running in a PowerShell window that read your profile BEFORE' -ForegroundColor Yellow
Write-Host '    Theme Manager wrote to it. Close the console and open a NEW window.' -ForegroundColor Yellow
Write-Host '    (Consoles from 2026-09-16 on re-read the file and do not need this.)' -ForegroundColor Yellow
