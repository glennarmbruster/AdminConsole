#Requires -Version 5.1
# TOOLKIT-HELP-RUNS: nothing against any tenant. It reads this folder and, with -Apply, edits Config\Tools.json, Config\Sections.json and Config\Themes.json after backing them up.
# TOOLKIT-HELP-BEGIN
# Puts NEW SCRIPTS ON BUTTONS without shipping a Tools.json over yours.
#
# The rule this exists to obey: no update may deliver Config\Tools.json,
# Config\Sections.json or Config\Settings.json. Those files hold an arrangement
# somebody built by hand -- buttons dragged, sections grown, accents chosen,
# paths set -- and replacing all of that to add one button is not an update, it
# is a reset.
#
# So new tools arrive as .ps1 files only, and this merges them in: it reads
# what is already there, adds only what is missing, and writes back after a
# dated backup under Config\Backup\. Run it twice and the second run does
# nothing.
#
#   .\Add-ConsoleAdditions.ps1           # say what would change, change nothing
#   .\Add-ConsoleAdditions.ps1 -Apply    # back up, then merge
#
# It finds work in two ways:
#
#   1. Config\Additions-*.json -- a curated list that names each tool, the
#      section it belongs in and what it works on. This is what a delivery
#      comes with when somebody has thought about where things go.
#
#   2. A SCAN for .ps1 files in this folder that no button points at. This is
#      the safety net for "I dropped a script in and forgot". A scanned tool
#      has no section to go in, so it lands in the first one and the summary
#      says so -- drag it and press Save Arrangement.
#
# TWO KINDS OF FILE ARE NEVER OFFERED, and the list is in the script:
#
#   Infrastructure -- the console itself, Common.ps1, the publisher, the
#   tester, this file, the dot-sourced query libraries.
#
#   Anything the console already reaches -- Settings, Tenant Manager, Theme
#   Manager, Tool Manager, Tool Status, QuickLaunch, Config Backup, Unblock
#   Files, Bug Fix and Feature Request from the Dashboard, and Unified Lookup
#   from the Look Up box at the top of the console. One route to a window is
#   enough; a second copy on the grid only raises the question of whether
#   they differ.
#
# Decline anything else and it is written to Config\Additions-Ignore.json so
# you are not asked again; delete that file to be asked afresh.
#
# Publish-Toolkit runs this at the end of every publish, which is the moment
# new scripts are in the folder and the arrangement is in front of you.
# TOOLKIT-HELP-END
<#
    Add-ConsoleAdditions.ps1
    Merge new tools and themes into the live Config, without replacing it.

    Author   : Glenn Armbruster
    Revised  : 2026-09-12
#>
param(
    # Without this nothing is written. The plan is printed and the script ends.
    [switch]$Apply,

    # Skip the scan and use only Config\Additions-*.json. For a delivery that
    # named its own tools and where a stray .ps1 in the folder is not wanted.
    [switch]$NoScan,

    # One line of output instead of a report. Publish-Toolkit uses this to ask
    # its question without printing the plan twice.
    [switch]$Brief
)

$ErrorActionPreference = 'Stop'

# COMMON.PS1, AND IT WAS MISSING UNTIL 2026-09-17.
#
# This script keeps its own tiny Read-JsonFile / Write-JsonFile rather than
# Common's, which is why nobody noticed: it looked standalone and was, right up
# until it learned about PARKED scripts. That change added three calls --
# Get-ParkedKey, Get-EffectiveParked, Get-ParkedScriptIndex -- and did not add
# the line that makes them exist. The script then failed on the first parked
# script it met, with "Get-ParkedKey is not recognized", which reads like a
# corrupt install rather than a missing import.
#
# NOTHING COLLIDES. Common.ps1 declares no top-level variables and needs no
# assembly loaded to be sourced, and none of this file's own functions share a
# name with one of its. Test-Toolkit now has a rule for this whole class --
# calling Common without importing it -- so the next one is caught at scan time
# rather than at somebody's keyboard.
. (Join-Path $PSScriptRoot 'Common.ps1')

$cfg = Join-Path $PSScriptRoot 'Config'

# NEVER OFFERED. Two kinds of file, one list, and both are permanent.
#
# 1. INFRASTRUCTURE -- not a tool at all.
#
#      AdminConsole / Common      the console and its shared code
#      Publish-Toolkit / Deploy   build-time, run from a prompt
#      Test-Toolkit               the self-test, run from a prompt
#      Add-ConsoleAdditions       this file
#      *Query.ps1                 dot-sourced libraries (matched separately)
#
# 2. IT LIVES ON THE DASHBOARD, AND THAT IS ENOUGH.
#
#    Settings, Tenant Manager, Theme Manager, Tool Manager, Tool Status,
#    QuickLaunch, Config Backup, Unblock Files, Bug Fix and Feature Request
#    all open from the Dashboard. Putting a second copy on the grid gives the
#    operator two routes to one window and invites the question of whether
#    they differ -- and it spends a button on something nobody reaches for
#    from the grid anyway. These are configuration and housekeeping, not
#    tools you run against a mailbox.
#
#    UNIFIED LOOKUP IS HERE FOR THE SAME REASON BY A DIFFERENT ROUTE. Since
#    2026-09-12 the console's own Look Up button opens it with whatever was
#    typed, search already run (ARCHITECTURE 45d). It is the front door, not
#    a tool on the grid -- and a button that opens it EMPTY, beside a box
#    that opens it filled in, is the worse of the two and indistinguishable
#    at a glance. Keep the file: the console launches it BY NAME.
#
#    A scan cannot work any of this out for itself: they are ordinary scripts
#    with ordinary help blocks, and they look exactly like tools. So the decision
#    is written down here, once, rather than re-declined on every publish.
#
# ADDING TO THIS LIST IS THE RIGHT MOVE when something else earns a permanent
# home on the Dashboard. Config\Additions-Ignore.json is for one operator
# declining one script; this list is for what the toolkit itself has decided.
$script:NeverOffer = @(
    # infrastructure
    'AdminConsole.ps1', 'Common.ps1', 'Publish-Toolkit.ps1', 'Deploy.ps1',
    'Test-Toolkit.ps1', 'Add-ConsoleAdditions.ps1', 'REVERT-MultiTenant.ps1',
    'Diagnose-Theme.ps1',
    # reached from the Dashboard
    'Settings.ps1', 'TenantManager.ps1', 'ThemeManager.ps1', 'ToolManager.ps1',
    'ToolStatus.ps1', 'QuickLaunchTool.ps1', 'ConfigBackup.ps1',
    'UnBlock.ps1', 'BugFix.ps1', 'FeatureRequest.ps1', 'ParkInPlace.ps1',
    'Customizations.ps1', 'ViewManager.ps1',
    # reached from the Look Up button on the console header
    'Unified-Lookup.ps1'
)

function Get-DashboardScript {
    <#
        Every script the Dashboard already has a button for.

        THE RULE: IF IT IS ON THE DASHBOARD, IT LIVES THERE AND NOWHERE ELSE.
        A second button on the tool grid gives the operator two routes to one
        window and invites the question of whether they differ. It also spends
        a grid button on configuration and housekeeping -- things nobody
        reaches for from the grid anyway.

        WHY THIS IS READ AND NOT JUST LISTED. The list above is hand-written,
        and a hand-written list of Dashboard tools drifts the moment somebody
        adds a Dashboard tool and forgets it -- which is exactly what happened
        with ViewManager.ps1 on 2026-09-16: it shipped with a Dashboard button
        and was offered by the next publish anyway. Reading Dashboard.ps1 makes
        the rule enforce itself, so the NEXT one is covered on the day it is
        added rather than the day somebody notices.

        IT CAN ONLY EVER ADD EXCLUSIONS. A parse that finds nothing leaves the
        hand-written list exactly as it was, so a Dashboard that has been
        reshaped beyond recognition costs nothing here -- and both lists are
        belt and braces for each other.

        MANUAL ADDING IS UNAFFECTED. This is the SCAN declining to offer a
        script. Tool Manager will still add any of these by hand if somebody
        decides they want a grid button after all; that is a decision, and
        this list is only about what gets suggested.
    #>
    $out = @()
    try {
        $dash = Join-Path $PSScriptRoot 'Dashboard.ps1'
        if (-not (Test-Path -LiteralPath $dash)) { return $out }
        $text = Get-Content -LiteralPath $dash -Raw -Encoding UTF8
        foreach ($m in [regex]::Matches($text, "(?i)Script\s*=\s*'([^']+\.ps1)'")) {
            $n = [System.IO.Path]::GetFileName($m.Groups[1].Value)
            if ($n -and $out -notcontains $n) { $out += $n }
        }
    } catch { return @() }
    return $out
}

function Read-JsonFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json)
}
function Write-JsonFile {
    param([string]$Path, $Object)
    $Object | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Get-IgnoreList {
    $p = Join-Path $cfg 'Additions-Ignore.json'
    $h = @{}
    $j = Read-JsonFile -Path $p
    if ($j) { foreach ($n in @($j.Ignore)) { if ("$n") { $h["$n".ToLower()] = $true } } }
    return $h
}
function Add-ToIgnoreList {
    param([string[]]$Names)
    $p = Join-Path $cfg 'Additions-Ignore.json'
    $have = Get-IgnoreList
    foreach ($n in @($Names)) { if ("$n") { $have["$n".ToLower()] = $true } }
    $obj = [ordered]@{
        _comment = 'Scripts NOT to offer as new buttons. Written when you decline them in Add-ConsoleAdditions. Delete this file to be asked about them again.'
        Ignore   = @($have.Keys | Sort-Object)
    }
    Write-JsonFile -Path $p -Object ([PSCustomObject]$obj)
}

# ---- what is already on a button -------------------------------------------
$toolsPath = Join-Path $cfg 'Tools.json'
$secPath   = Join-Path $cfg 'Sections.json'
$themePath = Join-Path $cfg 'Themes.json'

$tools = @(Read-JsonFile -Path $toolsPath)
if ($tools.Count -eq 0) { throw "Config\Tools.json not found or empty. Run the console once first." }
$sec = Read-JsonFile -Path $secPath
if (-not $sec) { throw "Config\Sections.json not found. Run the console once first." }

$perRow = [int]$sec.ButtonsPerRow
if ($perRow -lt 1) { $perRow = 5 }
$secNames = @($sec.Sections | ForEach-Object { "$($_.Name)" })
$haveScript = @{}
foreach ($t in $tools) { $haveScript["$($t.Script)".ToLower()] = $true }

$plan     = New-Object System.Collections.Generic.List[string]
$newTools = @()
$scanned  = @()

function Add-PlannedTool {
    param([string]$Name, [string]$Script, [string]$Section, $ObjectTypes, [string]$Tier, [bool]$FromScan)

    $key = "$Script".ToLower()
    if ($haveScript[$key]) { return }
    $file = Join-Path $PSScriptRoot ("$Script" -replace '^\.\\', '')
    if (-not (Test-Path -LiteralPath $file)) {
        $plan.Add("  skip   $Name -- $Script is not in this folder")
        return
    }
    $s = "$Section"
    if ($secNames -notcontains $s) {
        $s = $secNames[0]
        if (-not $FromScan) { $plan.Add("  note   $Name -- section '$Section' does not exist here, using '$s'") }
    }
    # Order continues the section it lands in, counting what is already planned
    # as well as what is already saved.
    $inSec = @($tools | Where-Object { "$($_.Section)" -eq $s }) + @($script:newTools | Where-Object { "$($_.Section)" -eq $s })
    $order = 0
    foreach ($x in $inSec) { if ([int]$x.Order -ge $order) { $order = [int]$x.Order + 1 } }

    $script:newTools += [PSCustomObject]([ordered]@{
        Order       = $order
        Name        = $Name
        Script      = $Script
        ObjectTypes = @($ObjectTypes)
        Section     = $s
        Tier        = "$Tier"
        NewProcess  = $false
    })
    $haveScript[$key] = $true
    $plan.Add("  add    $Name -> $s$(if ($FromScan) { '   (found by scan -- move it where you want it)' })")
}

# ---- 1. curated Additions-*.json -------------------------------------------
$addFiles = @(Get-ChildItem -LiteralPath $cfg -Filter 'Additions-*.json' -File -ErrorAction SilentlyContinue |
              Where-Object { $_.Name -notlike '*Ignore*' } | Sort-Object Name)
$themesToAdd = @{}
foreach ($af in $addFiles) {
    $add = $null
    try { $add = Read-JsonFile -Path $af.FullName } catch {
        $plan.Add("  skip   $($af.Name) -- not valid JSON: $($_.Exception.Message)")
        continue
    }
    foreach ($t in @($add.Tools)) {
        Add-PlannedTool -Name "$($t.Name)" -Script "$($t.Script)" -Section "$($t.Section)" `
                        -ObjectTypes $t.ObjectTypes -Tier "$($t.Tier)" -FromScan:$false
    }
    if ($add.PSObject.Properties['SectionAccents']) {
        foreach ($s in $sec.Sections) {
            $has = ($s.PSObject.Properties['Accent'] -and "$($s.Accent)".Trim())
            if ($has) { continue }
            $p = $add.SectionAccents.PSObject.Properties["$($s.Name)"]
            if (-not $p -or -not "$($p.Value)") { continue }
            $plan.Add("  accent $($s.Name) -> $($p.Value)")
            $s | Add-Member -NotePropertyName 'Accent' -NotePropertyValue "$($p.Value)" -Force
            $script:secChanged = $true
        }
    }
    if ($add.PSObject.Properties['Themes']) {
        foreach ($p in $add.Themes.PSObject.Properties) { $themesToAdd[$p.Name] = $p.Value }
    }
}

# ---- 2. the scan -----------------------------------------------------------
if (-not $NoScan) {
    $ignore = Get-IgnoreList
    # Read ONCE, outside the loop -- Get-ParkedScriptIndex re-reads the file
    # every call, and doing that per script turns a scan into a hundred reads.
    # TENANT PARKING AND THIS OPERATOR'S OWN, as one list. Offering a tool
    # somebody hid from themselves is the same annoyance as offering one the
    # tenant parked -- they said no once already.
    $parkedIdx = @{}
    try { $parkedIdx = (Get-EffectiveParked).Scripts } catch {
        try { $parkedIdx = Get-ParkedScriptIndex } catch { $parkedIdx = @{} }
    }
    # ON THE DASHBOARD ALREADY? Then it is not offered here. Folded in at the
    # point of use rather than into the literal above, so the hand-written list
    # stays readable as 'what the toolkit decided' and this stays visible as
    # 'and whatever the Dashboard is currently showing'.
    foreach ($d in @(Get-DashboardScript)) {
        if ($script:NeverOffer -notcontains $d) { $script:NeverOffer += $d }
    }

    foreach ($f in @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.ps1' -File | Sort-Object Name)) {
        $leaf = $f.Name
        if ($script:NeverOffer -contains $leaf) { continue }
        if ($leaf -like '*Query.ps1') { continue }
        if ($ignore[$leaf.ToLower()]) { continue }
        if ($haveScript[".\$leaf".ToLower()]) { continue }
        # PARKED FOR THIS TENANT. Offering it again on every publish is exactly
        # what parking exists to stop.
        #
        # NOT THE SAME AS $script:NeverOffer above: that is what the TOOLKIT has
        # decided, for every install, permanently. This is what THIS TENANT
        # decided, and it changes when the tenant does.
        if ($parkedIdx.ContainsKey((Get-ParkedKey -Script $leaf))) { continue }

        # The button caption comes from the file name, tidied: hyphens and
        # underscores to spaces. It is a starting point, not a decision --
        # rename it in Tool Manager, which is one field.
        $nice = ($leaf -replace '\.ps1$', '') -replace '[-_]', ' '
        $scanned += $leaf
        Add-PlannedTool -Name $nice -Script ".\$leaf" -Section $secNames[0] `
                        -ObjectTypes @() -Tier '' -FromScan:$true
    }
}

# ---- themes ----------------------------------------------------------------
$themes = Read-JsonFile -Path $themePath
$themesChanged = $false
if ($themesToAdd.Count -gt 0) {
    if (-not $themes) {
        $plan.Add('  skip   themes -- Themes.json is not here yet (the console builds it from the template)')
    } else {
        foreach ($k in @($themesToAdd.Keys | Sort-Object)) {
            if ($themes.Themes.PSObject.Properties[$k]) { continue }
            $themes.Themes | Add-Member -NotePropertyName $k -NotePropertyValue $themesToAdd[$k] -Force
            $themesChanged = $true
            $plan.Add("  add    theme  $k")
        }
    }
}

# ---- rows ------------------------------------------------------------------
foreach ($s in $sec.Sections) {
    $n = @(@($tools) + @($newTools) | Where-Object { "$($_.Section)" -eq "$($s.Name)" }).Count
    $need = [int][Math]::Ceiling($n / [double]$perRow)
    if ($need -lt 1) { $need = 1 }
    if ($need -gt [int]$s.Rows) {
        $plan.Add("  rows   $($s.Name): $($s.Rows) -> $need")
        $s.Rows = $need
        $script:secChanged = $true
    }
}

$anything = ($newTools.Count -gt 0) -or $script:secChanged -or $themesChanged

# ---- report ----------------------------------------------------------------
if ($Brief) {
    if (-not $anything) { Write-Host 'No additions available.' -ForegroundColor Green }
    else { Write-Host "$($newTools.Count) tool(s) can be added to the console." -ForegroundColor Yellow }
    return
}

Write-Host ''
Write-Host "Config folder : $cfg"
if ($addFiles.Count -gt 0) { Write-Host "Additions file: $(($addFiles | ForEach-Object { $_.Name }) -join ', ')" }
Write-Host ''
foreach ($line in $plan) { Write-Host $line }
if (-not $anything) {
    Write-Host '  Nothing to do -- every script in this folder is already on a button.' -ForegroundColor Green
    return
}
Write-Host ''
if (-not $Apply) {
    Write-Host 'PREVIEW ONLY. Nothing was written. Run again with -Apply.' -ForegroundColor Yellow
    if ($scanned.Count -gt 0) {
        Write-Host ''
        Write-Host 'Found by scan, so they have no section of their own:' -ForegroundColor Yellow
        foreach ($s in $scanned) { Write-Host "    $s" }
        Write-Host '  If any of these should never be a button, decline once and they go into' -ForegroundColor Yellow
        Write-Host '  Config\Additions-Ignore.json. Use -NoScan to skip the scan entirely.' -ForegroundColor Yellow
    }
    return
}

# ---- apply -----------------------------------------------------------------
$stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
$bak = Join-Path (Join-Path $cfg 'Backup') $stamp
New-Item -Path $bak -ItemType Directory -Force | Out-Null
foreach ($f in @($toolsPath, $secPath, $themePath)) {
    if (Test-Path -LiteralPath $f) { Copy-Item -LiteralPath $f -Destination $bak -Force }
}
Write-Host "Backed up to $bak"

if ($newTools.Count -gt 0) {
    Write-JsonFile -Path $toolsPath -Object @(@($tools) + @($newTools))
    Write-Host "Tools.json: $($newTools.Count) added ($(@($tools).Count + $newTools.Count) total)."
}
if ($script:secChanged) { Write-JsonFile -Path $secPath -Object $sec; Write-Host 'Sections.json: updated.' }
if ($themesChanged)     { Write-JsonFile -Path $themePath -Object $themes; Write-Host 'Themes.json: updated.' }

Write-Host ''
Write-Host 'Done. Start the console -- the new buttons are at the end of their sections.' -ForegroundColor Green
Write-Host 'Drag them where you want them and press Save Arrangement.' -ForegroundColor Green
