#Requires -Version 5.1
# TOOLKIT-HELP-RUNS: nothing. It works on this folder's own files.
# TOOLKIT-HELP-BEGIN
# The last thing you run before handing the toolkit to another engineer.
#
#   .\Publish-Toolkit.ps1
#
# In order:
#
#   1.  UNBLOCK every file. Anything downloaded or copied from another machine
#       carries a mark-of-the-web tag that makes PowerShell refuse to run it.
#       The files actually blocked are listed -- that list should match what you
#       just downloaded, and anything surprising on it is worth a look.
#
#   2.  SHOW what changed since the last stamp.
#
#   2a. CONFIRM WHICH TENANT this build is for, and refuse to go further unless
#       Config\Tenant.json already says the same thing.
#
#       This is the one thing nothing else can catch. Config\ is excluded from
#       the version manifest, and Test-Toolkit reads the tenant names OUT of
#       Tenant.json -- so a toolkit built against one tenant and published for
#       another would ship the first one's domains and servers past every other
#       guard, and work perfectly, at the wrong company.
#
#   2b. LIST tools this tenant can never run, so they can be removed before
#       shipping rather than greyed out on someone else's screen.
#
#   3.  ASK minor or major, and work the number out from the last one -- 1.03
#       becomes 1.04, or 2.0. Nothing to type, so nothing to mistype.
#
#   4.  STAMP Version.json with the version, the tenant, and a SHA-256 per
#       script.
#
#   5.  OFFER ANY NEW TOOLS A BUTTON. Adding a script to the folder and adding
#       a button for it are two jobs, and the second is the one that gets
#       forgotten -- the file is there, the tester passes, and the tool is
#       simply unreachable. This lists scripts nothing points at, and asks.
#       Answer no and nothing is written. See Add-ConsoleAdditions.ps1, which
#       does the work and can be run on its own at any time.
#
# Unblocking happens BEFORE hashing on purpose. Whatever clearing the tag does
# or does not do to a file, the manifest then describes the bytes exactly as
# they will be handed over.
#
# This is the ONLY thing that writes Version.json. There was briefly a separate
# New-VersionManifest.ps1 that did the stamping alone -- a second way in that
# skipped the confirmation and the reused-version warning above. Two doors, one
# of which bypasses the guardrails, is worse than one door.
#
# RUN THIS ON THE MASTER COPY ONLY. On someone else's copy it re-stamps their
# edits as official, which is exactly backwards.
# TOOLKIT-HELP-END
<#
    Publish-Toolkit.ps1
    Unblock, review, and stamp the toolkit for release.

    Reasoning: ScriptNotes/Versioning.notes.md

    Author   : Glenn Armbruster
    Created  : 2026-08-07
#>

param(
    # Optional. Left out, the script works out the next number itself and asks
    # whether this is a minor or a major release. Supply it to override -- to
    # correct a bad stamp, or to jump deliberately.
    [ValidatePattern('^\d+\.\d+(\.\d+)?$')]
    [string]$Version,

    # Which tenant this build is FOR. Mandatory in spirit: if it is left out the
    # script asks, and it will not stamp until Tenant.json matches the answer.
    [string]$Tenant = '',

    [string]$Notes = '',

    # Skip the confirmation prompt. Deliberately not the default: the prompt
    # exists so the change list gets read, not dismissed.
    [switch]$Force
)

. (Join-Path $PSScriptRoot 'Common.ps1')

$root = $PSScriptRoot

function Write-Step {
    param([string]$Text)
    Write-Host ''
    Write-Host $Text -ForegroundColor Cyan
    Write-Host ('-' * $Text.Length) -ForegroundColor DarkCyan
}

function Get-NextVersion {
    <#
        Works out the next number from the previous one.

            minor : 1.03 -> 1.04
            major : 1.03 -> 2.0

        Deliberately NOT done with [version] arithmetic. [version]'1.03' is
        1.3 -- leading zeros are discarded -- so incrementing that way would
        give 1.4 and quietly destroy the 1.0x style on the way past.

        Instead the previous string is split and the minor field incremented as
        an integer, then re-padded to the width it already had. Whatever
        padding is in use is preserved without this function needing to know
        what the convention is.
    #>
    param(
        [Parameter(Mandatory)][string]$Previous,
        [Parameter(Mandatory)][ValidateSet('Minor', 'Major')][string]$Kind
    )

    $parts = "$Previous".Split('.')
    if ($parts.Count -lt 2) { return $null }

    $majorNum = 0
    $minorNum = 0
    if (-not [int]::TryParse($parts[0], [ref]$majorNum)) { return $null }
    if (-not [int]::TryParse($parts[1], [ref]$minorNum)) { return $null }

    if ($Kind -eq 'Major') {
        # As specified: 1.03 becomes 2.0, not 2.00. A major release is a
        # deliberate break in the sequence, so it reads as a fresh start.
        return "$($majorNum + 1).0"
    }

    $width = $parts[1].Length
    $next  = $minorNum + 1
    $text  = "$next"
    # PadLeft only, never truncate: at 1.99 -> 1.100 the number outgrows the
    # padding, and the number is what matters. Ordering still holds.
    if ($text.Length -lt $width) { $text = $text.PadLeft($width, '0') }
    return "$majorNum.$text"
}

function ConvertTo-VersionOrNull {
    <#
        A cast that answers "not a version" instead of throwing. Used so an
        unparseable stamp in an old Version.json degrades to "cannot compare"
        rather than taking the whole publish down.
    #>
    param([string]$Text)
    if (-not $Text) { return $null }
    try { return [System.Version]$Text } catch { return $null }
}

# ---------------------------------------------------------------------------
# 1. Unblock
# ---------------------------------------------------------------------------
Write-Step '1. Clearing the mark-of-the-web'

$blocked = New-Object System.Collections.ArrayList
$all     = @(Get-ChildItem -LiteralPath $root -File -Recurse -ErrorAction SilentlyContinue)

foreach ($f in $all) {
    # Zone.Identifier is an NTFS alternate data stream. On a non-NTFS volume
    # the stream parameter fails outright, which SilentlyContinue turns into
    # "not blocked" -- correct, because there is nothing there to block it.
    $z = Get-Item -LiteralPath $f.FullName -Stream 'Zone.Identifier' -ErrorAction SilentlyContinue
    if ($z) { [void]$blocked.Add($f) }
}

if ($blocked.Count -eq 0) {
    Write-Host '  Nothing was blocked.'
} else {
    Write-Host "  $($blocked.Count) blocked file(s) -- these came from off this machine:"
    foreach ($f in $blocked) {
        Write-Host ("    {0}" -f (Get-ToolkitRelativePath -FullPath $f.FullName)) -ForegroundColor Yellow
    }
    $failed = @()
    foreach ($f in $blocked) {
        try { Unblock-File -LiteralPath $f.FullName -ErrorAction Stop }
        catch { $failed += (Get-ToolkitRelativePath -FullPath $f.FullName) }
    }
    if ($failed.Count -gt 0) {
        # Stopping here rather than stamping. A manifest is a statement that
        # this copy is ready to hand over, and a file PowerShell will refuse
        # to run is not ready to hand over.
        Write-Host ''
        Write-Host "  Could not unblock $($failed.Count) file(s). Nothing stamped:" -ForegroundColor Red
        foreach ($f in $failed) { Write-Host "      $f" -ForegroundColor Red }
        return
    }
    Write-Host '  Cleared.' -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# 2. What changed since the last stamp
# ---------------------------------------------------------------------------
Write-Step '2. Changes since the last stamp'

# Test-ToolkitIntegrity compares disk against the manifest. On a COPY those
# three lists mean "drift". Here on the master they mean something more
# useful, so they are relabelled rather than reimplemented.
$before = Test-ToolkitIntegrity

if (-not $before.HasManifest) {
    Write-Host '  No previous Version.json -- this is the first stamp.'
    Write-Host "  $(@(Get-ToolkitCodeFile).Count) script(s) will be covered."
} else {
    Write-Host "  Previous : v$($before.Version)  released $($before.Released)"
    Write-Host ''
    $any = $false
    foreach ($set in @(
        @{ Label = 'CHANGED since that stamp'; Items = $before.Modified; Colour = 'Yellow' },
        @{ Label = 'NEW since that stamp';     Items = $before.Extra;    Colour = 'Green'  },
        @{ Label = 'REMOVED since that stamp'; Items = $before.Missing;  Colour = 'Red'    })) {
        if ($set.Items.Count -gt 0) {
            $any = $true
            Write-Host "  $($set.Label): $($set.Items.Count)"
            foreach ($i in $set.Items) { Write-Host "      $i" -ForegroundColor $set.Colour }
            Write-Host ''
        }
    }
    if (-not $any) {
        Write-Host '  Nothing has changed since the last stamp.' -ForegroundColor Green
    }

}

# ---------------------------------------------------------------------------
# 2a. Which tenant is this build for?
#
# THE ONE THING NOTHING ELSE CHECKS. Config\ is excluded from the version
# manifest, so Version.json does not cover Tenant.json. Test-Toolkit's TEN002
# rule reads the tenant names OUT of Tenant.json, so it cannot flag that file
# either -- a checker cannot catch a file it treats as its own definition of
# truth.
#
# Which means a toolkit developed against one tenant and published for another
# ships that first tenant's domains, domain controllers and Exchange server,
# silently, past every other guard in the toolkit. It would work. It would work
# at the wrong company.
#
# So this is the guard: name the target, and the stamp is refused unless
# Tenant.json already says the same thing. Swapping it is offered rather than
# done quietly, because replacing a config file is a decision.
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# WHAT IS IN Config\ AND WHAT WOULD HAPPEN TO IT
#
# Config\ holds three different kinds of file and they want opposite treatment:
#
#   *.TEMPLATE.json   ships. Blank, safe, the thing a new tenant starts from.
#   Tenant.json       tenant FACTS. Shipping one populated with somebody else's
#                     tenant is worse than shipping none -- it looks configured
#                     and is wrong.
#   the rest          per-engineer or per-deployment: paths, tool layout,
#                     colours, tenant bindings. Copying them over an existing
#                     install silently replaces that engineer's setup.
#
# This step exists because all three were copied wholesale in one go, which
# blanked a configured Settings.json and reverted a Tools.json that had just
# been edited. Nothing errored; the loss was only noticed later, by its
# symptoms. A publish step is the last moment anyone is paying attention.
#
# It REPORTS and, for tenant data, refuses. It does not delete anything --
# deciding what a deployment contains is not a publish script's call.
Write-Step '2p. Config\ contents'

$cfgDir = Join-Path $PSScriptRoot 'Config'
$shipSafe = @(); $perDeployment = @(); $tenantData = @()

foreach ($f in @(Get-ChildItem -LiteralPath $cfgDir -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
    if ($f.Name -like '*.TEMPLATE.json') { $shipSafe += $f.Name; continue }
    if ($f.Name -ieq 'Tenant.json')      { $tenantData += $f.Name; continue }

    # A themes file carrying tenant bindings is tenant data wearing a
    # per-deployment name, so it is judged on CONTENT rather than filename.
    if ($f.Name -ieq 'Themes.json') {
        try {
            $tj = Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json
            if ($tj.TenantThemes -and @($tj.TenantThemes.PSObject.Properties).Count -gt 0) {
                $tenantData += "$($f.Name) (binds $(@($tj.TenantThemes.PSObject.Properties).Count) domain(s) to themes)"
                continue
            }
        } catch { }
    }
    $perDeployment += $f.Name
}

if ($shipSafe.Count -gt 0) {
    Write-Host '  Ships (blank templates):' -ForegroundColor Green
    foreach ($n in $shipSafe) { Write-Host "    $n" -ForegroundColor Green }
}
if ($perDeployment.Count -gt 0) {
    Write-Host '  DO NOT copy over an existing install -- these are that engineer''s:' -ForegroundColor Yellow
    foreach ($n in $perDeployment) { Write-Host "    $n" -ForegroundColor Yellow }
    Write-Host '    Copy the .ps1 files and Projects\, and leave Config\ alone.' -ForegroundColor Yellow
}
if ($tenantData.Count -gt 0) {
    Write-Host ''
    Write-Host '  EXCLUDE when copying to another tenant -- this is tenant data:' -ForegroundColor Red
    foreach ($n in $tenantData) { Write-Host "    $n" -ForegroundColor Red }
    Write-Host '    It hands one tenant''s domains, directories and licence groups to' -ForegroundColor Red
    Write-Host '    the next. Tenant.TEMPLATE.json ships instead, and TenantManager' -ForegroundColor Red
    Write-Host '    builds the real file on first run.' -ForegroundColor Red

    # Reported, NOT prompted. This folder has to contain a real Tenant.json for
    # the console to run at all, so a confirmation here would fire on every
    # publish forever -- and a question asked every time is a question answered
    # without reading. The risk is in the COPY, which happens outside this
    # script, so the useful thing this can do is name the files.
}
if ($shipSafe.Count -eq 0) {
    Write-Host '  No *.TEMPLATE.json in Config\ -- a new tenant would have nothing to start from.' -ForegroundColor Yellow
}

Write-Step '2a. Target tenant'

$cfg = Get-TenantConfig -Refresh
$inFile = "$($cfg.TenantName)".Trim()

if (-not $inFile) {
    Write-Host '  Config\Tenant.json has no TenantName. Fill it in before publishing.' -ForegroundColor Red
    return
}

if (-not $Tenant) {
    Write-Host "  Config\Tenant.json currently says: $inFile"
    $Tenant = (Read-Host "  Publishing for which tenant? [$inFile]").Trim()
    if (-not $Tenant) { $Tenant = $inFile }
}

if ($Tenant -ne $inFile) {
    Write-Host ''
    Write-Host "  MISMATCH. Building for '$Tenant' but Config\Tenant.json says '$inFile'." -ForegroundColor Red
    Write-Host ''
    Write-Host "  Shipping as-is would send ${inFile}'s directories and servers to ${Tenant}:" -ForegroundColor Red
    foreach ($d in @($cfg.Directories)) { Write-Host "      $($d.Domain)" -ForegroundColor Red }
    if ($cfg.OnPremExchange.ServerUri) { Write-Host "      $($cfg.OnPremExchange.ServerUri)" -ForegroundColor Red }
    Write-Host ''

    # A matching file in TenantConfigs, one level up and deliberately outside
    # the shipped folder, is the intended source.
    $src = Join-Path (Split-Path -Parent $root) "TenantConfigs\Tenant.$Tenant.json"
    if (Test-Path -LiteralPath $src) {
        Write-Host "  Found $src"
        $ans = Read-Host "  Copy it over Config\Tenant.json and continue? [y/N]"
        if ($ans -match '^(y|yes)$') {
            Copy-Item -LiteralPath $src -Destination (Get-TenantConfigPath) -Force
            $cfg = Get-TenantConfig -Refresh
            Write-Host "  Tenant.json is now '$($cfg.TenantName)'." -ForegroundColor Green
            if ("$($cfg.TenantName)".Trim() -ne $Tenant) {
                Write-Host "  ...but it still does not say '$Tenant'. Nothing stamped." -ForegroundColor Red
                return
            }
        } else {
            Write-Host '  Nothing stamped.' -ForegroundColor Yellow
            return
        }
    } else {
        Write-Host "  No TenantConfigs\Tenant.$Tenant.json to copy from." -ForegroundColor Red
        Write-Host '  Create one from Tenant.TEMPLATE.json, or re-run naming the tenant this build really is for.' -ForegroundColor Red
        Write-Host '  Nothing stamped.' -ForegroundColor Red
        return
    }
} else {
    Write-Host "  Building for '$Tenant', and Config\Tenant.json agrees." -ForegroundColor Green
}

# The capability report below reads from whatever Tenant.json now says, so it
# has to come after this -- otherwise it describes the wrong tenant.
# ---------------------------------------------------------------------------
# 2b. Tools this tenant can never run
# ---------------------------------------------------------------------------
#
# The console greys these at run time, which is the right behaviour for an
# engineer who is already looking at the window. It is the wrong moment for the
# person SHIPPING it -- by then the dead buttons are already on someone else's
# machine, and "ignore the grey ones" is an instruction nobody remembers.
#
# So the same question gets asked here, where it can still be acted on. This
# only reports. Removing a tool is a decision about what a deployment should
# contain, and a publish script should not be quietly deleting scripts.
Write-Step '2b. Capability check'

$capNames = @{ ExO = 'ExchangeOnline'; SPO = 'SharePointOnline'
               EMS = 'OnPremExchange'; AD  = 'OnPremAD' }
$absent = @()
foreach ($id in $capNames.Keys) {
    $has = $true
    try { $has = Test-TenantCapability -Name $capNames[$id] } catch { }
    if (-not $has) { $absent += $id }
}

if ($absent.Count -eq 0) {
    Write-Host '  This tenant has every capability the toolkit knows about.'
} else {
    Write-Host "  Tenant.json says this tenant does NOT have: $($absent -join ', ')"
    Write-Host ''

    $dead = New-Object System.Collections.ArrayList
    try {
        $toolsPath = Get-ConfigFilePath -Name 'Tools.json'
        $raw   = Get-Content -LiteralPath $toolsPath -Raw -ErrorAction Stop | ConvertFrom-Json
        $tools = if ($raw.PSObject.Properties['Tools']) { @($raw.Tools) } else { @($raw) }
        foreach ($tool in $tools) {
            $rel = "$($tool.Script)" -replace '^\.[\\/]', ''
            $full = Join-Path $root $rel
            if (-not (Test-Path -LiteralPath $full)) { continue }
            $req = Get-ScriptCapabilityRequirement -Path $full
            $blocked = @(@($req.Requires) | Where-Object { $absent -contains $_ })
            if ($blocked.Count -gt 0) {
                [void]$dead.Add([PSCustomObject]@{
                    Name = "$($tool.Name)"; Script = $rel; Needs = ($blocked -join ', ')
                })
            }
        }
    } catch {
        Write-Host "  Could not read Tools.json: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    if ($dead.Count -eq 0) {
        Write-Host '  No tool in Tools.json needs a missing capability.' -ForegroundColor Green
    } else {
        Write-Host "  $($dead.Count) tool(s) can never run in this tenant:" -ForegroundColor Yellow
        foreach ($d in $dead) {
            Write-Host ("    {0,-28} {1,-38} needs {2}" -f $d.Name, $d.Script, $d.Needs) -ForegroundColor Yellow
        }
        Write-Host ''
        Write-Host '  Remove them from Tools.json and from this folder before shipping,' -ForegroundColor Yellow
        Write-Host '  or leave them and accept a row of permanently greyed buttons.' -ForegroundColor Yellow
    }
}

# EX- files ship to one tenant only. Worth listing at the same moment, for the
# same reason: this is the last point at which leaving one in is still a choice.
$exFiles = @(Get-ChildItem -LiteralPath $root -Filter 'EX-*.ps1' -File -ErrorAction SilentlyContinue)
if ($exFiles.Count -gt 0) {
    Write-Host ''
    Write-Host "  $($exFiles.Count) EX- file(s) present -- single-tenant by declaration:"
    foreach ($e in $exFiles) { Write-Host "    $($e.Name)" }
    Write-Host '  These are exempt from the tenant-leakage check and contain tenant data.'
}

Write-Host '  Run Tester first if you have not -- this does not check the code, only fingerprints it.' -ForegroundColor DarkGray

# ---------------------------------------------------------------------------
# 3. Pick the number
#
# Typed by hand this was the one step that could silently go wrong: a typo
# lands on a real-looking number, and 1.02 stamped over 1.03 is a stale version
# that looks entirely legitimate. Offering the next number instead removes the
# whole class rather than guarding against it -- a number never typed cannot be
# mistyped or reversed.
#
# The guards below still run. They now cover the -Version override path, and
# they would catch a bug in Get-NextVersion, which is the point of keeping a
# check that "cannot" fire.
# ---------------------------------------------------------------------------
Write-Step '3. Version'

$explicit = [bool]$Version

if (-not $explicit) {
    if (-not $before.HasManifest) {
        # Nothing to count from.
        $ans = Read-Host '  No previous version. Number for this first stamp [1.00]'
        $Version = if ([string]::IsNullOrWhiteSpace($ans)) { '1.00' } else { $ans.Trim() }
        if ($Version -notmatch '^\d+\.\d+(\.\d+)?$') {
            Write-Host "  '$Version' is not a version number. Nothing stamped." -ForegroundColor Red
            return
        }
    } else {
        $minor = Get-NextVersion -Previous $before.Version -Kind 'Minor'
        $major = Get-NextVersion -Previous $before.Version -Kind 'Major'
        if (-not $minor -or -not $major) {
            Write-Host "  Cannot read '$($before.Version)' as a version. Re-run with -Version." -ForegroundColor Red
            return
        }

        Write-Host "  Current: v$($before.Version)"
        Write-Host ''
        Write-Host "    [1]  minor  ->  v$minor"
        Write-Host "    [2]  major  ->  v$major"
        Write-Host '    [3]  cancel'
        Write-Host ''
        $ans = Read-Host '  Which? [1]'
        switch ("$ans".Trim()) {
            ''      { $Version = $minor }
            '1'     { $Version = $minor }
            '2'     { $Version = $major }
            default {
                Write-Host '  Cancelled. Nothing was stamped (the unblock above still stands).' -ForegroundColor Yellow
                return
            }
        }
        Write-Host "  -> v$Version" -ForegroundColor Green
    }
}

if ($before.HasManifest) {
    # A shipped version number is spent. Re-using it means two different
    # toolkits both answer "1.03" when asked, and the number is the thing said
    # out loud on a call. The hashes would still catch the difference, but
    # they would report it as drift with no way to tell which 1.03 is which.
    if ($before.Version -eq $Version -and $any) {
        Write-Host ''
        Write-Host "  WARNING: v$Version has already been stamped, and files have changed since." -ForegroundColor Red
        Write-Host "           If v$Version has left this machine, bump to a new number instead." -ForegroundColor Red
    }

    # Going BACKWARDS is almost always a typo, and it is the worst kind of
    # collision: a stale number that looks perfectly legitimate.
    #
    # Compared as [version], never as strings -- "1.10" -lt "1.9" is TRUE as
    # strings, which would call the tenth release older than the ninth.
    $prev = ConvertTo-VersionOrNull $before.Version
    $new  = ConvertTo-VersionOrNull $Version
    if ($prev -and $new) {
        if ($new -lt $prev) {
            Write-Host ''
            Write-Host "  STOPPING: v$Version is OLDER than the current stamp, v$($before.Version)." -ForegroundColor Red
            Write-Host '            Almost certainly a typo. Use -Force if it really is deliberate.' -ForegroundColor Red
            if (-not $Force) { return }
            Write-Host '            -Force given; continuing anyway.' -ForegroundColor Yellow
        }
        elseif ($new -eq $prev -and $Version -ne $before.Version) {
            # e.g. 1.01 and 1.1. Leading zeros are cosmetic: [version] parses
            # each part as an integer, so "01" and "1" are the same number.
            Write-Host ''
            Write-Host "  WARNING: v$Version and the current v$($before.Version) look different but are" -ForegroundColor Red
            Write-Host '           the SAME version -- leading zeros are not significant.' -ForegroundColor Red
            Write-Host '           Pick one padding style and keep it.' -ForegroundColor Red
        }
    }
}

# The menu above already was the confirmation, so only ask again when the
# number came from -Version and nobody has agreed to anything yet.
if ($explicit -and -not $Force) {
    Write-Host ''

    # DEFAULTS TO YES, and that is a considered change rather than laziness.
    #
    # -Version means the operator has already typed the number they want. The
    # prompt is a chance to notice a typo, not a decision -- so Enter agreeing
    # with what was asked for is the right default, and it turned out to be the
    # common path by a wide margin in practice.
    #
    # This is safe to default open where the earlier drift prompt was not.
    # Stamping is reversible in the only way that matters: it writes
    # Version.json and nothing else, and re-running with a different number
    # replaces it. A wrong stamp is corrected by stamping again.
    #
    # 'n' still cancels, and -Force still skips the prompt entirely for
    # scripted use.
    $ans = (Read-Host "  Stamp this folder as v$Version? [Y/n]").Trim()
    if ($ans -match '^(n|no)$') {
        Write-Host '  Cancelled. Nothing was stamped (the unblock above still stands).' -ForegroundColor Yellow
        return
    }
}

Write-Step "4. Stamping as v$Version"

$hashes = [ordered]@{}
$failed = @()
foreach ($f in (Get-ToolkitCodeFile)) {
    $rel = Get-ToolkitRelativePath -FullPath $f
    $h   = Get-FileHashHex -Path $f
    if (-not $h) { $failed += $rel; continue }
    $hashes[$rel] = $h
    Write-Host ("  {0}  {1}" -f $h.Substring(0, 12), $rel)
}

if ($failed.Count -gt 0) {
    # Refusing rather than writing a partial manifest: a file silently absent
    # from the list is a file the integrity check will never complain about --
    # a hole shaped exactly like the thing it is meant to detect.
    Write-Host ''
    Write-Host "Could not hash $($failed.Count) file(s). Manifest NOT written:" -ForegroundColor Red
    foreach ($f in $failed) { Write-Host "    $f" -ForegroundColor Red }
    return
}

$manifest = [ordered]@{
    Version   = $Version
    Tenant    = $Tenant
    Released  = (Get-Date).ToString('yyyy-MM-dd')
    StampedBy = $env:USERNAME
    Notes     = $Notes
    FileCount = $hashes.Count
    Files     = $hashes
}

$path = Get-ToolkitVersionPath
Set-Content -LiteralPath $path -Value ($manifest | ConvertTo-Json -Depth 4) -Encoding UTF8

Write-Host ''
Write-Host "Wrote $path" -ForegroundColor Green
Write-Host "  Version : $Version"
Write-Host "  Files   : $($hashes.Count)"

# Verify the manifest against the files it just described. If this fails
# immediately after writing, the writer and the checker disagree about which
# files are covered -- a bug worth finding here rather than having it surface
# as phantom drift on someone else's machine, where it looks like their problem.
Write-Host ''
Write-Host 'Verifying the manifest against the files it just described...'
$check = Test-ToolkitIntegrity
if ($check.Modified.Count -eq 0 -and $check.Missing.Count -eq 0 -and $check.Extra.Count -eq 0) {
    Write-Host "  OK -- $($check.Matched) files match." -ForegroundColor Green
    Write-Host ''
    Write-Host 'Ready to ship.' -ForegroundColor Green
} else {
    Write-Host '  MISMATCH straight after writing -- writer and checker disagree.' -ForegroundColor Red
    Write-Host "  $($check.Summary)" -ForegroundColor Red
    Write-Host '  Do not ship this.' -ForegroundColor Red
}

#region Step 5 -- new tools on buttons ------------------------------------------
<#
    THE LAST STEP, AND THE RIGHT PLACE FOR IT.

    Adding a script to the folder and adding a BUTTON for it are two jobs, and
    the second one is easy to forget -- the file is there, the tester passes,
    the tool is simply unreachable. Publish is the moment both halves are in
    front of you: the new scripts are in the folder because you just stamped
    them, and this is the copy whose Config is the real one.

    AFTER THE STAMP, NOT BEFORE. Config\ is excluded from the manifest, so
    adding a button changes nothing that was hashed -- but running it first
    would mean a failed publish had still edited Tools.json, and a step that
    half-happens is worse than one that does not.

    ASKED, NEVER ASSUMED. It prints what it would do and waits. Answer no and
    the scripts stay in the folder, unreferenced, exactly as they were.
#>
$addScript = Join-Path $PSScriptRoot 'Add-ConsoleAdditions.ps1'
if (Test-Path -LiteralPath $addScript) {
    Write-Step 'New tools that are not on a button yet'
    try {
        & $addScript
        # The preview prints its own "nothing to do" and returns; the only way
        # to know whether it found work without parsing its output is to ask
        # it again briefly. Two cheap reads of the same JSON, and no shared
        # state between this script and that one -- which is what keeps the
        # tool usable on its own.
        $brief = & $addScript -Brief 6>&1 | Out-String
        if ($brief -match 'can be added') {
            Write-Host ''
            $ans = Read-Host 'Add these to the console now? (y/N)'
            if ("$ans".Trim() -match '^(?i)y') {
                & $addScript -Apply
            } else {
                Write-Host '  Left alone. Run Add-ConsoleAdditions.ps1 -Apply whenever you want them.' -ForegroundColor Yellow
            }
        }
    } catch {
        Write-Host "  Could not check for new tools: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host '  The stamp above is unaffected.' -ForegroundColor Yellow
    }
}
#endregion
