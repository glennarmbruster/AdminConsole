#Requires -Version 5.1
# TOOLKIT-HELP-RUNS: nothing. It works on this folder's own files.
# TOOLKIT-HELP-BEGIN
# The last thing you run before handing the toolkit to another engineer.
#
#   .\Publish-Toolkit.ps1
#
# Four steps, in this order:
#
#   1. UNBLOCK every file. Anything downloaded or copied from another machine
#      carries a mark-of-the-web tag that makes PowerShell refuse to run it.
#      The files that were actually blocked are listed -- that list should
#      match what you just downloaded, and anything surprising on it is worth
#      a look before you ship.
#
#   2. SHOW what changed since the last stamp.
#
#   3. ASK whether this is a minor or a major release, and work out the number
#      from the last one -- 1.03 becomes 1.04, or 2.0. Nothing to type, so
#      nothing to mistype. Pass -Version to override.
#
#   4. STAMP Version.json with the new version and a SHA-256 per script.
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
    $ans = Read-Host "  Stamp this folder as v$Version? [y/N]"
    if ($ans -notmatch '^(y|yes)$') {
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
