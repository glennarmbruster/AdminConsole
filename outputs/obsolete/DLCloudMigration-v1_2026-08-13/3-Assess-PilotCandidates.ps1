#Requires -Version 5.1
<#
    3-Assess-PilotCandidates.ps1
    DL cloud migration -- which groups to migrate first.

    Reasoning: ScriptNotes/DLCloudMigration.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-08
#>
# TOOLKIT-HELP-BEGIN
# Scores every exported group for pilot suitability and proposes a pilot set.
# READ ONLY -- it reads the CSVs from 1-Export-Groups.ps1 and writes one more.
# It queries nothing and changes nothing.
#
# Run 1-Export-Groups.ps1 first. 2-Export-Usage.ps1 is optional but makes the
# scoring considerably better -- without it there is no evidence which groups
# are actually in use.
#
# TWO DIFFERENT QUESTIONS, and the output answers both:
#
#   PilotScore  -- how SAFE is this group to move? Small, owned, unentangled.
#   Pilot set   -- which groups should you ACTUALLY pilot? Not simply the
#                  safest, because a pilot made only of trivial groups proves
#                  only that trivial groups migrate. The proposed set is chosen
#                  to cover the FEATURES you will meet later -- moderation,
#                  Send-As, delivery restrictions, room lists -- at the lowest
#                  risk available for each.
# TOOLKIT-HELP-END

# Directory names are NOT in a ValidateSet.
#
# This said [ValidateSet('HQ','Franchise')] -- one organisation's internal
# structure written into a type constraint, in a script meant to be portable.
# It is the exact pattern removed from Get-ADServerString for the same reason,
# reintroduced here by the person who removed it.
#
# Two ways it failed. A tenant whose directories are named anything else could
# not run these at all. And an EMPTY value -- which the console used to pass
# before a lookup -- was rejected by the validator before the default could
# apply, so the script died on its own default value.
#
# Validated against Tenant.json at runtime instead, where the answer lives.
param(
    [string]$Environment = '',
    [int]$PilotSize = 12
)

$toolkitRoot = $PSScriptRoot
while ($toolkitRoot -and -not (Test-Path -LiteralPath (Join-Path $toolkitRoot 'Common.ps1'))) {
    $toolkitRoot = Split-Path -Parent $toolkitRoot
}
if (-not $toolkitRoot) { throw "Common.ps1 was not found above '$PSScriptRoot'." }
. (Join-Path $toolkitRoot 'Common.ps1')

# Resolve the directory before anything uses it.
$script:DirIds = @(@(Get-TenantDirectory) | ForEach-Object { "$($_.Id)" } | Where-Object { $_ })
if (-not "$Environment".Trim()) {
    $def = @(@(Get-TenantDirectory) | Where-Object { $_.IsDefault })
    if ($def.Count -gt 0) { $Environment = "$($def[0].Id)" }
    elseif ($script:DirIds.Count -gt 0) { $Environment = $script:DirIds[0] }
}
if ($script:DirIds.Count -gt 0 -and ($script:DirIds -notcontains $Environment)) {
    throw ("Directory '$Environment' is not defined in Tenant.json. " +
           "Configured directories: $($script:DirIds -join ', ')")
}
$envFolder = if ($Environment -eq 'HQ') { 'HQ' } else { 'Fran' }
$outDir    = Join-Path (Join-Path $PSScriptRoot 'Exports') $envFolder

#region Load one snapshot -------------------------------------------------------
# Everything is loaded from ONE timestamp, not from whatever is newest per file.
# Mixing a fresh Groups file with last week's Members would produce scores that
# look authoritative and describe a set of groups that never existed.
$groupsFile = @(Get-ChildItem -LiteralPath $outDir -Filter "Groups_$Environment`_*.csv" -File -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1)
if (@($groupsFile).Count -eq 0) {
    Write-Host ''
    Write-Host "No Groups_$Environment`_*.csv in $outDir. Run 1-Export-Groups.ps1 first." -ForegroundColor Red
    return
}
$stamp = ''
if ($groupsFile[0].Name -match "Groups_$Environment`_(.+)\.csv$") { $stamp = $Matches[1] }

function Import-Part {
    param([string]$Kind, [switch]$Optional)
    $p = Join-Path $outDir "$Kind`_$Environment`_$stamp.csv"
    if (-not (Test-Path -LiteralPath $p)) {
        if (-not $Optional) {
            Write-Host "  Missing from this snapshot: $Kind" -ForegroundColor Yellow
        }
        return @()
    }
    try { return @(Import-Csv -LiteralPath $p) } catch { return @() }
}

Write-Host ''
Write-Host "=== DL Cloud Migration -- pilot assessment ($Environment) ===" -ForegroundColor Cyan
Write-Host "Snapshot: $stamp" -ForegroundColor DarkGray
Write-Host ''

$groups   = Import-Part -Kind 'Groups'
$blockers = Import-Part -Kind 'Blockers'
$rules    = Import-Part -Kind 'TransportRuleRefs'
$perms    = Import-Part -Kind 'Permissions'
$members  = Import-Part -Kind 'Members'
$usage    = Import-Part -Kind 'Usage' -Optional

if ($groups.Count -eq 0) { Write-Host 'No groups in the snapshot.' -ForegroundColor Yellow; return }

$haveUsage = $usage.Count -gt 0
if (-not $haveUsage) {
    Write-Host '  No Usage file in this snapshot. Scoring will run, but with no' -ForegroundColor Yellow
    Write-Host '  evidence of which groups are live -- run 2-Export-Usage.ps1 for a' -ForegroundColor Yellow
    Write-Host '  materially better answer.' -ForegroundColor Yellow
    Write-Host ''
}

# Index by address once. Nested loops over 4,778 rows per group is the
# difference between seconds and half an hour when phase 2 arrives.
$byBlocker = @{}; foreach ($b in $blockers) { $byBlocker["$($b.GroupPrimarySmtp)".ToLower()] = $b }
$byUsage   = @{}; foreach ($u in $usage)    { $byUsage["$($u.GroupPrimarySmtp)".ToLower()]   = $u }
$ruleCount = @{}; foreach ($r in $rules)    { $k = "$($r.GroupPrimarySmtp)".ToLower(); $ruleCount[$k] = [int]$ruleCount[$k] + 1 }
$sendAs    = @{}; $sendOb = @{}
foreach ($p in $perms) {
    $k = "$($p.GroupPrimarySmtp)".ToLower()
    if ($p.PermissionType -eq 'SendAs') { $sendAs[$k] = [int]$sendAs[$k] + 1 } else { $sendOb[$k] = [int]$sendOb[$k] + 1 }
}
$badMembers = @{}
foreach ($m in $members) {
    if ("$($m.Reproducible)" -eq 'False') {
        $k = "$($m.GroupPrimarySmtp)".ToLower(); $badMembers[$k] = [int]$badMembers[$k] + 1
    }
}
#endregion

#region Score -------------------------------------------------------------------
function Get-Truthy { param($V) return ("$V" -match '(?i)^(true|yes|1)$') }

$rows = @()
foreach ($g in $groups) {
    $key = "$($g.PrimarySmtpAddress)".ToLower()
    $b   = $byBlocker[$key]
    $u   = $byUsage[$key]

    $memberCount = [int]"$($g.MemberCount)"
    $nested      = [int]"$($g.NestedGroupCount)"
    $owners      = [int]"$($g.OwnerCount)"
    $blockCount  = if ($b) { [int]"$($b.BlockerCount)" } else { 0 }
    $rulesHit    = [int]$ruleCount[$key]
    $saCount     = [int]$sendAs[$key]
    $sobCount    = [int]$sendOb[$key]
    $badMem      = [int]$badMembers[$key]

    $moderated   = Get-Truthy $g.ModerationEnabled
    $restricted  = (Get-Truthy $g.RequireSenderAuthenticationEnabled) -or
                   ("$($g.AcceptMessagesOnlyFrom)".Trim() -ne '') -or
                   ("$($g.AcceptMessagesOnlyFromDLMembers)".Trim() -ne '') -or
                   ("$($g.RejectMessagesFrom)".Trim() -ne '')
    $policyAddr  = Get-Truthy $g.EmailAddressPolicyEnabled
    $isRoomList  = ("$($g.RecipientTypeDetails)" -eq 'RoomList')
    $hidden      = Get-Truthy $g.HiddenFromAddressListsEnabled

    # Score starts at 100 and things that make a migration harder take points
    # off. Written as subtractions rather than a weighted sum so that every
    # number in the Why column can be traced back to one named reason.
    $score = 100
    $why   = @()

    if ($blockCount -gt 0) { $score -= 40; $why += "$blockCount blocker(s)" }
    if ($owners -eq 0)     { $score -= 25; $why += 'no owner to confirm success' }
    if ($nested -gt 0)     { $score -= 20; $why += "$nested nested group(s)" }
    if ($rulesHit -gt 0)   { $score -= 20; $why += "$rulesHit transport rule reference(s)" }
    if ($badMem -gt 0)     { $score -= 15; $why += "$badMem member(s) cannot be reproduced" }
    if ($moderated)        { $score -= 10; $why += 'moderation' }
    if ($policyAddr)       { $score -= 10; $why += 'policy-generated addresses' }
    if ($restricted)       { $score -= 5;  $why += 'delivery restrictions' }
    if ($saCount + $sobCount -gt 0) { $score -= 5; $why += "$($saCount + $sobCount) delegate(s)" }

    # Size. A pilot wants a group big enough that someone notices if it breaks
    # and small enough that not many people are affected when it does.
    if ($memberCount -eq 0)       { $score -= 30; $why += 'empty' }
    elseif ($memberCount -eq 1)   { $score -= 15; $why += 'single member' }
    elseif ($memberCount -gt 200) { $score -= 20; $why += "$memberCount members -- large blast radius" }
    elseif ($memberCount -gt 50)  { $score -= 8;  $why += "$memberCount members" }

    # Usage. A pilot on a dead list proves nothing: nobody would report a
    # failure, so a clean result is indistinguishable from no result.
    $usageVerdict = if ($u) { "$($u.Verdict)" } else { 'NotMeasured' }
    if ($haveUsage) {
        if ($usageVerdict -eq 'NoTrafficInWindow') { $score -= 25; $why += 'no traffic in the measured window -- a failure here would go unnoticed' }
        elseif ($usageVerdict -eq 'Light')         { $score += 5 }
        elseif ($usageVerdict -eq 'Active')        { $score += 10 }
    }

    if ($score -lt 0)   { $score = 0 }
    if ($score -gt 100) { $score = 100 }

    $recommend =
        if ($blockCount -gt 0)   { 'Blocked -- resolve first' }
        elseif ($score -ge 85)   { 'Pilot candidate' }
        elseif ($score -ge 65)   { 'Wave 1' }
        elseif ($score -ge 40)   { 'Wave 2' }
        else                     { 'Needs attention' }

    # Which migration FEATURES this group exercises. Used to pick a pilot set
    # that covers the work rather than one that avoids it.
    $features = @()
    if ($moderated)                 { $features += 'Moderation' }
    if ($restricted)                { $features += 'Restrictions' }
    if ($saCount -gt 0)             { $features += 'SendAs' }
    if ($sobCount -gt 0)            { $features += 'SendOnBehalf' }
    if ($isRoomList)                { $features += 'RoomList' }
    if ($hidden)                    { $features += 'Hidden' }
    if ($policyAddr)                { $features += 'PolicyAddresses' }
    if ("$($g.X500Addresses)".Trim()) { $features += 'ExistingX500' }
    if ($memberCount -gt 50)        { $features += 'LargeMembership' }
    if ($features.Count -eq 0)      { $features += 'Plain' }

    $rows += [PSCustomObject]@{
        PrimarySmtpAddress = "$($g.PrimarySmtpAddress)"
        Name               = "$($g.Name)"
        Type               = "$($g.RecipientTypeDetails)"
        PilotScore         = $score
        Recommendation     = $recommend
        Members            = $memberCount
        Owners             = $owners
        Nested             = $nested
        TransportRuleRefs  = $rulesHit
        SendAs             = $saCount
        SendOnBehalf       = $sobCount
        UnreproducibleMembers = $badMem
        Moderated          = [string]$moderated
        Restricted         = [string]$restricted
        PolicyAddresses    = [string]$policyAddr
        UsageVerdict       = $usageVerdict
        MessageCount       = if ($u) { "$($u.MessageCount)" } else { '' }
        Features           = ($features -join '; ')
        Why                = ($why -join '; ')
    }
}
#endregion

#region Propose a covering pilot set --------------------------------------------
# Greedy coverage, not top-N-by-score.
#
# Taking the twelve highest scores would produce twelve plain, small, unmoderated
# groups -- a pilot that succeeds and tells you nothing about the moderated
# group with four Send-As delegates you have to move in week three. So: repeatedly
# take the highest-scoring group that exercises a feature nothing chosen yet
# covers, and only fall back to plain high scorers once every feature is covered.
$eligible = @($rows | Where-Object { $_.Recommendation -ne 'Blocked -- resolve first' } |
              Sort-Object PilotScore -Descending)
$covered  = @{}
$pilot    = @()

foreach ($r in $eligible) {
    if ($pilot.Count -ge $PilotSize) { break }
    $new = @($r.Features -split ';\s*' | Where-Object { $_ -and -not $covered.ContainsKey($_) })
    if ($new.Count -gt 0) {
        foreach ($f in $new) { $covered[$f] = $true }
        $pilot += $r
    }
}
foreach ($r in $eligible) {
    if ($pilot.Count -ge $PilotSize) { break }
    if ($pilot -notcontains $r) { $pilot += $r }
}
$pilotKeys = @{}
foreach ($p in $pilot) { $pilotKeys["$($p.PrimarySmtpAddress)".ToLower()] = $true }
foreach ($r in $rows) {
    Add-Member -InputObject $r -NotePropertyName 'ProposedPilot' `
        -NotePropertyValue ([string][bool]$pilotKeys["$($r.PrimarySmtpAddress)".ToLower()]) -Force
}
#endregion

#region Write -------------------------------------------------------------------
$path = Join-Path $outDir "PilotCandidates_$Environment`_$stamp.csv"
$rows | Sort-Object PilotScore -Descending |
    Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8

$byRec = $rows | Group-Object Recommendation | Sort-Object Count -Descending
Write-Host "--- $Environment assessment ---" -ForegroundColor Cyan
foreach ($grp in $byRec) {
    Write-Host ("  {0,-26} {1}" -f $grp.Name, $grp.Count)
}
Write-Host ''
Write-Host "Proposed pilot set ($($pilot.Count) group(s)), chosen to cover the work:" -ForegroundColor Green
foreach ($p in $pilot) {
    Write-Host ("  {0,3}  {1,-42} {2,4} members  [{3}]" -f `
        $p.PilotScore, $p.PrimarySmtpAddress, $p.Members, $p.Features) -ForegroundColor Gray
}
Write-Host ''
Write-Host ("Features covered by the pilot: {0}" -f (($covered.Keys | Sort-Object) -join ', ')) -ForegroundColor DarkGray

$allFeatures = @{}
foreach ($r in $rows) { foreach ($f in ($r.Features -split ';\s*')) { if ($f) { $allFeatures[$f] = $true } } }
$missed = @($allFeatures.Keys | Where-Object { -not $covered.ContainsKey($_) } | Sort-Object)
if ($missed.Count -gt 0) {
    Write-Host ("NOT covered by the pilot: {0}" -f ($missed -join ', ')) -ForegroundColor Yellow
    Write-Host 'Those features appear only on groups that are blocked or low-scoring.' -ForegroundColor Yellow
    Write-Host 'You will meet them later with no pilot experience of them -- worth' -ForegroundColor Yellow
    Write-Host 'deliberately adding one, even at a lower score.' -ForegroundColor Yellow
}
Write-Host ''
Write-Host "Saved to: $path" -ForegroundColor Green
Write-Host ''

try {
    Write-CompletedLog -Action 'DLCloudMigration-Pilot' -Target $Environment -Result 'Success' `
        -Detail "$($rows.Count) scored, $($pilot.Count) proposed, snapshot $stamp"
} catch { }
#endregion
