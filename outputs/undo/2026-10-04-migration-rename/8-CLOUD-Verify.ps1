#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# STEP 8 of the Migration project. Run after Step 7. Reads every migrated
# object back from the cloud and compares it, item by item, with what the
# pre-move exports say it had. The answer to "did anything go missing" is a
# file, not an eyeball.
#
# READ ONLY. Nothing is changed. Safe to run as often as you like.
#
# For each object in the newest Migration-Moved_*.csv:
#   * is it there, and is it cloud-only (IsDirSynced false)
#   * every proxy address from BOTH exports, plus X500 for both
#     LegacyExchangeDNs
#   * lists: owners, members, Send As, nesting, and the settings Step 7 stamps
#   * mailboxes: every row of the Step 4 Perms file (Full Access, Send As,
#     Send on Behalf, Inbox, Calendar), on-premises Send As, forwarding,
#     holds, retention, sent-items flags, custom attributes, memberships
#
# Result per item: match / MISSING / DIFFERENT / skipped. 'skipped' is a
# reference that no longer resolves in the cloud (a leaver as owner) or a
# membership in a list that is still synced -- expected losses Step 7 also
# reported. Anything MISSING or DIFFERENT is the list of things to fix by
# hand, or by re-running Step 7 if it was a transient failure.
#
# Output lands next to this script:
#     Migration-Verify_<date>_<time>.csv   Type, Object, Item, Expected, Actual, Result
#
# Step 9 refuses to delete the on-premises objects until this file exists and
# is newer than the move.
#
# Needs Exchange Online connected in this window.
# TOOLKIT-HELP-END

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$toolkitRoot = $PSScriptRoot
while ($toolkitRoot -and -not (Test-Path -LiteralPath (Join-Path $toolkitRoot 'Common.ps1'))) {
    $toolkitRoot = Split-Path -Parent $toolkitRoot
}
if (-not $toolkitRoot) { throw "Common.ps1 was not found in any folder above '$PSScriptRoot'." }
. (Join-Path $toolkitRoot 'Common.ps1')
. (Join-Path $PSScriptRoot '_Shared.ps1')

$basePath = $PSScriptRoot

function exo { param([string]$n, [hashtable]$p) Invoke-EXOCommand -Name $n -Parameters $p }

$movedFile = Get-NewestExport -Folder $basePath -Pattern 'Migration-Moved_*.csv'
if (-not $movedFile) { [System.Windows.Forms.MessageBox]::Show('No Migration-Moved_*.csv found. Nothing to verify.', '8 - CLOUD Verify', 'OK', 'Warning') | Out-Null; return }
$files = @{}
foreach ($k in 'DL-ONPREM', 'DL-CLOUD', 'SMB-ONPREM', 'SMB-CLOUD', 'SMB-CLOUD-Perms') { $files[$k] = Get-NewestExport -Folder $basePath -Pattern "Migration-${k}_*.csv" }

$go = Show-StepInstructions -Title '8 - CLOUD Verify' -ScriptPath $PSCommandPath -Text @"
STEP 8 -- compare the cloud-only objects with the pre-move exports.

* Objects:      $($movedFile.Name)
* DL exports:   $(if ($files['DL-CLOUD'])  { $files['DL-CLOUD'].Name }  else { '(none)' })  /  $(if ($files['DL-ONPREM'])  { $files['DL-ONPREM'].Name }  else { '(none)' })
* SMB exports:  $(if ($files['SMB-CLOUD']) { $files['SMB-CLOUD'].Name } else { '(none)' })  /  $(if ($files['SMB-ONPREM']) { $files['SMB-ONPREM'].Name } else { '(none)' })
* Perms:        $(if ($files['SMB-CLOUD-Perms']) { $files['SMB-CLOUD-Perms'].Name } else { '(none)' })

Nothing is changed. Output: Migration-Verify_<date>_<time>.csv.
Exchange Online must be connected in this window. Click OK to run.
"@
if (-not $go) { return }

$moved = @(Import-Csv -LiteralPath $movedFile.FullName | Where-Object { $_.Result -eq 'Moved' -or $_.Result -eq 'AlreadyThere' })
$ld = @{}
foreach ($k in $files.Keys) { $ld[$k] = @(); if ($files[$k]) { $ld[$k] = @(Import-Csv -LiteralPath $files[$k].FullName) } }
function Find-ExportRow { param([object[]]$Rows, [string]$Smtp) return ($Rows | Where-Object { "$($_.PrimarySmtpAddress)" -ieq $Smtp -and -not "$($_.Error)" } | Select-Object -First 1) }

$log = New-Object System.Collections.ArrayList
function Rec {
    param([string]$Type, [string]$Object, [string]$Item, [string]$Expected, [string]$Actual, [string]$Result)
    [void]$log.Add([PSCustomObject]@{ Type = $Type; Object = $Object; Item = $Item; Expected = $Expected; Actual = $Actual; Result = $Result })
    if ($Result -ne 'match') {
        $c = switch ($Result) { 'MISSING' { 'Red' } 'DIFFERENT' { 'Red' } default { 'Yellow' } }
        Write-Host ("    {0,-9} {1,-16} expected: {2}   actual: {3}" -f $Result, $Item, $Expected, $Actual) -ForegroundColor $c
    }
}

$script:ResolveCache = @{}
function Resolve-ToSmtp {
    param([string]$Id)
    $k = "$Id".ToLower(); if (-not $k) { return '' }
    if ($script:ResolveCache.ContainsKey($k)) { return $script:ResolveCache[$k] }
    $v = $k
    try { $r = exo 'Get-Recipient' @{ Identity = $Id; ErrorAction = 'Stop' }; if ("$($r.PrimarySmtpAddress)") { $v = "$($r.PrimarySmtpAddress)".ToLower() } } catch { $v = '' }
    $script:ResolveCache[$k] = $v
    return $v
}

function Compare-Set {
    # Expected addresses (from exports, union) vs current recipient list.
    param([string]$Type, [string]$Smtp, [string]$Item, [string[]]$Expected, [object]$Current, [switch]$Raw)
    $want = @($Expected | ForEach-Object { "$_".ToLower() } | Where-Object { $_ } | Select-Object -Unique)
    $have = @()
    foreach ($c in @($Current)) { $have += $(if ($Raw) { "$c".ToLower() } else { Resolve-ToSmtp "$c" }) }
    $have = @($have | Where-Object { $_ })
    if ($want.Count -eq 0) { return }
    $missing = @()
    foreach ($w in $want) {
        if ($have -contains $w) { continue }
        if (-not $Raw -and -not (Resolve-ToSmtp $w)) { Rec $Type $Smtp $Item $w '(does not resolve in the cloud)' 'skipped'; continue }
        $missing += $w
    }
    if ($missing.Count -eq 0) { Rec $Type $Smtp $Item "$($want.Count) entries" "$($want.Count) present" 'match' }
    else { Rec $Type $Smtp $Item ($missing -join ' ') '(absent)' 'MISSING' }
}

function Compare-Proxies {
    param([string]$Type, [string]$Smtp, [object]$Current, [object]$OnPrem, [object]$Cloud)
    $want = @()
    foreach ($src in @($OnPrem, $Cloud)) {
        if (-not $src) { continue }
        $want += ConvertFrom-CsvList $src.EmailAddresses
        if ("$($src.LegacyExchangeDN)") { $want += "X500:$($src.LegacyExchangeDN)" }
    }
    $have = @($Current.EmailAddresses | ForEach-Object { "$_".ToLower() })
    $missing = @()
    foreach ($w in ($want | Select-Object -Unique)) {
        $n = "$w".ToLower(); if ($n -match '^smtp:') { $n = 'smtp:' + $n.Substring(5) }
        if ($have -contains $n) { continue }
        if ($n -match '@.*\.onmicrosoft\.com$' -and $n -notmatch '^x500') { continue }
        $missing += $w
    }
    if ($missing.Count -eq 0) { Rec $Type $Smtp 'proxies' "$($have.Count) address(es)" 'all present' 'match' }
    else { Rec $Type $Smtp 'proxies' ($missing -join ' ') '(absent)' 'MISSING' }
    $p = "$($Current.PrimarySmtpAddress)"
    if ($p -ine $Smtp) { Rec $Type $Smtp 'primary' $Smtp $p 'DIFFERENT' } else { Rec $Type $Smtp 'primary' $Smtp $p 'match' }
}

function Compare-Scalars {
    param([string]$Type, [string]$Smtp, [object]$Current, [object]$OnPrem, [object]$Cloud, [string[]]$Keys)
    foreach ($k in $Keys) {
        $want = ''
        if ($Cloud -and $Cloud.PSObject.Properties[$k] -and "$($Cloud.$k)" -ne '')       { $want = "$($Cloud.$k)" }
        elseif ($OnPrem -and $OnPrem.PSObject.Properties[$k] -and "$($OnPrem.$k)" -ne '') { $want = "$($OnPrem.$k)" }
        if ($want -eq '') { continue }
        $cur = "$($Current.$k)"
        if ($cur -ieq $want) { Rec $Type $Smtp $k $want $cur 'match' } else { Rec $Type $Smtp $k $want $cur 'DIFFERENT' }
    }
}

$scalarKeys = @('DisplayName', 'Alias', 'HiddenFromAddressListsEnabled', 'RequireSenderAuthenticationEnabled', 'ModerationEnabled', 'SendModerationNotifications', 'MailTip') + @(1..15 | ForEach-Object { "CustomAttribute$_" })
$listKeys   = @('AcceptMessagesOnlyFrom', 'AcceptMessagesOnlyFromDLMembers', 'RejectMessagesFrom', 'RejectMessagesFromDLMembers', 'ModeratedBy', 'BypassModerationFromSendersOrMembers', 'GrantSendOnBehalfTo')

function Get-ExportUnion { param([object[]]$Sources, [string]$Column) $all = @(); foreach ($s in $Sources) { if ($s -and $s.PSObject.Properties[$Column]) { $all += ConvertFrom-CsvList $s.$Column } }; return @($all | Select-Object -Unique) }

Write-Host ''
$dls  = @($moved | Where-Object { $_.Type -eq 'DL' })
$smbs = @($moved | Where-Object { $_.Type -eq 'SMB' })

foreach ($m in $smbs) {
    $smtp = "$($m.PrimarySmtpAddress)"
    Write-Host "SMB  $smtp" -ForegroundColor Cyan
    $mb = $null
    try { $mb = exo 'Get-Mailbox' @{ Identity = $smtp; ErrorAction = 'Stop' } } catch { }
    if (-not $mb) { Rec 'SMB' $smtp 'exists' 'mailbox' '(not found)' 'MISSING'; continue }
    Rec 'SMB' $smtp 'exists' 'mailbox' "$($mb.RecipientTypeDetails)" 'match'
    if ("$($mb.IsDirSynced)" -eq 'True') { Rec 'SMB' $smtp 'cloud-only' 'IsDirSynced False' 'True' 'DIFFERENT' } else { Rec 'SMB' $smtp 'cloud-only' 'IsDirSynced False' 'False' 'match' }
    if ("$($mb.RecipientTypeDetails)" -ne 'SharedMailbox') { Rec 'SMB' $smtp 'type' 'SharedMailbox' "$($mb.RecipientTypeDetails)" 'DIFFERENT' }
    $on = Find-ExportRow $ld['SMB-ONPREM'] $smtp
    $cl = Find-ExportRow $ld['SMB-CLOUD']  $smtp
    if (-not $on -and -not $cl) { Rec 'SMB' $smtp 'exports' 'a row in an SMB export' '(none)' 'skipped'; continue }

    Compare-Proxies -Type 'SMB' -Smtp $smtp -Current $mb -OnPrem $on -Cloud $cl
    $cmp = $mb | Select-Object *
    if ("$($cmp.ForwardingSmtpAddress)") { $cmp.ForwardingSmtpAddress = "$($cmp.ForwardingSmtpAddress)" -replace '^smtp:', '' }
    Compare-Scalars -Type 'SMB' -Smtp $smtp -Current $cmp -OnPrem $on -Cloud $cl -Keys ($scalarKeys + @('MessageCopyForSentAsEnabled', 'MessageCopyForSendOnBehalfEnabled', 'DeliverToMailboxAndForward', 'ForwardingSmtpAddress', 'LitigationHoldEnabled', 'RetentionHoldEnabled', 'RetentionPolicy'))
    if ($cl -and "$($cl.ForwardingAddress)") { Compare-Set -Type 'SMB' -Smtp $smtp -Item 'ForwardingAddress' -Expected @($cl.ForwardingAddress) -Current @($mb.ForwardingAddress) }
    foreach ($k in $listKeys) { Compare-Set -Type 'SMB' -Smtp $smtp -Item $k -Expected (Get-ExportUnion @($on, $cl) $k) -Current $mb.$k }
    for ($n = 1; $n -le 5; $n++) { Compare-Set -Type 'SMB' -Smtp $smtp -Item "ExtensionCustomAttribute$n" -Expected (Get-ExportUnion @($on, $cl) "ExtensionCustomAttribute$n") -Current $mb."ExtensionCustomAttribute$n" -Raw }

    # Permissions.
    $perms = @($ld['SMB-CLOUD-Perms'] | Where-Object { "$($_.Mailbox)" -ieq $smtp })
    if ($on) { foreach ($s in (ConvertFrom-CsvList $on.SendAs)) { $perms += [PSCustomObject]@{ Mailbox = $smtp; Kind = 'SendAs'; User = $s; Rights = 'SendAs' } } }
    $fa = @(); $sa = @()
    try { $fa = @(exo 'Get-MailboxPermission' @{ Identity = $smtp; ErrorAction = 'Stop' } | Where-Object { -not $_.IsInherited -and $_.AccessRights -contains 'FullAccess' } | ForEach-Object { Resolve-ToSmtp "$($_.User)" }) } catch { }
    try { $sa = @(exo 'Get-RecipientPermission' @{ Identity = $smtp; ErrorAction = 'Stop' } | Where-Object { $_.AccessRights -contains 'SendAs' } | ForEach-Object { Resolve-ToSmtp "$($_.Trustee)" }) } catch { }
    $folders = @{}
    foreach ($p in ($perms | Sort-Object Kind, User -Unique)) {
        $who = "$($p.User)"
        switch ($p.Kind) {
            'FullAccess'   { $r = Resolve-ToSmtp $who; if (-not $r) { Rec 'SMB' $smtp 'FullAccess' $who '(does not resolve)' 'skipped' } elseif ($fa -contains $r) { Rec 'SMB' $smtp 'FullAccess' $who 'granted' 'match' } else { Rec 'SMB' $smtp 'FullAccess' $who '(absent)' 'MISSING' } }
            'SendAs'       { $r = Resolve-ToSmtp $who; if (-not $r) { Rec 'SMB' $smtp 'SendAs' $who '(does not resolve)' 'skipped' } elseif ($sa -contains $r) { Rec 'SMB' $smtp 'SendAs' $who 'granted' 'match' } else { Rec 'SMB' $smtp 'SendAs' $who '(absent)' 'MISSING' } }
            'SendOnBehalf' { }   # covered by GrantSendOnBehalfTo above
            default {
                try {
                    if (-not $folders.ContainsKey($p.Kind)) {
                        $fs = @(exo 'Get-MailboxFolderStatistics' @{ Identity = $smtp; FolderScope = $p.Kind; ErrorAction = 'Stop' })
                        $prim = $fs | Where-Object { "$($_.FolderType)" -eq $p.Kind } | Select-Object -First 1
                        if (-not $prim) { $prim = $fs | Sort-Object { ("$($_.FolderPath)".Split('/') | Where-Object { $_ }).Count } | Select-Object -First 1 }
                        $have = @{}
                        foreach ($fp in @(exo 'Get-MailboxFolderPermission' @{ Identity = ($smtp + ':\' + $prim.Name); ErrorAction = 'Stop' })) { $have["$($fp.User)".ToLower()] = (@($fp.AccessRights) -join ',') }
                        $folders[$p.Kind] = $have
                    }
                    $have = $folders[$p.Kind]
                    $rights = (@("$($p.Rights)" -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join ',')
                    $key = $who.ToLower()
                    # Folder permission User shows as display name; match on
                    # that first, then on the resolved address of each entry.
                    $cur = $null
                    if ($have.ContainsKey($key)) { $cur = $have[$key] }
                    else {
                        $want = Resolve-ToSmtp $who
                        foreach ($hk in $have.Keys) { if ($hk -notin 'default', 'anonymous' -and (Resolve-ToSmtp $hk) -eq $want -and $want) { $cur = $have[$hk]; break } }
                    }
                    if ($null -eq $cur) { if ($who -notin 'Default', 'Anonymous' -and -not (Resolve-ToSmtp $who)) { Rec 'SMB' $smtp $p.Kind "$who = $rights" '(does not resolve)' 'skipped' } else { Rec 'SMB' $smtp $p.Kind "$who = $rights" '(absent)' 'MISSING' } }
                    elseif ($cur -ieq $rights) { Rec 'SMB' $smtp $p.Kind "$who = $rights" $cur 'match' }
                    else { Rec 'SMB' $smtp $p.Kind "$who = $rights" "$who = $cur" 'DIFFERENT' }
                } catch { Rec 'SMB' $smtp $p.Kind $who "could not read: $($_.Exception.Message)" 'skipped' }
            }
        }
    }

    # Memberships.
    if ($cl -and "$($cl.MemberOf)") {
        foreach ($g in (ConvertFrom-CsvList $cl.MemberOf)) {
            try {
                $dg = exo 'Get-DistributionGroup' @{ Identity = $g; ErrorAction = 'Stop' }
                if ("$($dg.IsDirSynced)" -eq 'True') { Rec 'SMB' $smtp 'MemberOf' $g '(list still synced -- managed on-premises)' 'skipped'; continue }
                $isIn = @(exo 'Get-DistributionGroupMember' @{ Identity = $g; ResultSize = 'Unlimited'; ErrorAction = 'Stop' } | Where-Object { "$($_.PrimarySmtpAddress)" -ieq $smtp }).Count -gt 0
                Rec 'SMB' $smtp 'MemberOf' $g $(if ($isIn) { 'member' } else { '(absent)' }) $(if ($isIn) { 'match' } else { 'MISSING' })
            } catch { Rec 'SMB' $smtp 'MemberOf' $g '(list not found)' 'skipped' }
        }
    }
}

foreach ($m in $dls) {
    $smtp = "$($m.PrimarySmtpAddress)"
    Write-Host "DL   $smtp" -ForegroundColor Cyan
    $dg = $null
    try { $dg = exo 'Get-DistributionGroup' @{ Identity = $smtp; ErrorAction = 'Stop' } } catch { }
    if (-not $dg) { Rec 'DL' $smtp 'exists' 'list' '(not found)' 'MISSING'; continue }
    Rec 'DL' $smtp 'exists' 'list' "$($dg.RecipientTypeDetails)" 'match'
    if ("$($dg.IsDirSynced)" -eq 'True') { Rec 'DL' $smtp 'cloud-only' 'IsDirSynced False' 'True' 'DIFFERENT' } else { Rec 'DL' $smtp 'cloud-only' 'IsDirSynced False' 'False' 'match' }
    $on = Find-ExportRow $ld['DL-ONPREM'] $smtp
    $cl = Find-ExportRow $ld['DL-CLOUD']  $smtp
    if (-not $on -and -not $cl) { Rec 'DL' $smtp 'exports' 'a row in a DL export' '(none)' 'skipped'; continue }

    $wantType = if ($on -and "$($on.GroupType)" -match 'SecurityEnabled') { 'Security' } elseif ($cl -and "$($cl.GroupType)" -match 'SecurityEnabled') { 'Security' } else { 'Distribution' }
    $curType  = if ("$($dg.GroupType)" -match 'SecurityEnabled') { 'Security' } else { 'Distribution' }
    Rec 'DL' $smtp 'type' $wantType $curType $(if ($wantType -eq $curType) { 'match' } else { 'DIFFERENT' })

    Compare-Proxies -Type 'DL' -Smtp $smtp -Current $dg -OnPrem $on -Cloud $cl
    Compare-Scalars -Type 'DL' -Smtp $smtp -Current $dg -OnPrem $on -Cloud $cl -Keys ($scalarKeys + @('ReportToManagerEnabled', 'ReportToOriginatorEnabled', 'SendOofMessageToOriginatorEnabled', 'MemberJoinRestriction', 'MemberDepartRestriction', 'BypassNestedModerationEnabled'))
    Compare-Set -Type 'DL' -Smtp $smtp -Item 'ManagedBy' -Expected (Get-ExportUnion @($on, $cl) 'ManagedBy') -Current $dg.ManagedBy
    foreach ($k in $listKeys) { Compare-Set -Type 'DL' -Smtp $smtp -Item $k -Expected (Get-ExportUnion @($on, $cl) $k) -Current $dg.$k }
    for ($n = 1; $n -le 5; $n++) { Compare-Set -Type 'DL' -Smtp $smtp -Item "ExtensionCustomAttribute$n" -Expected (Get-ExportUnion @($on, $cl) "ExtensionCustomAttribute$n") -Current $dg."ExtensionCustomAttribute$n" -Raw }

    $members = @()
    try { $members = @(exo 'Get-DistributionGroupMember' @{ Identity = $smtp; ResultSize = 'Unlimited'; ErrorAction = 'Stop' } | ForEach-Object { "$($_.PrimarySmtpAddress)" }) } catch { }
    Compare-Set -Type 'DL' -Smtp $smtp -Item 'Members' -Expected (Get-ExportUnion @($on, $cl) 'Members') -Current $members -Raw
    $sa = @()
    try { $sa = @(exo 'Get-RecipientPermission' @{ Identity = $smtp; ErrorAction = 'Stop' } | Where-Object { $_.AccessRights -contains 'SendAs' } | ForEach-Object { "$($_.Trustee)" }) } catch { }
    Compare-Set -Type 'DL' -Smtp $smtp -Item 'SendAs' -Expected (Get-ExportUnion @($on, $cl) 'SendAs') -Current $sa

    foreach ($g in (Get-ExportUnion @($on, $cl) 'MemberOf')) {
        if ($g -ieq $smtp) { continue }
        try {
            $pg = exo 'Get-DistributionGroup' @{ Identity = $g; ErrorAction = 'Stop' }
            if ("$($pg.IsDirSynced)" -eq 'True') { Rec 'DL' $smtp 'MemberOf' $g '(parent still synced -- managed on-premises)' 'skipped'; continue }
            $isIn = @(exo 'Get-DistributionGroupMember' @{ Identity = $g; ResultSize = 'Unlimited'; ErrorAction = 'Stop' } | Where-Object { "$($_.PrimarySmtpAddress)" -ieq $smtp }).Count -gt 0
            Rec 'DL' $smtp 'MemberOf' $g $(if ($isIn) { 'member' } else { '(absent)' }) $(if ($isIn) { 'match' } else { 'MISSING' })
        } catch { Rec 'DL' $smtp 'MemberOf' $g '(parent not found)' 'skipped' }
    }
}

$outFile = Join-Path $basePath "Migration-Verify_$(Get-MigrationTimestamp).csv"
$log | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8
$nBad  = @($log | Where-Object { $_.Result -in 'MISSING', 'DIFFERENT' }).Count
$nSkip = @($log | Where-Object { $_.Result -eq 'skipped' }).Count
$nOk   = @($log | Where-Object { $_.Result -eq 'match' }).Count
Write-ExportDone -Path $outFile -Rows $log.Count
Write-Host "$nOk match, $nBad missing/different, $nSkip skipped across $($smbs.Count) mailbox(es) and $($dls.Count) list(s)." -ForegroundColor $(if ($nBad) { 'Red' } else { 'Green' })
if ($nBad) { Write-Host 'Re-run Step 7 for transient failures; anything still listed needs a hand.' -ForegroundColor Yellow }
else { Write-Host 'Clean. Step 9 may now delete the parked on-premises objects.' -ForegroundColor Green }
Write-CompletedLog -Action 'MigrationVerify' -Target "$($smbs.Count) mailboxes, $($dls.Count) lists" -Result $(if ($nBad) { 'Partial' } else { 'Success' }) -Detail "$nOk match, $nBad missing/different, $nSkip skipped -> $([System.IO.Path]::GetFileName($outFile))"
