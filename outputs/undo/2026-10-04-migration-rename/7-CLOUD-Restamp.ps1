#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# STEP 7 of the Migration project. Puts back, on the cloud-only objects Step
# 6 brought back, everything the exports from Steps 1-4 say they had.
#
# IDEMPOTENT BY DESIGN. Every item is "does it have this? no? add it". Run
# it once, run it five times: the second run reports 'present' for
# everything and changes nothing. That is also why it is safe to run against
# a mailbox that came back intact -- it will find nothing to do and say so.
#
# FOR EACH LIST (recreated in Step 6, so it starts nearly empty):
#   * every proxy address from both exports, plus X500 for BOTH
#     LegacyExchangeDNs -- on-premises and old cloud. Without those, a
#     reply to any mail sent before today bounces with IMCEAEX.
#   * owners, members (union of both exports), Send As, and the lists it
#     was nested in (cloud-only parents only; synced parents are reported)
#   * hidden-from-GAL, sender restrictions, moderation, send on behalf,
#     report/OOF flags, MailTip, custom attributes 1-15, extension 1-5
#
# FOR EACH MAILBOX (restored intact in Step 6, so mostly verification):
#   * proxies and X500 as above
#   * Full Access, Send As, Send on Behalf, Inbox and Calendar folder rights
#     from the Step 4 Perms file, plus Send As granted on-premises (Step 3)
#   * forwarding, litigation hold, retention policy/hold, sent-items copy
#     flags, hidden-from-GAL, sender restrictions, MailTip, custom attributes
#   * membership of cloud lists it was in. A list that is STILL SYNCED cannot
#     be edited here and is reported -- that membership is managed
#     on-premises, and the object is no longer there. Migrate the list or
#     re-add on-premises with a contact.
#
# Values come from the CLOUD export where it exists (it is what the live
# object had) and the on-premises export otherwise; multi-valued things are
# the union of both. Anything that will not resolve -- an owner who has
# left, a member that no longer exists -- is skipped and listed, never
# invented.
#
# Output lands next to this script:
#     Migration-Restamp_<date>_<time>.csv   one row per item: present / added / set / FAILED / skipped
#
# Needs Exchange Online connected in this window. When the file has no
# FAILED or skipped rows you care about, move the exports to
# Exports-Complete\ and the project is done.
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

# ---- inputs ----
$movedFile = Get-NewestExport -Folder $basePath -Pattern 'Migration-Moved_*.csv'
if (-not $movedFile) { [System.Windows.Forms.MessageBox]::Show('No Migration-Moved_*.csv found. Run Step 5 (and 6) first.', '7 - CLOUD Restamp', 'OK', 'Warning') | Out-Null; return }
$files = @{}
foreach ($k in 'DL-ONPREM', 'DL-CLOUD', 'SMB-ONPREM', 'SMB-CLOUD', 'SMB-CLOUD-Perms') {
    $files[$k] = Get-NewestExport -Folder $basePath -Pattern "Migration-${k}_*.csv"
}

$go = Show-StepInstructions -Title '7 - CLOUD Restamp' -ScriptPath $PSCommandPath -Text @"
STEP 7 of 7 -- restamp the cloud-only objects from the exports.

* Objects:      $($movedFile.Name)
* DL on-prem:   $(if ($files['DL-ONPREM'])  { $files['DL-ONPREM'].Name }  else { '(none)' })
* DL cloud:     $(if ($files['DL-CLOUD'])   { $files['DL-CLOUD'].Name }   else { '(none)' })
* SMB on-prem:  $(if ($files['SMB-ONPREM']) { $files['SMB-ONPREM'].Name } else { '(none)' })
* SMB cloud:    $(if ($files['SMB-CLOUD'])  { $files['SMB-CLOUD'].Name }  else { '(none)' })
* SMB perms:    $(if ($files['SMB-CLOUD-Perms']) { $files['SMB-CLOUD-Perms'].Name } else { '(none)' })

Only what is missing is added; nothing is removed. Safe to re-run.
Exchange Online must be connected in this window. Click OK to run.
"@
if (-not $go) { return }

$moved = @(Import-Csv -LiteralPath $movedFile.FullName | Where-Object { $_.Result -eq 'Moved' -or $_.Result -eq 'AlreadyThere' })
$ld = @{}
foreach ($k in $files.Keys) { $ld[$k] = @(); if ($files[$k]) { $ld[$k] = @(Import-Csv -LiteralPath $files[$k].FullName) } }
function Find-ExportRow { param([object[]]$Rows, [string]$Smtp) return ($Rows | Where-Object { "$($_.PrimarySmtpAddress)" -ieq $Smtp -and -not "$($_.Error)" } | Select-Object -First 1) }

# ---- result log ----
$log = New-Object System.Collections.ArrayList
function Note {
    param([string]$Type, [string]$Object, [string]$Item, [string]$Action, [string]$Result, [string]$Detail = '')
    [void]$log.Add([PSCustomObject]@{ Type = $Type; Object = $Object; Item = $Item; Action = $Action; Result = $Result; Detail = $Detail })
    $c = switch ($Result) { 'added' { 'Green' } 'set' { 'Green' } 'present' { 'DarkGray' } 'FAILED' { 'Red' } default { 'Yellow' } }
    Write-Host ("    {0,-8} {1,-14} {2}{3}" -f $Result, $Item, $Action, $(if ($Detail) { "  ($Detail)" } else { '' })) -ForegroundColor $c
}

# ---- helpers ----
function Test-Resolvable { param([string]$Id) try { [void](exo 'Get-Recipient' @{ Identity = $Id; ErrorAction = 'Stop' }); return $true } catch { return $false } }

function Add-MissingProxies {
    <#
        Desired = every proxy from both exports + X500 for both
        LegacyExchangeDNs. The primary (SMTP: upper-case) is never touched:
        a second 'SMTP:' entry is refused, so it is compared case-blind and
        added as lower-case 'smtp:' when absent.
    #>
    param([string]$Type, [string]$Smtp, [object]$Current, [object]$OnPrem, [object]$Cloud, [string]$SetCmd)
    $want = @()
    foreach ($src in @($OnPrem, $Cloud)) {
        if (-not $src) { continue }
        $want += ConvertFrom-CsvList $src.EmailAddresses
        if ("$($src.LegacyExchangeDN)") { $want += "X500:$($src.LegacyExchangeDN)" }
    }
    $have = @($Current.EmailAddresses | ForEach-Object { "$_".ToLower() })
    $add = @()
    foreach ($w in ($want | Select-Object -Unique)) {
        $norm = "$w".ToLower()
        if ($norm -match '^smtp:') { $norm = 'smtp:' + $norm.Substring(5) }
        if ($have -contains $norm) { continue }
        if ($norm -match '@.*\.onmicrosoft\.com$' -and $norm -notmatch '^x500') { continue }   # tenant-managed
        $add += $(if ($w -match '^SMTP:') { 'smtp:' + $w.Substring(5) } else { $w })
    }
    if ($add.Count -eq 0) { Note $Type $Smtp 'proxies' "$($have.Count) address(es)" 'present'; return }
    $base = Get-SetBase -SetCmd $SetCmd -Smtp $Smtp
    try {
        exo $SetCmd ($base + @{ EmailAddresses = @{ Add = $add } }) | Out-Null
        Note $Type $Smtp 'proxies' ($add -join ' ') 'added'
    } catch {
        # One bad address fails the batch; retry singly so the good ones land.
        foreach ($a in $add) {
            try { exo $SetCmd ($base + @{ EmailAddresses = @{ Add = $a } }) | Out-Null; Note $Type $Smtp 'proxy' $a 'added' }
            catch { Note $Type $Smtp 'proxy' $a 'FAILED' $_.Exception.Message }
        }
    }
}

function Get-SetBase {
    # Set-DistributionGroup refuses a caller who is not an owner unless told
    # to bypass the check -- even an admin. Set-Mailbox has no such switch.
    param([string]$SetCmd, [string]$Smtp)
    $b = @{ Identity = $Smtp; ErrorAction = 'Stop' }
    if ($SetCmd -eq 'Set-DistributionGroup') { $b.BypassSecurityGroupManagerCheck = $true }
    return $b
}

$script:ResolveCache = @{}
function Resolve-ToSmtp {
    # A current recipient-list value (display or canonical name) as an address,
    # so it can be compared with the exports. Cached: owners repeat.
    param([string]$Id)
    $k = "$Id".ToLower()
    if (-not $k) { return '' }
    if ($script:ResolveCache.ContainsKey($k)) { return $script:ResolveCache[$k] }
    $v = $k
    try { $r = exo 'Get-Recipient' @{ Identity = $Id; ErrorAction = 'Stop' }; if ("$($r.PrimarySmtpAddress)") { $v = "$($r.PrimarySmtpAddress)".ToLower() } } catch { }
    $script:ResolveCache[$k] = $v
    return $v
}

function Set-MissingScalars {
    <#
        Desired scalars vs current, set only the differences, one call; if the
        batch is refused, one call per property so a single bad value does
        not block the rest. Recipient-list properties (AcceptMessagesOnlyFrom
        ...) are set as the union of the two exports, dropping anything that
        does not resolve.
    #>
    param([string]$Type, [string]$Smtp, [object]$Current, [hashtable]$Desired, [string]$SetCmd)
    $diff = @{}
    foreach ($k in $Desired.Keys) {
        $want = $Desired[$k]
        if ($null -eq $want) { continue }
        $cur = $Current.$k
        if ($want -is [array]) {
            $wantN = @($want | ForEach-Object { "$_".ToLower() } | Sort-Object -Unique)
            if ($wantN.Count -eq 0) { continue }                       # never blank a list from an empty export
            # Current recipient lists come back as names; the exports hold
            # addresses. Resolve before comparing or every owner is "missing".
            $curN = if ($k -like 'ExtensionCustomAttribute*') { @($cur | ForEach-Object { "$_".ToLower() }) } else { @($cur | ForEach-Object { Resolve-ToSmtp "$_" }) }
            $missing = @($wantN | Where-Object { $curN -notcontains $_ })
            if ($missing.Count -eq 0) { continue }
            $diff[$k] = @{ Add = @($want | Where-Object { $missing -contains "$_".ToLower() }) }
        } else {
            if ("$want" -eq '') { continue }                          # blank export value = leave alone
            $c = "$cur"; if ($cur -is [bool]) { $c = "$cur" }
            if ($c -ieq "$want") { continue }
            $v = $want
            if ("$want" -ieq 'True') { $v = $true } elseif ("$want" -ieq 'False') { $v = $false }
            $diff[$k] = $v
        }
    }
    if ($diff.Count -eq 0) { Note $Type $Smtp 'settings' "$($Desired.Count) checked" 'present'; return }
    $base = Get-SetBase -SetCmd $SetCmd -Smtp $Smtp
    try {
        exo $SetCmd ($base + $diff) | Out-Null
        Note $Type $Smtp 'settings' (($diff.Keys | Sort-Object) -join ', ') 'set'
    } catch {
        foreach ($k in $diff.Keys) {
            $one = $base.Clone(); $one[$k] = $diff[$k]
            $shown = if ($diff[$k] -is [hashtable]) { '+ ' + (@($diff[$k].Add) -join ', ') } else { "$($diff[$k])" }
            try { exo $SetCmd $one | Out-Null; Note $Type $Smtp $k $shown 'set' }
            catch { Note $Type $Smtp $k $shown 'FAILED' $_.Exception.Message }
        }
    }
}

function Get-UnionList {
    # Union of a ';' list column across the exports, resolvable entries only.
    param([object[]]$Sources, [string]$Column, [string]$Type, [string]$Smtp)
    $all = @()
    foreach ($s in $Sources) { if ($s -and $s.PSObject.Properties[$Column]) { $all += ConvertFrom-CsvList $s.$Column } }
    $out = @()
    foreach ($v in ($all | Select-Object -Unique)) {
        if (Test-Resolvable $v) { $out += $v } else { Note $Type $Smtp $Column $v 'skipped' 'does not resolve in the cloud' }
    }
    return @($out)
}

function Get-FolderIdentity {
    param([string]$Smtp, [string]$Kind)
    $folders = @(exo 'Get-MailboxFolderStatistics' @{ Identity = $Smtp; FolderScope = $Kind; ErrorAction = 'Stop' })
    $primary = $folders | Where-Object { "$($_.FolderType)" -eq $Kind } | Select-Object -First 1
    if (-not $primary) { $primary = $folders | Sort-Object { ("$($_.FolderPath)".Split('/') | Where-Object { $_ }).Count } | Select-Object -First 1 }
    if (-not $primary) { throw "No $Kind folder on $Smtp." }
    return ($Smtp + ':\' + $primary.Name)
}

$scalarKeys = @('HiddenFromAddressListsEnabled', 'RequireSenderAuthenticationEnabled', 'ModerationEnabled', 'SendModerationNotifications',
                'BypassNestedModerationEnabled', 'MailTip') + @(1..15 | ForEach-Object { "CustomAttribute$_" })
$listKeys   = @('AcceptMessagesOnlyFrom', 'AcceptMessagesOnlyFromDLMembers', 'RejectMessagesFrom', 'RejectMessagesFromDLMembers', 'ModeratedBy',
                'BypassModerationFromSendersOrMembers', 'GrantSendOnBehalfTo') + @(1..5 | ForEach-Object { "ExtensionCustomAttribute$_" })

function Get-DesiredSettings {
    # Cloud value wins for scalars; lists are the union. Extension attributes
    # are lists but not recipients, so they skip the resolve check.
    param([object]$OnPrem, [object]$Cloud, [string[]]$Scalars, [string[]]$Lists, [string]$Type, [string]$Smtp)
    $d = @{}
    foreach ($k in $Scalars) {
        $v = ''
        if ($Cloud -and $Cloud.PSObject.Properties[$k] -and "$($Cloud.$k)" -ne '')       { $v = "$($Cloud.$k)" }
        elseif ($OnPrem -and $OnPrem.PSObject.Properties[$k] -and "$($OnPrem.$k)" -ne '') { $v = "$($OnPrem.$k)" }
        if ($v -ne '') { $d[$k] = $v }
    }
    foreach ($k in $Lists) {
        if ($k -like 'ExtensionCustomAttribute*') {
            $all = @(); foreach ($s in @($OnPrem, $Cloud)) { if ($s -and $s.PSObject.Properties[$k]) { $all += ConvertFrom-CsvList $s.$k } }
            $d[$k] = @($all | Select-Object -Unique)
        } else {
            # @() again on the outside: a one-element return unrolls to a
            # string, and a string is not an array to Set-MissingScalars.
            $d[$k] = @(Get-UnionList -Sources @($OnPrem, $Cloud) -Column $k -Type $Type -Smtp $Smtp)
        }
    }
    return $d
}

# ---- go ----
Write-Host ''
$dls  = @($moved | Where-Object { $_.Type -eq 'DL' })
$smbs = @($moved | Where-Object { $_.Type -eq 'SMB' })

# MAILBOXES FIRST: list members and permission trustees include them.
foreach ($m in $smbs) {
    $smtp = "$($m.PrimarySmtpAddress)"
    Write-Host "SMB  $smtp" -ForegroundColor Cyan
    $mb = $null
    try { $mb = exo 'Get-Mailbox' @{ Identity = $smtp; ErrorAction = 'Stop' } } catch { }
    if (-not $mb) { Note 'SMB' $smtp 'mailbox' 'not found -- Step 6 not run for it?' 'skipped'; continue }
    if ("$($mb.IsDirSynced)" -eq 'True') { Note 'SMB' $smtp 'mailbox' 'still directory-synced -- restamp would be overwritten' 'skipped'; continue }
    $on = Find-ExportRow $ld['SMB-ONPREM'] $smtp
    $cl = Find-ExportRow $ld['SMB-CLOUD']  $smtp
    if (-not $on -and -not $cl) { Note 'SMB' $smtp 'exports' 'no row in either SMB export' 'skipped'; continue }

    Add-MissingProxies -Type 'SMB' -Smtp $smtp -Current $mb -OnPrem $on -Cloud $cl -SetCmd 'Set-Mailbox'

    $desired = Get-DesiredSettings -OnPrem $on -Cloud $cl -Scalars ($scalarKeys + @('MessageCopyForSentAsEnabled', 'MessageCopyForSendOnBehalfEnabled', 'DeliverToMailboxAndForward', 'LitigationHoldEnabled', 'RetentionHoldEnabled', 'RetentionPolicy')) -Lists $listKeys -Type 'SMB' -Smtp $smtp
    if ($cl) {
        if ("$($cl.ForwardingSmtpAddress)") { $desired['ForwardingSmtpAddress'] = "$($cl.ForwardingSmtpAddress)" }
        if ("$($cl.ForwardingAddress)" -and (Test-Resolvable $cl.ForwardingAddress)) { $desired['ForwardingAddress'] = "$($cl.ForwardingAddress)" }
        if ("$($cl.LitigationHoldDuration)" -and "$($cl.LitigationHoldDuration)" -ne 'Unlimited') { $desired['LitigationHoldDuration'] = "$($cl.LitigationHoldDuration)" }
    }
    # Compare forwarding as the cmdlet reports it.
    $cmp = $mb | Select-Object *
    if ("$($cmp.ForwardingSmtpAddress)") { $cmp.ForwardingSmtpAddress = "$($cmp.ForwardingSmtpAddress)" -replace '^smtp:', '' }
    Set-MissingScalars -Type 'SMB' -Smtp $smtp -Current $cmp -Desired $desired -SetCmd 'Set-Mailbox'

    # Permissions: Step 4 Perms file + on-premises Send As.
    $perms = @($ld['SMB-CLOUD-Perms'] | Where-Object { "$($_.Mailbox)" -ieq $smtp })
    if ($on) { foreach ($s in (ConvertFrom-CsvList $on.SendAs)) { $perms += [PSCustomObject]@{ Mailbox = $smtp; Kind = 'SendAs'; User = $s; Rights = 'SendAs' } } }
    if ($perms.Count) {
        $curFA = @(); $curSA = @()
        try { $curFA = @(exo 'Get-MailboxPermission' @{ Identity = $smtp; ErrorAction = 'Stop' } | Where-Object { -not $_.IsInherited -and $_.AccessRights -contains 'FullAccess' } | ForEach-Object { "$($_.User)".ToLower() }) } catch { }
        try { $curSA = @(exo 'Get-RecipientPermission' @{ Identity = $smtp; ErrorAction = 'Stop' } | Where-Object { $_.AccessRights -contains 'SendAs' } | ForEach-Object { "$($_.Trustee)".ToLower() }) } catch { }
        $folderCache = @{}
        foreach ($p in ($perms | Sort-Object Kind, User -Unique)) {
            $who = "$($p.User)"
            try {
                switch ($p.Kind) {
                    'FullAccess' {
                        if ($curFA -contains $who.ToLower()) { Note 'SMB' $smtp 'FullAccess' $who 'present'; break }
                        if (-not (Test-Resolvable $who)) { Note 'SMB' $smtp 'FullAccess' $who 'skipped' 'trustee does not resolve'; break }
                        exo 'Add-MailboxPermission' @{ Identity = $smtp; User = $who; AccessRights = 'FullAccess'; InheritanceType = 'All'; AutoMapping = $false; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                        Note 'SMB' $smtp 'FullAccess' $who 'added'
                    }
                    'SendAs' {
                        if ($curSA -contains $who.ToLower()) { Note 'SMB' $smtp 'SendAs' $who 'present'; break }
                        if (-not (Test-Resolvable $who)) { Note 'SMB' $smtp 'SendAs' $who 'skipped' 'trustee does not resolve'; break }
                        exo 'Add-RecipientPermission' @{ Identity = $smtp; Trustee = $who; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                        Note 'SMB' $smtp 'SendAs' $who 'added'
                    }
                    'SendOnBehalf' { <# handled by GrantSendOnBehalfTo in settings #> }
                    default {
                        # Inbox / Calendar folder rights.
                        if (-not $folderCache.ContainsKey($p.Kind)) {
                            $fid = Get-FolderIdentity -Smtp $smtp -Kind $p.Kind
                            $folderCache[$p.Kind] = @{ Id = $fid; Have = @{} }
                            foreach ($fp in @(exo 'Get-MailboxFolderPermission' @{ Identity = $fid; ErrorAction = 'Stop' })) { $folderCache[$p.Kind].Have["$($fp.User)".ToLower()] = (@($fp.AccessRights) -join ',') }
                        }
                        $fc = $folderCache[$p.Kind]
                        $rights = @("$($p.Rights)" -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                        $key = $who.ToLower()
                        if ($fc.Have.ContainsKey($key) -and $fc.Have[$key] -ieq ($rights -join ',')) { Note 'SMB' $smtp $p.Kind "$who = $($rights -join ',')" 'present'; break }
                        if ($who -notin 'Default', 'Anonymous' -and -not (Test-Resolvable $who)) { Note 'SMB' $smtp $p.Kind $who 'skipped' 'user does not resolve'; break }
                        if ($fc.Have.ContainsKey($key)) {
                            exo 'Set-MailboxFolderPermission' @{ Identity = $fc.Id; User = $who; AccessRights = $rights; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                        } else {
                            exo 'Add-MailboxFolderPermission' @{ Identity = $fc.Id; User = $who; AccessRights = $rights; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                        }
                        Note 'SMB' $smtp $p.Kind "$who = $($rights -join ',')" 'set'
                    }
                }
            } catch { Note 'SMB' $smtp $p.Kind $who 'FAILED' $_.Exception.Message }
        }
    }

    # Cloud list memberships.
    if ($cl -and "$($cl.MemberOf)") {
        foreach ($g in (ConvertFrom-CsvList $cl.MemberOf)) {
            try {
                $dg = exo 'Get-DistributionGroup' @{ Identity = $g; ErrorAction = 'Stop' }
                if ("$($dg.IsDirSynced)" -eq 'True') { Note 'SMB' $smtp 'MemberOf' $g 'skipped' 'list is still directory-synced -- membership is managed on-premises'; continue }
                $isIn = @(exo 'Get-DistributionGroupMember' @{ Identity = $g; ResultSize = 'Unlimited'; ErrorAction = 'Stop' } | Where-Object { "$($_.PrimarySmtpAddress)" -ieq $smtp }).Count -gt 0
                if ($isIn) { Note 'SMB' $smtp 'MemberOf' $g 'present'; continue }
                exo 'Add-DistributionGroupMember' @{ Identity = $g; Member = $smtp; BypassSecurityGroupManagerCheck = $true; ErrorAction = 'Stop' } | Out-Null
                Note 'SMB' $smtp 'MemberOf' $g 'added'
            } catch { Note 'SMB' $smtp 'MemberOf' $g 'FAILED' $_.Exception.Message }
        }
    }
}

foreach ($m in $dls) {
    $smtp = "$($m.PrimarySmtpAddress)"
    Write-Host "DL   $smtp" -ForegroundColor Cyan
    $dg = $null
    try { $dg = exo 'Get-DistributionGroup' @{ Identity = $smtp; ErrorAction = 'Stop' } } catch { }
    if (-not $dg) { Note 'DL' $smtp 'list' 'not found -- Step 6 not run for it?' 'skipped'; continue }
    if ("$($dg.IsDirSynced)" -eq 'True') { Note 'DL' $smtp 'list' 'still directory-synced -- restamp would be overwritten' 'skipped'; continue }
    $on = Find-ExportRow $ld['DL-ONPREM'] $smtp
    $cl = Find-ExportRow $ld['DL-CLOUD']  $smtp
    if (-not $on -and -not $cl) { Note 'DL' $smtp 'exports' 'no row in either DL export' 'skipped'; continue }

    Add-MissingProxies -Type 'DL' -Smtp $smtp -Current $dg -OnPrem $on -Cloud $cl -SetCmd 'Set-DistributionGroup'

    $desired = Get-DesiredSettings -OnPrem $on -Cloud $cl -Scalars ($scalarKeys + @('ReportToManagerEnabled', 'ReportToOriginatorEnabled', 'SendOofMessageToOriginatorEnabled', 'MemberJoinRestriction', 'MemberDepartRestriction')) -Lists ($listKeys + @('ManagedBy')) -Type 'DL' -Smtp $smtp
    Set-MissingScalars -Type 'DL' -Smtp $smtp -Current $dg -Desired $desired -SetCmd 'Set-DistributionGroup'

    # Members: union of both exports, add what is missing.
    $want = @()
    foreach ($s in @($on, $cl)) { if ($s) { $want += ConvertFrom-CsvList $s.Members } }
    $want = @($want | Select-Object -Unique)
    $have = @()
    try { $have = @(exo 'Get-DistributionGroupMember' @{ Identity = $smtp; ResultSize = 'Unlimited'; ErrorAction = 'Stop' } | ForEach-Object { "$($_.PrimarySmtpAddress)".ToLower() }) } catch { }
    $addM = @($want | Where-Object { $have -notcontains "$_".ToLower() })
    if ($addM.Count -eq 0) { Note 'DL' $smtp 'members' "$($have.Count) present, $($want.Count) wanted" 'present' }
    foreach ($mem in $addM) {
        try {
            exo 'Add-DistributionGroupMember' @{ Identity = $smtp; Member = $mem; BypassSecurityGroupManagerCheck = $true; ErrorAction = 'Stop' } | Out-Null
            Note 'DL' $smtp 'member' $mem 'added'
        } catch {
            $why = if ($_.Exception.Message -match "couldn't be found|not found|doesn't exist") { 'does not resolve in the cloud' } else { $_.Exception.Message }
            Note 'DL' $smtp 'member' $mem $(if ($why -match 'resolve') { 'skipped' } else { 'FAILED' }) $why
        }
    }

    # Send As: union of both exports.
    $wantSA = @(); foreach ($s in @($on, $cl)) { if ($s) { $wantSA += ConvertFrom-CsvList $s.SendAs } }
    if ($wantSA.Count) {
        $curSA = @()
        try { $curSA = @(exo 'Get-RecipientPermission' @{ Identity = $smtp; ErrorAction = 'Stop' } | Where-Object { $_.AccessRights -contains 'SendAs' } | ForEach-Object { "$($_.Trustee)".ToLower() }) } catch { }
        foreach ($t in ($wantSA | Select-Object -Unique)) {
            if ($curSA -contains $t.ToLower()) { Note 'DL' $smtp 'SendAs' $t 'present'; continue }
            if (-not (Test-Resolvable $t)) { Note 'DL' $smtp 'SendAs' $t 'skipped' 'trustee does not resolve'; continue }
            try { exo 'Add-RecipientPermission' @{ Identity = $smtp; Trustee = $t; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop' } | Out-Null; Note 'DL' $smtp 'SendAs' $t 'added' }
            catch { Note 'DL' $smtp 'SendAs' $t 'FAILED' $_.Exception.Message }
        }
    }

    # Nesting: put this list back into the lists it was a member of. Union of
    # both exports; a parent that is still synced is reported, not touched.
    $parents = @(); foreach ($s in @($on, $cl)) { if ($s -and $s.PSObject.Properties['MemberOf']) { $parents += ConvertFrom-CsvList $s.MemberOf } }
    foreach ($g in ($parents | Select-Object -Unique)) {
        if ($g -ieq $smtp) { continue }
        try {
            $pg = exo 'Get-DistributionGroup' @{ Identity = $g; ErrorAction = 'Stop' }
            if ("$($pg.IsDirSynced)" -eq 'True') { Note 'DL' $smtp 'MemberOf' $g 'skipped' 'parent list is still directory-synced -- membership is managed on-premises'; continue }
            $isIn = @(exo 'Get-DistributionGroupMember' @{ Identity = $g; ResultSize = 'Unlimited'; ErrorAction = 'Stop' } | Where-Object { "$($_.PrimarySmtpAddress)" -ieq $smtp }).Count -gt 0
            if ($isIn) { Note 'DL' $smtp 'MemberOf' $g 'present'; continue }
            exo 'Add-DistributionGroupMember' @{ Identity = $g; Member = $smtp; BypassSecurityGroupManagerCheck = $true; ErrorAction = 'Stop' } | Out-Null
            Note 'DL' $smtp 'MemberOf' $g 'added'
        } catch { Note 'DL' $smtp 'MemberOf' $g 'FAILED' $_.Exception.Message }
    }
}

# ---- write the log ----
$outFile = Join-Path $basePath "Migration-Restamp_$(Get-MigrationTimestamp).csv"
$log | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8
$nF = @($log | Where-Object { $_.Result -eq 'FAILED' }).Count
$nS = @($log | Where-Object { $_.Result -eq 'skipped' }).Count
$nA = @($log | Where-Object { $_.Result -in 'added', 'set' }).Count
Write-ExportDone -Path $outFile -Rows $log.Count
Write-Host "$nA change(s), $nF failed, $nS skipped. Skipped rows name things that no longer resolve -- read them before calling it done." -ForegroundColor $(if ($nF) { 'Yellow' } else { 'Cyan' })
Write-CompletedLog -Action 'MigrationRestamp' -Target "$($smbs.Count) mailboxes, $($dls.Count) lists" -Result $(if ($nF -and $nA) { 'Partial' } elseif ($nF) { 'Failed' } else { 'Success' }) -Detail "$nA changed, $nF failed, $nS skipped -> $([System.IO.Path]::GetFileName($outFile))"
