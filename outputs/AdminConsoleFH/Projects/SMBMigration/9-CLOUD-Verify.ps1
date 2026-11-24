#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# STEP 9 of the SMBMigration project. Reads every mailbox in the batch back
# from Exchange Online and compares it, item by item, with the Step 2 and
# Step 3 exports. The answer to "did anything go missing" is a file, not an
# eyeball.
#
# READ ONLY. Nothing is changed. Safe to run as often as you like.
#
# Per mailbox: exists and is cloud-only (IsDirSynced false); is a
# SharedMailbox; DisplayName equals the CLEAN name (tag removed, Step 8);
# every proxy from both exports plus X500 for both LegacyExchangeDNs; every
# row of the Perms file (Full Access, Send As, Send on Behalf, Inbox,
# Calendar) plus on-premises Send As; forwarding, holds, retention,
# sent-items flags, sender restrictions, custom attributes; cloud list
# memberships.
#
# Result per item: match / MISSING / DIFFERENT / skipped. 'skipped' is a
# reference that no longer resolves in the cloud or a membership in a list
# that is still synced. Anything MISSING or DIFFERENT is the list of things
# to fix by hand, or by re-running Step 7 if it was transient.
#
# Output: SMBMigration-<batch>-Verify.csv. Needs Exchange Online.
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

$batchId = Select-SmbBatch -Folder $basePath -Title '9 - CLOUD Verify'
if (-not $batchId) { return }
$files = @{}
foreach ($k in 'ONPREM', 'CLOUD', 'CLOUD-Perms') {
    $p = Get-SmbBatchFile -Folder $basePath -BatchId $batchId -Stage $k
    $files[$k] = $(if (Test-Path -LiteralPath $p) { Get-Item -LiteralPath $p } else { $null })
}

$go = Show-StepInstructions -Title '9 - CLOUD Verify' -ScriptPath $PSCommandPath -Text @"
STEP 9 of 9 -- compare the restored mailboxes of batch $batchId with the exports.

* On-prem export:  $(if ($files['ONPREM'])      { $files['ONPREM'].Name }      else { '(MISSING)' })
* Cloud export:    $(if ($files['CLOUD'])       { $files['CLOUD'].Name }       else { '(MISSING)' })
* Perms:           $(if ($files['CLOUD-Perms']) { $files['CLOUD-Perms'].Name } else { '(MISSING)' })

Nothing is changed. Output: SMBMigration-$batchId-Verify.csv.
Exchange Online must be connected in this window. Click OK to run.
"@
if (-not $go) { return }

$names = @(Read-SmbBatch -Folder $basePath -BatchId $batchId)
$moved = @($names | ForEach-Object { [PSCustomObject]@{ Type = 'SMB'; PrimarySmtpAddress = $_ } })
$contacts = @()
$ctFile = Get-SmbBatchFile -Folder $basePath -BatchId $batchId -Stage 'Contacts'
if (Test-Path -LiteralPath $ctFile) { $contacts = @(Import-Csv -LiteralPath $ctFile) }
$onmsDomain = "$((Get-TenantConfig).OnMicrosoftDomain)".Trim()
$routingDomain = ''
try { $routingDomain = "$(Get-TenantRemoteRoutingDomain)".Trim() } catch { }
$ld = @{}
foreach ($k in $files.Keys) { $ld['SMB-' + $k] = @(); if ($files[$k]) { $ld['SMB-' + $k] = @(Import-Csv -LiteralPath $files[$k].FullName) } }
$ld['DL-ONPREM'] = @(); $ld['DL-CLOUD'] = @()
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
    # THE ON-PREMISES CONTACT DELIVERS TO <alias>@<tenant>.onmicrosoft.com.
    # If that address is not on the cloud mailbox, on-premises mail bounces
    # even though every other check passes -- so it is checked by itself.
    $ctRow = $contacts | Where-Object { "$($_.PrimarySmtpAddress)" -ieq $smtp } | Select-Object -First 1
    $routeTargets = @()
    if ($ctRow -and "$($ctRow.ExternalEmailAddress)") { $routeTargets += "$($ctRow.ExternalEmailAddress)" }
    elseif ("$($mb.Alias)" -and $onmsDomain) { $routeTargets += "$($mb.Alias)@$onmsDomain" }
    if ("$($mb.Alias)" -and $routingDomain) { $routeTargets += "$($mb.Alias)@$routingDomain" }
    $haveAddr = @($mb.EmailAddresses | ForEach-Object { ("$_" -replace '^(?i)smtp:', '').ToLower() })
    foreach ($rt in ($routeTargets | Select-Object -Unique)) {
        if ($haveAddr -contains $rt.ToLower()) { Rec 'SMB' $smtp 'routing' $rt 'present' 'match' }
        else { Rec 'SMB' $smtp 'routing' $rt '(absent -- mail via the on-premises contact will bounce)' 'MISSING' }
    }
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

$outFile = Get-SmbBatchFile -Folder $basePath -BatchId $batchId -Stage 'Verify'
$log | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8
$nBad  = @($log | Where-Object { $_.Result -in 'MISSING', 'DIFFERENT' }).Count
$nSkip = @($log | Where-Object { $_.Result -eq 'skipped' }).Count
$nOk   = @($log | Where-Object { $_.Result -eq 'match' }).Count
Write-ExportDone -Path $outFile -Rows $log.Count
Write-Host "$nOk match, $nBad missing/different, $nSkip skipped across $($smbs.Count) mailbox(es)." -ForegroundColor $(if ($nBad) { 'Red' } else { 'Green' })
if ($nBad) { Write-Host 'Re-run Step 7 for transient failures (and Step 8 for a DisplayName difference); anything still listed needs a hand.' -ForegroundColor Yellow }
else { Write-Host "Clean. Batch $batchId is done." -ForegroundColor Green }
Write-CompletedLog -Action 'SmbMigrationVerify' -Target "batch $batchId ($($smbs.Count) mailboxes)" -Result $(if ($nBad) { 'Partial' } else { 'Success' }) -Detail "$nOk match, $nBad missing/different, $nSkip skipped -> $([System.IO.Path]::GetFileName($outFile))"
