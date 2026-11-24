#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# STEP 3 of the SMBMigration project. The CLOUD copy of Step 2: every
# Exchange Online attribute AND every permission of the batch's mailboxes.
#
# READ ONLY. Nothing is changed.
#
# Two files, because permissions are one row per grant and everything else
# is one row per mailbox:
#     SMBMigration-<batch>-CLOUD.csv        one row per mailbox
#     SMBMigration-<batch>-CLOUD-Perms.csv  Mailbox, Kind, User, Rights
# Kinds: FullAccess, SendAs, SendOnBehalf, Inbox, Calendar (folder rights by
# FolderType, never by name). Default and Anonymous are kept.
#
# The mailbox row carries proxies, the cloud LegacyExchangeDN, forwarding,
# holds, retention, archive, quotas, sent-items flags, audit, hidden, sender
# restrictions, custom attributes, regional settings and cloud MemberOf.
# DisplayName is recorded CLEAN (Step 1's tag removed); TaggedInCloud says
# whether the tag has synced.
#
# THE TAG GATE. It ends with READY or NOT READY: Step 4 should not run until
# every mailbox shows the tag in the cloud, or the admin-center search in
# Step 6 will miss some. Soft -- Step 4 warns and lets you override.
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

$batchId = Select-SmbBatch -Folder $basePath -Title '3 - CLOUD Export'
if (-not $batchId) { return }
$go = Show-StepInstructions -Title '3 - CLOUD Export' -ScriptPath $PSCommandPath -Text @"
STEP 3 of 9 -- export the CLOUD mailboxes of batch $batchId and every permission on them.

* Nothing is changed. This only reads.
* Output, next to this script:
    SMBMigration-$batchId-CLOUD.csv         one row per mailbox
    SMBMigration-$batchId-CLOUD-Perms.csv   one row per permission
* Ends with READY / NOT READY: whether Step 1's tag has synced to every mailbox.

Exchange Online must be connected in this window. Click OK to run.
"@
if (-not $go) { return }

function exo { param([string]$n, [hashtable]$p) Invoke-EXOCommand -Name $n -Parameters $p }

function Get-FolderIdentity {
    # Matched on FolderType, never on name -- names are localised.
    param([Parameter(Mandatory)][string]$Smtp, [Parameter(Mandatory)][ValidateSet('Inbox', 'Calendar')][string]$Kind)
    $folders = @(exo 'Get-MailboxFolderStatistics' @{ Identity = $Smtp; FolderScope = $Kind; ErrorAction = 'Stop' })
    $primary = $folders | Where-Object { "$($_.FolderType)" -eq $Kind } | Select-Object -First 1
    if (-not $primary) { $primary = $folders | Sort-Object { ("$($_.FolderPath)".Split('/') | Where-Object { $_ }).Count } | Select-Object -First 1 }
    if (-not $primary) { throw "No $Kind folder found on $Smtp." }
    return ($Smtp + ':\' + $primary.Name)
}

$names = @(Read-SmbBatch -Folder $basePath -BatchId $batchId)
Write-Host "$($names.Count) mailbox(es) in batch $batchId." -ForegroundColor Cyan
$outFile  = Get-SmbBatchFile -Folder $basePath -BatchId $batchId -Stage 'CLOUD'
$permFile = Get-SmbBatchFile -Folder $basePath -BatchId $batchId -Stage 'CLOUD-Perms'

$resolver = { param($id) Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{ Identity = $id; ErrorAction = 'Stop' } }

$rows = @(); $perms = @()
$i = 0
foreach ($name in $names) {
    $i++
    Write-Host "[$i/$($names.Count)] $name" -NoNewline
    try {
        $mb = exo 'Get-Mailbox' @{ Identity = $name; ErrorAction = 'Stop' }
        $smtp = "$($mb.PrimarySmtpAddress)"
        $pc = 0

        # Full Access: explicit, non-inherited, not SELF, not deny.
        try {
            foreach ($p in @(exo 'Get-MailboxPermission' @{ Identity = $smtp; ErrorAction = 'Stop' } |
                     Where-Object { -not $_.IsInherited -and -not $_.Deny -and "$($_.User)" -notmatch 'NT AUTHORITY\\SELF|^S-1-5' -and $_.AccessRights -contains 'FullAccess' })) {
                $perms += [PSCustomObject]@{ Mailbox = $smtp; Kind = 'FullAccess'; User = "$($p.User)"; Rights = 'FullAccess' }; $pc++
            }
        } catch { Write-Host "  (FullAccess not read: $($_.Exception.Message))" -ForegroundColor Yellow -NoNewline }

        # Send As in the cloud is Get-RecipientPermission.
        try {
            foreach ($p in @(exo 'Get-RecipientPermission' @{ Identity = $smtp; ErrorAction = 'Stop' } |
                     Where-Object { $_.AccessRights -contains 'SendAs' -and "$($_.Trustee)" -notmatch 'NT AUTHORITY\\SELF|^S-1-5' })) {
                $perms += [PSCustomObject]@{ Mailbox = $smtp; Kind = 'SendAs'; User = "$($p.Trustee)"; Rights = 'SendAs' }; $pc++
            }
        } catch { Write-Host "  (SendAs not read: $($_.Exception.Message))" -ForegroundColor Yellow -NoNewline }

        $sob = @(Get-RecipientSmtp -Refs $mb.GrantSendOnBehalfTo -Resolver $resolver)
        foreach ($s in $sob) { $perms += [PSCustomObject]@{ Mailbox = $smtp; Kind = 'SendOnBehalf'; User = $s; Rights = 'SendOnBehalf' }; $pc++ }

        foreach ($kind in 'Inbox', 'Calendar') {
            try {
                $fid = Get-FolderIdentity -Smtp $smtp -Kind $kind
                foreach ($p in @(exo 'Get-MailboxFolderPermission' @{ Identity = $fid; ErrorAction = 'Stop' })) {
                    $who = "$($p.User)"
                    # A user whose account is gone shows as a bare SID/'NT User:' --
                    # nothing to restamp, so skipped.
                    if ($who -match '^NT User:|^S-1-5') { continue }
                    $perms += [PSCustomObject]@{ Mailbox = $smtp; Kind = $kind; User = $who; Rights = (@($p.AccessRights) -join ',') }; $pc++
                }
            } catch { Write-Host "  ($kind perms not read: $($_.Exception.Message))" -ForegroundColor Yellow -NoNewline }
        }

        # Cloud lists this mailbox belongs to. Members -eq takes the DN.
        $memberOf = @()
        try {
            $dnQ = "$($mb.DistinguishedName)" -replace "'", "''"
            $memberOf = @(exo 'Get-DistributionGroup' @{ Filter = "Members -eq '$dnQ'"; ResultSize = 'Unlimited'; ErrorAction = 'Stop' } |
                          ForEach-Object { "$($_.PrimarySmtpAddress)" })
        } catch { Write-Host "  (MemberOf not read: $($_.Exception.Message))" -ForegroundColor Yellow -NoNewline }

        $cas = $null; $reg = $null
        try { $cas = exo 'Get-CASMailbox' @{ Identity = $smtp; ErrorAction = 'Stop' } } catch { }
        try { $reg = exo 'Get-MailboxRegionalConfiguration' @{ Identity = $smtp; ErrorAction = 'Stop' } } catch { }

        $fwd = ''
        if ("$($mb.ForwardingAddress)") { $fwd = (Get-RecipientSmtp -Refs $mb.ForwardingAddress -Resolver $resolver) -join '' }

        $row = [ordered]@{
            Name                                 = "$($mb.Name)"
            DisplayName                          = (Remove-SmbTag "$($mb.DisplayName)")
            TaggedDisplayName                    = "$($mb.DisplayName)"
            TaggedInCloud                        = $(if (Test-SmbTagged "$($mb.DisplayName)") { 'True' } else { 'False' })
            Alias                                = "$($mb.Alias)"
            UserPrincipalName                    = "$($mb.UserPrincipalName)"
            PrimarySmtpAddress                   = $smtp
            EmailAddresses                       = ConvertTo-CsvList ($mb.EmailAddresses | ForEach-Object { "$_" })
            LegacyExchangeDN                     = "$($mb.LegacyExchangeDN)"
            ExchangeGuid                         = "$($mb.ExchangeGuid)"
            ArchiveGuid                          = "$($mb.ArchiveGuid)"
            ArchiveStatus                        = "$($mb.ArchiveStatus)"
            ArchiveName                          = ConvertTo-CsvList $mb.ArchiveName
            ExternalDirectoryObjectId            = "$($mb.ExternalDirectoryObjectId)"
            ImmutableId                          = "$($mb.ImmutableId)"
            IsDirSynced                          = "$($mb.IsDirSynced)"
            RecipientTypeDetails                 = "$($mb.RecipientTypeDetails)"
            AccountDisabled                      = "$($mb.AccountDisabled)"
            HiddenFromAddressListsEnabled        = "$($mb.HiddenFromAddressListsEnabled)"
            ForwardingAddress                    = $fwd
            ForwardingSmtpAddress                = "$($mb.ForwardingSmtpAddress)" -replace '^smtp:', ''
            DeliverToMailboxAndForward           = "$($mb.DeliverToMailboxAndForward)"
            LitigationHoldEnabled                = "$($mb.LitigationHoldEnabled)"
            LitigationHoldDuration               = "$($mb.LitigationHoldDuration)"
            RetentionPolicy                      = "$($mb.RetentionPolicy)"
            RetentionHoldEnabled                 = "$($mb.RetentionHoldEnabled)"
            InPlaceHolds                         = ConvertTo-CsvList $mb.InPlaceHolds
            ProhibitSendQuota                    = "$($mb.ProhibitSendQuota)"
            ProhibitSendReceiveQuota             = "$($mb.ProhibitSendReceiveQuota)"
            IssueWarningQuota                    = "$($mb.IssueWarningQuota)"
            MessageCopyForSentAsEnabled          = "$($mb.MessageCopyForSentAsEnabled)"
            MessageCopyForSendOnBehalfEnabled    = "$($mb.MessageCopyForSendOnBehalfEnabled)"
            AuditEnabled                         = "$($mb.AuditEnabled)"
            RequireSenderAuthenticationEnabled   = "$($mb.RequireSenderAuthenticationEnabled)"
            AcceptMessagesOnlyFrom               = ConvertTo-CsvList (Get-RecipientSmtp -Refs $mb.AcceptMessagesOnlyFrom -Resolver $resolver)
            AcceptMessagesOnlyFromDLMembers      = ConvertTo-CsvList (Get-RecipientSmtp -Refs $mb.AcceptMessagesOnlyFromDLMembers -Resolver $resolver)
            RejectMessagesFrom                   = ConvertTo-CsvList (Get-RecipientSmtp -Refs $mb.RejectMessagesFrom -Resolver $resolver)
            RejectMessagesFromDLMembers          = ConvertTo-CsvList (Get-RecipientSmtp -Refs $mb.RejectMessagesFromDLMembers -Resolver $resolver)
            ModerationEnabled                    = "$($mb.ModerationEnabled)"
            ModeratedBy                          = ConvertTo-CsvList (Get-RecipientSmtp -Refs $mb.ModeratedBy -Resolver $resolver)
            GrantSendOnBehalfTo                  = ConvertTo-CsvList $sob
            MailTip                              = "$($mb.MailTip)"
            MemberOf                             = ConvertTo-CsvList $memberOf
            MemberOfCount                        = $memberOf.Count
            Language                             = $(if ($reg) { "$($reg.Language)" } else { '' })
            TimeZone                             = $(if ($reg) { "$($reg.TimeZone)" } else { '' })
            OwaEnabled                           = $(if ($cas) { "$($cas.OWAEnabled)" } else { '' })
            ActiveSyncEnabled                    = $(if ($cas) { "$($cas.ActiveSyncEnabled)" } else { '' })
            PopEnabled                           = $(if ($cas) { "$($cas.PopEnabled)" } else { '' })
            ImapEnabled                          = $(if ($cas) { "$($cas.ImapEnabled)" } else { '' })
            SmtpClientAuthenticationDisabled     = $(if ($cas) { "$($cas.SmtpClientAuthenticationDisabled)" } else { '' })
            PermissionCount                      = $pc
            WhenCreated                          = "$($mb.WhenCreated)"
        }
        for ($n = 1; $n -le 15; $n++) { $row["CustomAttribute$n"] = "$($mb."CustomAttribute$n")" }
        for ($n = 1; $n -le 5; $n++)  { $row["ExtensionCustomAttribute$n"] = ConvertTo-CsvList $mb."ExtensionCustomAttribute$n" }
        $row['ExportedUtc'] = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
        $row['Error'] = ''
        $rows += [PSCustomObject]$row
        Write-Host "  ok  ($pc permission(s), member of $($memberOf.Count))" -ForegroundColor Green
    } catch {
        Write-Host "  FAILED: $($_.Exception.Message)" -ForegroundColor Red
        $rows += [PSCustomObject]@{ Name = $name; Error = "$($_.Exception.Message)" }
    }
}

$header = @($rows | Where-Object { -not $_.Error } | Select-Object -First 1)
if ($header.Count -gt 0) {
    $cols = @($header[0].PSObject.Properties.Name)
    $rows = @($rows | ForEach-Object { $r = $_; $o = [ordered]@{}; foreach ($c in $cols) { $o[$c] = $(if ($r.PSObject.Properties[$c]) { $r.$c } else { '' }) }; [PSCustomObject]$o })
}
$rows  | Export-Csv -LiteralPath $outFile  -NoTypeInformation -Encoding UTF8
# The Perms file is always written, even empty, so Step 7 never has to guess
# whether "no file" means "no permissions" or "Step 3 was not run".
if ($perms.Count -eq 0) { '"Mailbox","Kind","User","Rights"' | Set-Content -LiteralPath $permFile -Encoding UTF8 }
else { $perms | Export-Csv -LiteralPath $permFile -NoTypeInformation -Encoding UTF8 }
Write-ExportDone -Path $outFile -Rows $rows.Count
Write-ExportDone -Path $permFile -Rows $perms.Count
$bad = @($rows | Where-Object { $_.Error }).Count
if ($bad) { Write-Host "$bad mailbox(es) failed -- see the Error column." -ForegroundColor Yellow }

# ---- things a cloud-only mailbox loses or needs, said now while there is still a choice ----
$goodRows = @($rows | Where-Object { -not $_.Error })
$held = @($goodRows | Where-Object { "$($_.LitigationHoldEnabled)" -eq 'True' -or "$($_.InPlaceHolds)".Trim() -or ("$($_.ArchiveStatus)" -and "$($_.ArchiveStatus)" -ne 'None') })
if ($held.Count) {
    Write-Host ''
    Write-Host "NOTE: $($held.Count) mailbox(es) have a litigation hold, an in-place hold or an online archive. Those need a licence; if the licence came from an on-premises group, it ends when the account is deleted in Step 4:" -ForegroundColor Yellow
    $held | ForEach-Object { Write-Host "  $($_.PrimarySmtpAddress)  hold=$($_.LitigationHoldEnabled) inplace=$(if ("$($_.InPlaceHolds)".Trim()) { 'yes' } else { 'no' }) archive=$($_.ArchiveStatus)" -ForegroundColor Yellow }
}
$inLists = @($goodRows | Where-Object { "$($_.MemberOf)".Trim() })
if ($inLists.Count) {
    Write-Host ''
    Write-Host "NOTE: $($inLists.Count) mailbox(es) are members of cloud lists. Membership of a list that is SYNCED from on-premises cannot be put back on a cloud-only mailbox (Step 7 reports it as skipped) -- mail sent to those lists will no longer reach these mailboxes unless the list itself is made cloud-only:" -ForegroundColor Yellow
    $inLists | ForEach-Object { Write-Host "  $($_.PrimarySmtpAddress)  <-  $($_.MemberOf)" -ForegroundColor Yellow }
}

# ---- the tag gate ----
$good = @($rows | Where-Object { -not $_.Error })
$untagged = @($good | Where-Object { "$($_.TaggedInCloud)" -ne 'True' })
Write-Host ''
if ($good.Count -and $untagged.Count -eq 0) {
    Write-Host "READY: the display-name tag from Step 1 is in the cloud on all $($good.Count) mailbox(es). Run Step 4." -ForegroundColor Green
} elseif ($good.Count) {
    Write-Host "NOT READY: $($untagged.Count) of $($good.Count) mailbox(es) do not yet show the Step 1 tag in the cloud:" -ForegroundColor Yellow
    $untagged | ForEach-Object { Write-Host "  $($_.PrimarySmtpAddress)" -ForegroundColor Yellow }
    Write-Host "Wait for the next sync and re-run this step. Step 4 will warn if it is run before this says READY -- the tag is what makes the admin-center restore one search." -ForegroundColor Yellow
}
Write-CompletedLog -Action 'SmbMigrationExportCloud' -Target "batch $batchId" -Result $(if ($bad) { 'Partial' } else { 'Success' }) -Detail "$($rows.Count - $bad) exported, $bad failed, $($perms.Count) permission rows -> $([System.IO.Path]::GetFileName($outFile))"
