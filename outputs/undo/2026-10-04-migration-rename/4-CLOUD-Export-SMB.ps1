#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# STEP 4 of the Migration project. The CLOUD copy of Step 3: exports every
# Exchange Online attribute AND every permission of the shared mailboxes
# named in SMB-Input.csv.
#
# READ ONLY. Nothing is changed.
#
# THIS IS THE FILE THE RESTAMP LEANS ON. A restored mailbox comes back with
# its permissions intact, so most of this is a safety net -- but it is the
# safety net, and the one time a Full Access entry does not survive is the
# time somebody asks what it was. Two files, because permissions are one
# row per grant and everything else is one row per mailbox:
#
#     Migration-SMB-CLOUD_<date>_<time>.csv        one row per mailbox
#     Migration-SMB-CLOUD-Perms_<date>_<time>.csv  Mailbox, Kind, User, Rights
#
# Kinds in the Perms file: FullAccess, SendAs, SendOnBehalf, Inbox, Calendar.
# Inbox and Calendar are folder permissions found by FolderType, never by
# name (names are localised). Default and Anonymous are kept -- they are
# real settings.
#
# The mailbox row also carries: proxies, both LegacyExchangeDNs' worth of
# X500, forwarding, litigation hold, retention policy and hold, archive,
# quotas, sent-items copy flags, audit, hidden-from-GAL, sender
# restrictions, custom attributes, regional settings, and MemberOf -- every
# cloud list the mailbox is in.
#
# Needs Exchange Online connected in this window. Run Step 5 next -- but
# Step 5 is the one that starts the clock, so read its help first.
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

$go = Show-StepInstructions -Title '4 - CLOUD Export SMB' -ScriptPath $PSCommandPath -Text @"
STEP 4 of 7 -- export the CLOUD shared mailboxes and every permission on them.

* Reads SMB-Input.csv (same file as Step 3).
* Nothing is changed. This only reads.
* Output, next to this script:
    Migration-SMB-CLOUD_<date>_<time>.csv         one row per mailbox
    Migration-SMB-CLOUD-Perms_<date>_<time>.csv   one row per permission

Exchange Online must be connected in this window. Click OK to review the
input file and run.
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

$names = Read-InputList -Path (Join-Path $basePath 'SMB-Input.csv') -What 'mailbox'
$ts = Get-MigrationTimestamp
$outFile  = Join-Path $basePath "Migration-SMB-CLOUD_$ts.csv"
$permFile = Join-Path $basePath "Migration-SMB-CLOUD-Perms_$ts.csv"

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
            DisplayName                          = "$($mb.DisplayName)"
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
# whether "no file" means "no permissions" or "Step 4 was not run".
if ($perms.Count -eq 0) { '"Mailbox","Kind","User","Rights"' | Set-Content -LiteralPath $permFile -Encoding UTF8 }
else { $perms | Export-Csv -LiteralPath $permFile -NoTypeInformation -Encoding UTF8 }
Write-ExportDone -Path $outFile -Rows $rows.Count
Write-ExportDone -Path $permFile -Rows $perms.Count
$bad = @($rows | Where-Object { $_.Error }).Count
if ($bad) { Write-Host "$bad mailbox(es) failed -- see the Error column." -ForegroundColor Yellow }
Write-CompletedLog -Action 'MigrationExportSMB' -Target 'cloud' -Result $(if ($bad) { 'Partial' } else { 'Success' }) -Detail "$($rows.Count - $bad) exported, $bad failed, $($perms.Count) permission rows -> $([System.IO.Path]::GetFileName($outFile))"
