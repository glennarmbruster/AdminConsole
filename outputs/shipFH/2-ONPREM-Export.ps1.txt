#Requires -Version 5.1
# TOOLKIT-REQUIRES: OwnProcess  -- Get-Recipient and Get-User exist in BOTH
#                                 Exchanges; Get-RemoteMailbox is on-premises
#                                 only. This script means the on-premises
#                                 copies and must not share a process with a
#                                 cloud-connected console.
# TOOLKIT-HELP-BEGIN
# STEP 2 of the SMBMigration project. Exports every ON-PREMISES attribute of
# the shared mailboxes in the batch.
#
# READ ONLY. Nothing is changed.
#
# The on-premises object is a REMOTE MAILBOX: an AD account with mail
# attributes and a pointer to the cloud. What only exists here and is
# captured: the AD account (sam, UPN, OU, description...), every proxy
# address and the on-premises LegacyExchangeDN (both become X500 later),
# Send As granted on-premises, sender restrictions, custom attributes, and
# MemberOf. DisplayName is recorded CLEAN (Step 1's tag removed) and as
# tagged, so Steps 5, 8 and 9 use the real name.
#
# Step 4 deletes these objects, so this file is the only on-premises record
# of them afterwards. Keep it.
#
# Output: SMBMigration-<batch>-ONPREM.csv, and SMBMigration-<batch>-Groups.csv
# with one row per group membership (name, mail, category, scope, DN --
# security groups included) for cleanup after the account is gone.
# Run Step 3 (the cloud copy) next.
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

$batchId = Select-SmbBatch -Folder $basePath -Title '2 - ONPREM Export'
if (-not $batchId) { return }
$go = Show-StepInstructions -Title '2 - ONPREM Export' -ScriptPath $PSCommandPath -Text @"
STEP 2 of 9 -- export the ON-PREMISES shared mailbox accounts of batch $batchId.

* Nothing is changed. This only reads.
* Output: SMBMigration-$batchId-ONPREM.csv next to this script -- after
  Step 4 this is the only on-premises record of these objects.
* DisplayName is recorded clean (Step 1's tag removed) and as tagged.

This window must be connected to ON-PREMISES Exchange. Click OK to run.
"@
if (-not $go) { return }

Import-Module ActiveDirectory -ErrorAction SilentlyContinue
Enter-MigrationOnPremScope

$names = @(Read-SmbBatch -Folder $basePath -BatchId $batchId)
Write-Host "$($names.Count) mailbox(es) in batch $batchId." -ForegroundColor Cyan
$outFile = Get-SmbBatchFile -Folder $basePath -BatchId $batchId -Stage 'ONPREM'

$resolver = { param($id) Get-Recipient -Identity $id -ErrorAction Stop }

$rows = @()
$groupRows = @()
$i = 0
foreach ($name in $names) {
    $i++
    Write-Host "[$i/$($names.Count)] $name" -NoNewline
    try {
        # Get-RemoteMailbox is on-premises ONLY. Its absence from the cloud is
        # the point: this is the object that decides what the cloud copy looks
        # like -- until Step 4 deletes it on purpose.
        $rm = Get-RemoteMailbox -Identity $name -ErrorAction Stop
        # Every AD read below goes to a controller in THIS object's domain.
        $srv = Get-MigrationAdServer -DistinguishedName "$($rm.DistinguishedName)"
        $adp = @{}; $xdc = @{}
        if ($srv) { $adp['Server'] = $srv; $xdc['DomainController'] = $srv }
        $u  = $null
        try { $u = Get-User -Identity $rm.Identity -ErrorAction Stop } catch { }

        $ad = $null; $memberOf = @()
        try {
            $ad = Get-ADUser -Identity $rm.DistinguishedName @adp -Properties memberOf, Enabled, Description, Department, Title, Manager, Company, Office, telephoneNumber, whenCreated, whenChanged, msExchRecipientTypeDetails, msExchRemoteRecipientType -ErrorAction Stop
            foreach ($gdn in @($ad.memberOf)) {
                # Address if the group has one, name if not -- same rule as
                # every other reference in these exports.
                $g = $null
                $gsrv = Get-MigrationAdServer -DistinguishedName "$gdn"
                $gp = @{}; if ($gsrv) { $gp['Server'] = $gsrv }
                try { $g = Get-ADGroup -Identity $gdn @gp -Properties mail -ErrorAction Stop } catch { }
                $memberOf += $(if ($g -and "$($g.mail)") { "$($g.mail)" } elseif ($g) { "$($g.Name)" } else { "$gdn" })
                # ONE ROW PER MEMBERSHIP, with the DN, for cleanup later: the
                # account is deleted in Step 4 and AD forgets it was ever in
                # these groups. Security groups included -- they are the ones
                # with no address and nothing else records them.
                $groupRows += [PSCustomObject]@{
                    Batch = $batchId; PrimarySmtpAddress = "$($rm.PrimarySmtpAddress)"; Name = "$($rm.Name)"
                    GroupName = $(if ($g) { "$($g.Name)" } else { '' }); GroupMail = $(if ($g) { "$($g.mail)" } else { '' })
                    GroupCategory = $(if ($g) { "$($g.GroupCategory)" } else { '' }); GroupScope = $(if ($g) { "$($g.GroupScope)" } else { '' })
                    GroupDn = "$gdn"; Resolved = $(if ($g) { 'True' } else { 'False' })
                }
            }
        } catch { Write-Host "  (AD read failed: $($_.Exception.Message))" -ForegroundColor Yellow -NoNewline }

        $sendAs = @()
        try {
            $aces = @(Get-ADPermission -Identity $rm.DistinguishedName @xdc -ErrorAction Stop |
                      Where-Object { -not $_.IsInherited -and $_.ExtendedRights -match 'Send-As' -and "$($_.User)" -notmatch 'NT AUTHORITY\\SELF|^S-1-5' })
            foreach ($a in $aces) {
                $sam = ("$($a.User)" -split '\\')[-1]
                $r = $null
                try { $r = Get-Recipient -Identity $sam -ErrorAction Stop } catch { }
                $sendAs += $(if ($r) { "$($r.PrimarySmtpAddress)" } else { "$($a.User)" })
            }
        } catch { Write-Host "  (Send As not read: $($_.Exception.Message))" -ForegroundColor Yellow -NoNewline }

        $mgr = ''
        if ($ad -and "$($ad.Manager)") {
            $msrv = Get-MigrationAdServer -DistinguishedName "$($ad.Manager)"
            $mp = @{}; if ($msrv) { $mp['Server'] = $msrv }
            try { $mgr = "$((Get-ADUser -Identity $ad.Manager @mp -Properties mail -ErrorAction Stop).mail)" } catch { $mgr = "$($ad.Manager)" }
        }

        $row = [ordered]@{
            Name                               = "$($rm.Name)"
            DisplayName                        = (Remove-SmbTag "$($rm.DisplayName)")
            TaggedDisplayName                  = "$($rm.DisplayName)"
            TaggedOnPrem                       = $(if (Test-SmbTagged "$($rm.DisplayName)") { 'True' } else { 'False' })
            Alias                              = "$($rm.Alias)"
            SamAccountName                     = "$($rm.SamAccountName)"
            UserPrincipalName                  = "$($rm.UserPrincipalName)"
            PrimarySmtpAddress                 = "$($rm.PrimarySmtpAddress)"
            EmailAddresses                     = ConvertTo-CsvList ($rm.EmailAddresses | ForEach-Object { "$_" })   # "$_", not .ProxyAddressString: deserialised entries have no such property (2026-10-04, empty column)
            RemoteRoutingAddress               = "$($rm.RemoteRoutingAddress)"
            LegacyExchangeDN                   = "$($rm.LegacyExchangeDN)"
            ExchangeGuid                       = "$($rm.ExchangeGuid)"
            ArchiveGuid                        = "$($rm.ArchiveGuid)"
            RecipientTypeDetails               = "$($rm.RecipientTypeDetails)"
            RemoteRecipientType                = "$($rm.RemoteRecipientType)"
            Guid                               = "$($rm.Guid)"
            DistinguishedName                  = "$($rm.DistinguishedName)"
            OrganizationalUnit                 = "$($rm.OrganizationalUnit)"
            AccountEnabled                     = $(if ($ad) { "$($ad.Enabled)" } else { '' })
            Description                        = $(if ($ad) { "$($ad.Description)" } else { '' })
            Department                         = $(if ($ad) { "$($ad.Department)" } else { '' })
            Title                              = $(if ($ad) { "$($ad.Title)" } else { '' })
            Company                            = $(if ($ad) { "$($ad.Company)" } else { '' })
            Office                             = $(if ($ad) { "$($ad.Office)" } else { '' })
            Phone                              = $(if ($ad) { "$($ad.telephoneNumber)" } else { '' })
            Manager                            = $mgr
            MemberOf                           = ConvertTo-CsvList $memberOf
            MemberOfCount                      = $memberOf.Count
            HiddenFromAddressListsEnabled      = "$($rm.HiddenFromAddressListsEnabled)"
            RequireSenderAuthenticationEnabled = "$($rm.RequireSenderAuthenticationEnabled)"
            AcceptMessagesOnlyFrom             = ConvertTo-CsvList (Get-RecipientSmtp -Refs $rm.AcceptMessagesOnlyFrom -Resolver $resolver)
            AcceptMessagesOnlyFromDLMembers    = ConvertTo-CsvList (Get-RecipientSmtp -Refs $rm.AcceptMessagesOnlyFromDLMembers -Resolver $resolver)
            RejectMessagesFrom                 = ConvertTo-CsvList (Get-RecipientSmtp -Refs $rm.RejectMessagesFrom -Resolver $resolver)
            RejectMessagesFromDLMembers        = ConvertTo-CsvList (Get-RecipientSmtp -Refs $rm.RejectMessagesFromDLMembers -Resolver $resolver)
            ModerationEnabled                  = "$($rm.ModerationEnabled)"
            ModeratedBy                        = ConvertTo-CsvList (Get-RecipientSmtp -Refs $rm.ModeratedBy -Resolver $resolver)
            GrantSendOnBehalfTo                = ConvertTo-CsvList (Get-RecipientSmtp -Refs $rm.GrantSendOnBehalfTo -Resolver $resolver)
            SendAs                             = ConvertTo-CsvList $sendAs
            MailTip                            = "$($rm.MailTip)"
            EmailAddressPolicyEnabled          = "$($rm.EmailAddressPolicyEnabled)"
            WhenCreated                        = "$($rm.WhenCreated)"
        }
        for ($n = 1; $n -le 15; $n++) { $row["CustomAttribute$n"] = "$($rm."CustomAttribute$n")" }
        for ($n = 1; $n -le 5; $n++)  { $row["ExtensionCustomAttribute$n"] = ConvertTo-CsvList $rm."ExtensionCustomAttribute$n" }
        $row['ExportedUtc'] = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
        $row['Error'] = ''
        $rows += [PSCustomObject]$row
        Write-Host "  ok  (member of $($memberOf.Count))" -ForegroundColor Green
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
$rows | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8
Write-ExportDone -Path $outFile -Rows $rows.Count
# Group memberships, one row each, always written (empty = header only) so
# "no file" can never be mistaken for "no groups".
$grpFile = Get-SmbBatchFile -Folder $basePath -BatchId $batchId -Stage 'Groups'
if ($groupRows.Count -eq 0) { '"Batch","PrimarySmtpAddress","Name","GroupName","GroupMail","GroupCategory","GroupScope","GroupDn","Resolved"' | Set-Content -LiteralPath $grpFile -Encoding UTF8 }
else { $groupRows | Export-Csv -LiteralPath $grpFile -NoTypeInformation -Encoding UTF8 }
Write-ExportDone -Path $grpFile -Rows $groupRows.Count
$bad = @($rows | Where-Object { $_.Error }).Count
if ($bad) { Write-Host "$bad mailbox(es) failed -- see the Error column." -ForegroundColor Yellow }

# ---- tag check: Step 1 should have run ----
$untagged = @($rows | Where-Object { -not $_.Error -and "$($_.TaggedOnPrem)" -ne 'True' })
if ($untagged.Count) {
    Write-Host ''
    Write-Host "NOTE: $($untagged.Count) mailbox(es) do not carry the Step 1 display-name tag on-premises -- the admin-center search in Step 6 will not find them by tag:" -ForegroundColor Yellow
    $untagged | ForEach-Object { Write-Host "  $($_.PrimarySmtpAddress)  ($($_.TaggedDisplayName))" -ForegroundColor Yellow }
}

# MemberOf is informational here: a mailbox that stops existing on-premises
# drops out of every on-premises group, and those memberships are not
# recreated (lists are not part of this project).
$inGroups = @($rows | Where-Object { -not $_.Error -and [int]$_.MemberOfCount -gt 0 })
if ($inGroups.Count) {
    Write-Host ''
    Write-Host "NOTE: $($inGroups.Count) mailbox(es) are members of on-premises groups; those memberships end when the account is deleted in Step 4 and are NOT recreated:" -ForegroundColor Yellow
    $inGroups | ForEach-Object { Write-Host "  $($_.PrimarySmtpAddress)  <-  $($_.MemberOf)" -ForegroundColor Yellow }
}
Write-Host "NEXT: Step 3 (cloud export)." -ForegroundColor Cyan
Write-CompletedLog -Action 'SmbMigrationExportOnPrem' -Target "batch $batchId" -Result $(if ($bad) { 'Partial' } else { 'Success' }) -Detail "$($rows.Count - $bad) exported, $bad failed -> $([System.IO.Path]::GetFileName($outFile))"
