#Requires -Version 5.1
# TOOLKIT-REQUIRES: OwnProcess  -- Get-Recipient and Get-User exist in BOTH
#                                 Exchanges; Get-RemoteMailbox is on-premises
#                                 only. This script means the on-premises
#                                 copies and must not share a process with a
#                                 cloud-connected console.
# TOOLKIT-HELP-BEGIN
# STEP 3 of the Migration project. Exports every ON-PREMISES attribute of the
# shared mailboxes named in SMB-Input.csv.
#
# READ ONLY. Nothing is changed.
#
# In a hybrid directory the on-premises object is a REMOTE MAILBOX: an AD
# account with mail attributes and a pointer to the cloud. The mailbox
# itself, its permissions and its contents live in the cloud (Step 4). What
# only exists here: the AD account (sam, UPN, OU, enabled, description,
# department, manager), the on-premises LegacyExchangeDN, Send As granted
# on-premises, and -- the one that bites -- MEMBER OF: every on-premises
# list this mailbox belongs to.
#
# MEMBER OF MATTERS. When the account leaves the sync scope in Step 5, every
# synced list drops it as a member on the next sync, and the cloud copies of
# those lists are read-only. Lists in DL-Input.csv are recreated in the
# cloud with their members (Steps 6-7), so those come back. Any list in the
# MemberOf column that is NOT in DL-Input.csv is a decision: migrate it too,
# or accept the loss. The column exists so that decision is made before
# Step 5, not discovered after.
#
# SMB-Input.csv has one column, Name. Alias, address or display name work.
#
# Output lands next to this script:
#     Migration-SMB-ONPREM_<date>_<time>.csv
#
# Run Step 4 (the cloud copy) next.
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

$go = Show-StepInstructions -Title '3 - ONPREM Export SMB' -ScriptPath $PSCommandPath -Text @"
STEP 3 of 7 -- export the ON-PREMISES shared mailbox accounts.

* SMB-Input.csv: one column, Name. One mailbox per row.
* Nothing is changed. This only reads.
* Output: Migration-SMB-ONPREM_<date>_<time>.csv next to this script.
* CHECK THE MemberOf COLUMN afterwards: lists there that are not in
  DL-Input.csv will lose this member when it leaves sync scope.

This window must be connected to ON-PREMISES Exchange. Click OK to review
the input file and run.
"@
if (-not $go) { return }

Import-Module ActiveDirectory -ErrorAction SilentlyContinue

$names = Read-InputList -Path (Join-Path $basePath 'SMB-Input.csv') -What 'mailbox'
$outFile = Join-Path $basePath "Migration-SMB-ONPREM_$(Get-MigrationTimestamp).csv"

$resolver = { param($id) Get-Recipient -Identity $id -ErrorAction Stop }

$rows = @()
$i = 0
foreach ($name in $names) {
    $i++
    Write-Host "[$i/$($names.Count)] $name" -NoNewline
    try {
        # Get-RemoteMailbox is on-premises ONLY. Its absence from the cloud is
        # the point: this is the object that decides what the cloud copy looks
        # like -- until Step 5 breaks that link on purpose.
        $rm = Get-RemoteMailbox -Identity $name -ErrorAction Stop
        $u  = $null
        try { $u = Get-User -Identity $rm.Identity -ErrorAction Stop } catch { }

        $ad = $null; $memberOf = @()
        try {
            $ad = Get-ADUser -Identity $rm.DistinguishedName -Properties memberOf, Enabled, Description, Department, Title, Manager, Company, Office, telephoneNumber, whenCreated, whenChanged, msExchRecipientTypeDetails, msExchRemoteRecipientType -ErrorAction Stop
            foreach ($gdn in @($ad.memberOf)) {
                # Address if the group has one, name if not -- same rule as
                # every other reference in these exports.
                $g = $null
                try { $g = Get-ADGroup -Identity $gdn -Properties mail -ErrorAction Stop } catch { }
                $memberOf += $(if ($g -and "$($g.mail)") { "$($g.mail)" } elseif ($g) { "$($g.Name)" } else { "$gdn" })
            }
        } catch { Write-Host "  (AD read failed: $($_.Exception.Message))" -ForegroundColor Yellow -NoNewline }

        $sendAs = @()
        try {
            $aces = @(Get-ADPermission -Identity $rm.DistinguishedName -ErrorAction Stop |
                      Where-Object { -not $_.IsInherited -and $_.ExtendedRights -match 'Send-As' -and "$($_.User)" -notmatch 'NT AUTHORITY\\SELF|^S-1-5' })
            foreach ($a in $aces) {
                $sam = ("$($a.User)" -split '\\')[-1]
                $r = $null
                try { $r = Get-Recipient -Identity $sam -ErrorAction Stop } catch { }
                $sendAs += $(if ($r) { "$($r.PrimarySmtpAddress)" } else { "$($a.User)" })
            }
        } catch { Write-Host "  (Send As not read: $($_.Exception.Message))" -ForegroundColor Yellow -NoNewline }

        $mgr = ''
        if ($ad -and "$($ad.Manager)") { try { $mgr = "$((Get-ADUser -Identity $ad.Manager -Properties mail -ErrorAction Stop).mail)" } catch { $mgr = "$($ad.Manager)" } }

        $row = [ordered]@{
            Name                               = "$($rm.Name)"
            DisplayName                        = "$($rm.DisplayName)"
            Alias                              = "$($rm.Alias)"
            SamAccountName                     = "$($rm.SamAccountName)"
            UserPrincipalName                  = "$($rm.UserPrincipalName)"
            PrimarySmtpAddress                 = "$($rm.PrimarySmtpAddress)"
            EmailAddresses                     = ConvertTo-CsvList ($rm.EmailAddresses | ForEach-Object { $_.ProxyAddressString })
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
$bad = @($rows | Where-Object { $_.Error }).Count
if ($bad) { Write-Host "$bad mailbox(es) failed -- see the Error column." -ForegroundColor Yellow }

# The MemberOf warning, computed here so nobody has to eyeball it in Excel.
$dlInput = Join-Path $basePath 'DL-Input.csv'
$migrating = @()
if (Test-Path -LiteralPath $dlInput) { $migrating = @(Import-Csv -LiteralPath $dlInput | ForEach-Object { "$($_.Name)".Trim().ToLower() } | Where-Object { $_ }) }
$orphans = @{}
foreach ($r in ($rows | Where-Object { -not $_.Error })) {
    foreach ($g in (ConvertFrom-CsvList $r.MemberOf)) {
        $key = $g.ToLower(); $local = ($key -split '@')[0]
        if ($migrating -contains $key -or $migrating -contains $local) { continue }
        if (-not $orphans.ContainsKey($g)) { $orphans[$g] = @() }
        $orphans[$g] += $r.PrimarySmtpAddress
    }
}
if ($orphans.Count) {
    Write-Host ''
    Write-Host "MEMBERSHIPS THAT WILL BE LOST -- groups NOT in DL-Input.csv that contain these mailboxes:" -ForegroundColor Yellow
    foreach ($g in ($orphans.Keys | Sort-Object)) { Write-Host "  $g  <-  $($orphans[$g] -join ', ')" -ForegroundColor Yellow }
    Write-Host "Add them to DL-Input.csv and re-run Steps 1-2, or accept the loss, BEFORE Step 5." -ForegroundColor Yellow
}
Write-CompletedLog -Action 'MigrationExportSMB' -Target 'on-premises' -Result $(if ($bad) { 'Partial' } else { 'Success' }) -Detail "$($rows.Count - $bad) exported, $bad failed -> $([System.IO.Path]::GetFileName($outFile))"
