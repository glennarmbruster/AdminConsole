#Requires -Version 5.1
# TOOLKIT-REQUIRES: OwnProcess  -- Get-DistributionGroup, Get-Group,
#                                 Get-DistributionGroupMember and
#                                 Get-Recipient exist in BOTH Exchanges. This
#                                 script means the on-premises copies and must
#                                 not share a process with a cloud-connected
#                                 console.
# TOOLKIT-HELP-BEGIN
# STEP DL-1 of the Migration project (DLs and shared mailboxes to cloud-only).
# Exports every ON-PREMISES attribute of the distribution lists named in
# DL-Input.csv -- everything Step 5 needs to restamp the recreated cloud list.
#
# READ ONLY. Nothing is changed.
#
# Captured per list: names, alias, primary address, every proxy address,
# LegacyExchangeDN (becomes an X500 address later -- without it, replies to
# old mail bounce), owners, members, join/depart rules, sender restrictions,
# moderation, hidden-from-GAL, send on behalf, Send As, custom attributes 1-15,
# extension attributes 1-5, MailTip, report/OOF flags, group type, OU.
#
# Owners, members and every recipient reference are written as PRIMARY SMTP
# ADDRESSES, not directory names: an address is the only form that still
# resolves once the object is cloud-only.
#
# DL-Input.csv has one column, Name. Alias, address or display name all
# work. The file opens for review before anything runs.
#
# Output lands next to this script:
#     Migration-DL-ONPREM_<date>_<time>.csv
#
# Run Step DL-2 (the cloud copy of this export) next.
# TOOLKIT-HELP-END

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Common.ps1 lives at the TOOLKIT ROOT and there is exactly one of it. Walk up
# to it rather than counting folders (see Projects\README.md).
$toolkitRoot = $PSScriptRoot
while ($toolkitRoot -and -not (Test-Path -LiteralPath (Join-Path $toolkitRoot 'Common.ps1'))) {
    $toolkitRoot = Split-Path -Parent $toolkitRoot
}
if (-not $toolkitRoot) { throw "Common.ps1 was not found in any folder above '$PSScriptRoot'." }
. (Join-Path $toolkitRoot 'Common.ps1')
. (Join-Path $PSScriptRoot '_Shared.ps1')

$basePath = $PSScriptRoot

$go = Show-StepInstructions -Title 'DL-1 - ONPREM Export DL' -ScriptPath $PSCommandPath -Text @"
STEP DL-1 of 7 -- export the ON-PREMISES distribution lists.

* DL-Input.csv: one column, Name. One list per row.
* Nothing is changed. This only reads.
* Output: Migration-DL-ONPREM_<date>_<time>.csv next to this script.

This window must be connected to ON-PREMISES Exchange (it was opened that way
by the console). Click OK to review the input file and run.
"@
if (-not $go) { return }

$names = Read-InputList -Path (Join-Path $basePath 'DL-Input.csv') -What 'list'
$outFile = Join-Path $basePath "Migration-DL-ONPREM_$(Get-MigrationTimestamp).csv"

$resolver = { param($id) Get-Recipient -Identity $id -ErrorAction Stop }

$rows = @()
$i = 0
foreach ($name in $names) {
    $i++
    Write-Host "[$i/$($names.Count)] $name" -NoNewline
    try {
        $dg = Get-DistributionGroup -Identity $name -ErrorAction Stop
        # Get-Group carries Notes; Get-DistributionGroup does not expose it.
        $grp = $null
        try { $grp = Get-Group -Identity $dg.Identity -ErrorAction Stop } catch { }

        # -ResultSize Unlimited: the 5,000 cap is Get-ADGroupMember's, but the
        # Exchange default page is 1,000 and a company-wide list is bigger.
        $members = @(Get-DistributionGroupMember -Identity $dg.Identity -ResultSize Unlimited -ErrorAction Stop |
                     ForEach-Object { if ("$($_.PrimarySmtpAddress)") { "$($_.PrimarySmtpAddress)" } else { "$($_.Name)" } })

        # Send As on-premises is an AD ACE (Get-ADPermission), reported as
        # DOMAIN\sam. Resolved to an address here; SELF and inherited entries
        # are noise.
        # MEMBER OF: the lists THIS list is nested in. A list leaving sync scope
        # is dropped from every synced parent, and the cloud copies of those
        # parents are read-only -- the same trap Step 1 flags for mailboxes.
        $memberOf = @()
        try {
            $dnQ = "$($dg.DistinguishedName)" -replace "'", "''"
            $memberOf = @(Get-DistributionGroup -Filter "Members -eq '$dnQ'" -ResultSize Unlimited -ErrorAction Stop |
                          ForEach-Object { "$($_.PrimarySmtpAddress)" })
        } catch { Write-Host "  (MemberOf not read: $($_.Exception.Message))" -ForegroundColor Yellow -NoNewline }

        $sendAs = @()
        try {
            $aces = @(Get-ADPermission -Identity $dg.DistinguishedName -ErrorAction Stop |
                      Where-Object { -not $_.IsInherited -and $_.ExtendedRights -match 'Send-As' -and "$($_.User)" -notmatch 'NT AUTHORITY\\SELF|^S-1-5' })
            foreach ($a in $aces) {
                $sam = ("$($a.User)" -split '\\')[-1]
                $r = $null
                try { $r = Get-Recipient -Identity $sam -ErrorAction Stop } catch { }
                $sendAs += $(if ($r) { "$($r.PrimarySmtpAddress)" } else { "$($a.User)" })
            }
        } catch { Write-Host "  (Send As not read: $($_.Exception.Message))" -ForegroundColor Yellow -NoNewline }

        $row = [ordered]@{
            Name                                 = "$($dg.Name)"
            DisplayName                          = "$($dg.DisplayName)"
            Alias                                = "$($dg.Alias)"
            PrimarySmtpAddress                   = "$($dg.PrimarySmtpAddress)"
            EmailAddresses                       = ConvertTo-CsvList ($dg.EmailAddresses | ForEach-Object { $_.ProxyAddressString })
            LegacyExchangeDN                     = "$($dg.LegacyExchangeDN)"
            GroupType                            = "$($dg.GroupType)"
            RecipientTypeDetails                 = "$($dg.RecipientTypeDetails)"
            SamAccountName                       = "$($dg.SamAccountName)"
            Guid                                 = "$($dg.Guid)"
            DistinguishedName                    = "$($dg.DistinguishedName)"
            OrganizationalUnit                   = "$($dg.OrganizationalUnit)"
            ManagedBy                            = ConvertTo-CsvList (Get-RecipientSmtp -Refs $dg.ManagedBy -Resolver $resolver)
            Members                              = ConvertTo-CsvList $members
            MemberCount                          = $members.Count
            MemberOf                             = ConvertTo-CsvList $memberOf
            MemberOfCount                        = $memberOf.Count
            MemberJoinRestriction                = "$($dg.MemberJoinRestriction)"
            MemberDepartRestriction              = "$($dg.MemberDepartRestriction)"
            RequireSenderAuthenticationEnabled   = "$($dg.RequireSenderAuthenticationEnabled)"
            AcceptMessagesOnlyFrom               = ConvertTo-CsvList (Get-RecipientSmtp -Refs $dg.AcceptMessagesOnlyFrom -Resolver $resolver)
            AcceptMessagesOnlyFromDLMembers      = ConvertTo-CsvList (Get-RecipientSmtp -Refs $dg.AcceptMessagesOnlyFromDLMembers -Resolver $resolver)
            RejectMessagesFrom                   = ConvertTo-CsvList (Get-RecipientSmtp -Refs $dg.RejectMessagesFrom -Resolver $resolver)
            RejectMessagesFromDLMembers          = ConvertTo-CsvList (Get-RecipientSmtp -Refs $dg.RejectMessagesFromDLMembers -Resolver $resolver)
            ModerationEnabled                    = "$($dg.ModerationEnabled)"
            ModeratedBy                          = ConvertTo-CsvList (Get-RecipientSmtp -Refs $dg.ModeratedBy -Resolver $resolver)
            SendModerationNotifications          = "$($dg.SendModerationNotifications)"
            BypassModerationFromSendersOrMembers = ConvertTo-CsvList (Get-RecipientSmtp -Refs $dg.BypassModerationFromSendersOrMembers -Resolver $resolver)
            BypassNestedModerationEnabled        = "$($dg.BypassNestedModerationEnabled)"
            HiddenFromAddressListsEnabled        = "$($dg.HiddenFromAddressListsEnabled)"
            GrantSendOnBehalfTo                  = ConvertTo-CsvList (Get-RecipientSmtp -Refs $dg.GrantSendOnBehalfTo -Resolver $resolver)
            SendAs                               = ConvertTo-CsvList $sendAs
            ReportToManagerEnabled               = "$($dg.ReportToManagerEnabled)"
            ReportToOriginatorEnabled            = "$($dg.ReportToOriginatorEnabled)"
            SendOofMessageToOriginatorEnabled    = "$($dg.SendOofMessageToOriginatorEnabled)"
            MailTip                              = "$($dg.MailTip)"
            Notes                                = $(if ($grp) { "$($grp.Notes)" } else { '' })
            EmailAddressPolicyEnabled            = "$($dg.EmailAddressPolicyEnabled)"
            WhenCreated                          = "$($dg.WhenCreated)"
        }
        for ($n = 1; $n -le 15; $n++) { $row["CustomAttribute$n"] = "$($dg."CustomAttribute$n")" }
        for ($n = 1; $n -le 5; $n++)  { $row["ExtensionCustomAttribute$n"] = ConvertTo-CsvList $dg."ExtensionCustomAttribute$n" }
        $row['ExportedUtc'] = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
        $row['Error'] = ''
        $rows += [PSCustomObject]$row
        Write-Host "  ok  ($($members.Count) members, member of $($memberOf.Count))" -ForegroundColor Green
    } catch {
        Write-Host "  FAILED: $($_.Exception.Message)" -ForegroundColor Red
        $rows += [PSCustomObject]@{ Name = $name; Error = "$($_.Exception.Message)" }
    }
}

# Every row must carry every column or Export-Csv drops the extras from the
# later rows; a failed row is padded to the full header.
$header = @($rows | Where-Object { -not $_.Error } | Select-Object -First 1)
if ($header.Count -gt 0) {
    $cols = @($header[0].PSObject.Properties.Name)
    $rows = @($rows | ForEach-Object { $r = $_; $o = [ordered]@{}; foreach ($c in $cols) { $o[$c] = $(if ($r.PSObject.Properties[$c]) { $r.$c } else { '' }) }; [PSCustomObject]$o })
}
$rows | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8
Write-ExportDone -Path $outFile -Rows $rows.Count
$bad = @($rows | Where-Object { $_.Error }).Count
if ($bad) { Write-Host "$bad list(s) failed -- see the Error column. Fix DL-Input.csv and re-run before going on." -ForegroundColor Yellow }

# Parents NOT in DL-Input.csv lose this member at Step 3. Said here, not in Excel.
$migrating = @($names | ForEach-Object { $_.ToLower() }) + @($rows | Where-Object { -not $_.Error } | ForEach-Object { "$($_.PrimarySmtpAddress)".ToLower() })
$orphans = @{}
foreach ($r in ($rows | Where-Object { -not $_.Error })) {
    foreach ($g in (ConvertFrom-CsvList $r.MemberOf)) {
        $key = $g.ToLower(); if ($migrating -contains $key -or $migrating -contains ($key -split '@')[0]) { continue }
        if (-not $orphans.ContainsKey($g)) { $orphans[$g] = @() }
        $orphans[$g] += $r.PrimarySmtpAddress
    }
}
if ($orphans.Count) {
    Write-Host ''
    Write-Host "NESTINGS THAT WILL BE LOST -- parent lists NOT in DL-Input.csv that contain these lists:" -ForegroundColor Yellow
    foreach ($g in ($orphans.Keys | Sort-Object)) { Write-Host "  $g  <-  $($orphans[$g] -join ', ')" -ForegroundColor Yellow }
    Write-Host "Add them to DL-Input.csv and re-run Steps DL-1 and DL-2, or accept the loss, BEFORE Step 3." -ForegroundColor Yellow
}
Write-CompletedLog -Action 'MigrationExportDL' -Target 'on-premises' -Result $(if ($bad) { 'Partial' } else { 'Success' }) -Detail "$($rows.Count - $bad) exported, $bad failed -> $([System.IO.Path]::GetFileName($outFile))"
