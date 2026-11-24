#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# STEP DL-2 of the Migration project. The CLOUD copy of Step DL-1: exports every
# Exchange Online attribute of the distribution lists named in DL-Input.csv.
#
# READ ONLY. Nothing is changed.
#
# WHY BOTH SIDES. A synced list has two LegacyExchangeDNs -- the on-premises
# one (Step DL-1) and the cloud one (this) -- and both must come back as X500
# addresses on the recreated list, or replies to old mail bounce. The cloud
# side also holds things the on-premises object never had: Send As granted in
# the cloud, members that are cloud-only recipients, the @onmicrosoft.com
# proxy. Step 4 stamps the union of the two exports.
#
# Same columns as Step DL-1 plus the cloud-only ones, so the two files can be
# laid side by side in Excel.
#
# Output lands next to this script:
#     Migration-DL-CLOUD_<date>_<time>.csv
#
# Needs Exchange Online connected in this window. Run Step 1 next.
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

$go = Show-StepInstructions -Title 'DL-2 - CLOUD Export DL' -ScriptPath $PSCommandPath -Text @"
STEP DL-2 of 6 -- export the CLOUD copies of the same distribution lists.

* Reads DL-Input.csv (same file as Step DL-1).
* Nothing is changed. This only reads.
* Output: Migration-DL-CLOUD_<date>_<time>.csv next to this script.

Exchange Online must be connected in this window. Click OK to review the
input file and run.
"@
if (-not $go) { return }

function exo { param([string]$n, [hashtable]$p) Invoke-EXOCommand -Name $n -Parameters $p }

$names = Read-InputList -Path (Join-Path $basePath 'DL-Input.csv') -What 'list'
$outFile = Join-Path $basePath "Migration-DL-CLOUD_$(Get-MigrationTimestamp).csv"

$resolver = { param($id) Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{ Identity = $id; ErrorAction = 'Stop' } }

$rows = @()
$i = 0
foreach ($name in $names) {
    $i++
    Write-Host "[$i/$($names.Count)] $name" -NoNewline
    try {
        $dg = exo 'Get-DistributionGroup' @{ Identity = $name; ErrorAction = 'Stop' }
        $grp = $null
        try { $grp = exo 'Get-Group' @{ Identity = $dg.Identity; ErrorAction = 'Stop' } } catch { }

        $members = @(exo 'Get-DistributionGroupMember' @{ Identity = $dg.Identity; ResultSize = 'Unlimited'; ErrorAction = 'Stop' } |
                     ForEach-Object { if ("$($_.PrimarySmtpAddress)") { "$($_.PrimarySmtpAddress)" } else { "$($_.Name)" } })

        # SEND AS IN THE CLOUD IS Get-RecipientPermission, NOT Get-ADPermission.
        # Learn: Get-RecipientPermission -Identity; Trustee is an address or
        # name; AccessRights SendAs. SELF is on every object and is noise.
        $memberOf = @()
        try {
            $dnQ = "$($dg.DistinguishedName)" -replace "'", "''"
            $memberOf = @(exo 'Get-DistributionGroup' @{ Filter = "Members -eq '$dnQ'"; ResultSize = 'Unlimited'; ErrorAction = 'Stop' } |
                          ForEach-Object { "$($_.PrimarySmtpAddress)" })
        } catch { Write-Host "  (MemberOf not read: $($_.Exception.Message))" -ForegroundColor Yellow -NoNewline }

        $sendAs = @()
        try {
            $sendAs = @(exo 'Get-RecipientPermission' @{ Identity = $dg.Identity; ErrorAction = 'Stop' } |
                        Where-Object { $_.AccessRights -contains 'SendAs' -and "$($_.Trustee)" -notmatch 'NT AUTHORITY\\SELF|^S-1-5' } |
                        ForEach-Object { "$($_.Trustee)" })
        } catch { Write-Host "  (Send As not read: $($_.Exception.Message))" -ForegroundColor Yellow -NoNewline }

        $row = [ordered]@{
            Name                                 = "$($dg.Name)"
            DisplayName                          = "$($dg.DisplayName)"
            Alias                                = "$($dg.Alias)"
            PrimarySmtpAddress                   = "$($dg.PrimarySmtpAddress)"
            EmailAddresses                       = ConvertTo-CsvList ($dg.EmailAddresses | ForEach-Object { "$_" })
            LegacyExchangeDN                     = "$($dg.LegacyExchangeDN)"
            GroupType                            = "$($dg.GroupType)"
            RecipientTypeDetails                 = "$($dg.RecipientTypeDetails)"
            ExternalDirectoryObjectId            = "$($dg.ExternalDirectoryObjectId)"
            Guid                                 = "$($dg.Guid)"
            IsDirSynced                          = "$($dg.IsDirSynced)"
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

$header = @($rows | Where-Object { -not $_.Error } | Select-Object -First 1)
if ($header.Count -gt 0) {
    $cols = @($header[0].PSObject.Properties.Name)
    $rows = @($rows | ForEach-Object { $r = $_; $o = [ordered]@{}; foreach ($c in $cols) { $o[$c] = $(if ($r.PSObject.Properties[$c]) { $r.$c } else { '' }) }; [PSCustomObject]$o })
}
$rows | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8
Write-ExportDone -Path $outFile -Rows $rows.Count
$bad = @($rows | Where-Object { $_.Error }).Count
if ($bad) { Write-Host "$bad list(s) failed -- see the Error column." -ForegroundColor Yellow }
Write-CompletedLog -Action 'MigrationExportDL' -Target 'cloud' -Result $(if ($bad) { 'Partial' } else { 'Success' }) -Detail "$($rows.Count - $bad) exported, $bad failed -> $([System.IO.Path]::GetFileName($outFile))"
