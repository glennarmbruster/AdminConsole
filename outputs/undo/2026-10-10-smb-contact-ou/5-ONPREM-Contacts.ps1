#Requires -Version 5.1
# TOOLKIT-REQUIRES: OwnProcess  -- New-/Set-MailContact exist in BOTH
#                                 Exchanges. This script means the on-premises
#                                 copies and must not share a process with a
#                                 cloud-connected console.
# TOOLKIT-HELP-BEGIN
# STEP 5 of the SMBMigration project. Creates an on-premises MAIL CONTACT for
# every mailbox Step 4 deleted, so on-premises systems that send to the old
# address keep delivering -- to the cloud mailbox.
#
# Built from the Step 2 export, per mailbox:
#   Name / Alias           the mailbox's original Name and Alias
#   DisplayName            the original (clean) display name
#   ExternalEmailAddress   <alias>@<tenant>.onmicrosoft.com  -- also the
#                          primary SMTP, the way every other contact here
#                          is built
#   Proxies                every old smtp: address, <alias>@<tenant>.mail.
#                          onmicrosoft.com (the hybrid routing address), and
#                          every X500: the mailbox had plus one for its
#                          on-premises LegacyExchangeDN, so old replies resolve
#   Hidden from the address book; email address policy off so the list sticks.
#
# Where: the Mail contact OU in Tenant Manager (a NON-syncing OU -- these
# contacts must never reach the cloud, where the real mailbox lives).
#
# GATES, soft: Step 2 export exists; each mailbox no longer resolves
# on-premises (Step 4 ran); the contact OU is set. A contact that already
# exists for an address is reported and left alone, so re-running is safe.
#
# Output: SMBMigration-<batch>-Contacts.csv.
# TOOLKIT-HELP-END
# TOOLKIT-HELP-RUNS: On-premises Exchange. Creates mail contacts in the configured contact OU.

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

$batchId = Select-SmbBatch -Folder $basePath -Title '5 - ONPREM Contacts'
if (-not $batchId) { return }
$onFile = Get-SmbBatchFile -Folder $basePath -BatchId $batchId -Stage 'ONPREM'

$cfg = Get-TenantConfig
$onmsDomain = "$($cfg.OnMicrosoftDomain)".Trim()
$routingDomain = ''
try { $routingDomain = "$(Get-TenantRemoteRoutingDomain)".Trim() } catch { }
if (-not $routingDomain -and $onmsDomain) { $routingDomain = ($onmsDomain -replace '\.onmicrosoft\.com$', '.mail.onmicrosoft.com') }
$ou = ''
try { $ou = Get-TenantDirectoryOu -Kind 'MailContact' } catch { $ou = '' }

$go = Show-StepInstructions -Title '5 - ONPREM Contacts' -ScriptPath $PSCommandPath -Text @"
STEP 5 of 9 -- create on-premises mail contacts for batch $batchId.

* From: SMBMigration-$batchId-ONPREM.csv (Step 2)
* Into: $(if ($ou) { $ou } else { '(Mail contact OU NOT SET in Tenant Manager)' })
* Primary / external address: <alias>@$onmsDomain
* Proxies: every old address, <alias>@$routingDomain, and the X500s
* Hidden from the address book. Existing contacts are left alone.

This window must be connected to ON-PREMISES Exchange. Click OK to run.
"@
if (-not $go) { return }

if (-not $onmsDomain) { [System.Windows.Forms.MessageBox]::Show('Tenant.json has no onmicrosoft domain (Tenant Manager). Nothing was created.', '5 - ONPREM Contacts', 'OK', 'Error') | Out-Null; return }
if (-not $ou) { [System.Windows.Forms.MessageBox]::Show('Tenant Manager has no Mail contact OU for this directory. Set it, then run this again. Nothing was created.', '5 - ONPREM Contacts', 'OK', 'Error') | Out-Null; return }
if (-not (Test-Path -LiteralPath $onFile)) { [System.Windows.Forms.MessageBox]::Show("No Step 2 export for this batch ($([System.IO.Path]::GetFileName($onFile))). The contacts are built from it. Nothing was created.", '5 - ONPREM Contacts', 'OK', 'Error') | Out-Null; return }

Import-Module ActiveDirectory -ErrorAction SilentlyContinue
Enter-MigrationOnPremScope

$names = @(Read-SmbBatch -Folder $basePath -BatchId $batchId)
$export = @(Import-Csv -LiteralPath $onFile | Where-Object { -not "$($_.Error)" })

# ---- gate: the mailbox should be gone on-premises ----
$stillThere = @()
foreach ($n in $names) {
    try { $null = Get-RemoteMailbox -Identity $n -ErrorAction Stop; $stillThere += $n } catch { }
}
if ($stillThere.Count) {
    if (-not (Confirm-SmbGate -Title 'Mailboxes still on-premises' -Problem "$($stillThere.Count) mailbox(es) in this batch still exist on-premises -- Step 4 has not deleted them: $($stillThere -join ', '). A contact with the same addresses would collide with them.")) { return }
}

# The contact OU as an on-premises controller sees it: canonical from
# Tenant.json; Exchange accepts canonical for -OrganizationalUnit.
$plan = @(); $skip = @()
foreach ($n in $names) {
    $row = $export | Where-Object { "$($_.PrimarySmtpAddress)" -ieq $n } | Select-Object -First 1
    if (-not $row) { $skip += "$n -- no row in the Step 2 export"; continue }
    $alias = "$($row.Alias)".Trim()
    if (-not $alias) { $skip += "$n -- export row has no Alias"; continue }
    $external = "$alias@$onmsDomain"
    $proxies = New-Object System.Collections.ArrayList
    $seen = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    [void]$seen.Add("SMTP:$external"); [void]$proxies.Add("SMTP:$external")
    foreach ($e in (ConvertFrom-CsvList $row.EmailAddresses)) {
        $v = "$e"
        if ($v -match '^(?i)smtp:') { $v = 'smtp:' + $v.Substring(5) }
        if ($v -match '^(?i)x500:') { $v = 'X500:' + $v.Substring(5) }
        if ($v -match '^(?i)(smtp|x500):' -and $seen.Add($v)) { [void]$proxies.Add($v) }
    }
    $rt = "smtp:$alias@$routingDomain"
    if ($routingDomain -and $seen.Add($rt)) { [void]$proxies.Add($rt) }
    if ("$($row.LegacyExchangeDN)" -and $seen.Add("X500:$($row.LegacyExchangeDN)")) { [void]$proxies.Add("X500:$($row.LegacyExchangeDN)") }
    $plan += [PSCustomObject]@{
        Address = $n; Name = "$($row.Name)"; Alias = $alias; Display = "$($row.DisplayName)"
        External = $external; Proxies = @($proxies.ToArray()); Dn = "$($row.DistinguishedName)"
        Groups = @(ConvertFrom-CsvList $row.MemberOf)
    }
}
if ($skip.Count) { Write-Host 'Not built:' -ForegroundColor Yellow; $skip | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow } }
if ($plan.Count -eq 0) { Write-Host 'Nothing to create.' -ForegroundColor Yellow; return }

Write-Host ''
Write-Host "About to create $($plan.Count) contact(s) in $ou :" -ForegroundColor Cyan
$plan | ForEach-Object { Write-Host "  $($_.Name.PadRight(40)) $($_.External)   [$($_.Proxies.Count) address(es)]" }
$a = [System.Windows.Forms.MessageBox]::Show(
    ("Create $($plan.Count) mail contact(s) in" + "`r`n  $ou" + "`r`n`r`n  " +
     (@($plan | Select-Object -First 15 | ForEach-Object { "$($_.Name)  ->  $($_.External)" }) -join "`r`n  ") +
     $(if ($plan.Count -gt 15) { "`r`n  ... and $($plan.Count - 15) more" } else { '' })),
    'Confirm', 'OKCancel', 'Question')
if ($a -ne [System.Windows.Forms.DialogResult]::OK) { Write-Host 'Cancelled. Nothing was created.' -ForegroundColor Yellow; return }

# THE GROUPS THE MAILBOX WAS IN. The account that held those memberships is
# gone (Step 4). Putting the CONTACT into the same on-premises groups brings
# the mailbox back into them -- for mail that is expanded ON-PREMISES. A
# list that is synced to the cloud expands in the cloud for cloud-delivered
# mail, and its cloud copy never sees a contact from a non-syncing OU, so
# this is a partial fix and is offered, not assumed. Default No.
$withGroups = @($plan | Where-Object { @($_.Groups).Count -gt 0 })
$addToGroups = $false
if ($withGroups.Count) {
    $g = [System.Windows.Forms.MessageBox]::Show(
        ("$($withGroups.Count) of these mailboxes were members of on-premises groups (Step 2's MemberOf)." + "`r`n`r`n" +
         'Add each new CONTACT to the same groups? That restores delivery through those groups for mail expanded on-premises only; a list synced to the cloud expands there for cloud-delivered mail and will NOT include the contact.' + "`r`n`r`n" +
         'Yes = add the contacts to the groups.  No = create the contacts only (the groups are listed in the output file either way).'),
        'Group memberships', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question, [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    $addToGroups = ($g -eq [System.Windows.Forms.DialogResult]::Yes)
}

# A controller in the directory that owns the contact OU (the deleted DN's
# domain is the same directory).
$rows = @(); $ok = 0; $exists = 0; $failed = 0; $groupNotes = @()
foreach ($p in $plan) {
    $res = ''; $srv = Get-MigrationAdServer -DistinguishedName $p.Dn
    $dc = @{}; if ($srv) { $dc['DomainController'] = $srv }
    $have = $null
    try { $have = Get-MailContact -Identity $p.External @dc -ErrorAction Stop } catch { $have = $null }
    if (-not $have) { try { $have = Get-MailContact -Identity $p.Name @dc -ErrorAction Stop } catch { $have = $null } }
    if ($have) {
        $res = 'AlreadyExists'; $exists++
        Write-Host "  exists   $($p.Name)  ($($have.PrimarySmtpAddress))" -ForegroundColor DarkGray
    } else {
        try {
            # Learn (New-MailContact): -Name, -ExternalEmailAddress, -Alias,
            # -DisplayName, -OrganizationalUnit, -DomainController.
            New-MailContact -Name $p.Name -DisplayName $p.Display -Alias $p.Alias -ExternalEmailAddress $p.External -OrganizationalUnit $ou @dc -ErrorAction Stop | Out-Null
            # Learn (Set-MailContact): -EmailAddresses replaces the collection
            # (one upper-case SMTP: = primary), -EmailAddressPolicyEnabled,
            # -HiddenFromAddressListsEnabled, -DomainController. Policy off
            # FIRST, in the same call, or the policy re-stamps the list.
            Set-MailContact -Identity $p.External -EmailAddressPolicyEnabled $false -EmailAddresses $p.Proxies -HiddenFromAddressListsEnabled $true @dc -ErrorAction Stop
            $res = 'Created'; $ok++
            Write-Host "  created  $($p.Name)  ->  $($p.External)  [$($p.Proxies.Count) address(es)]" -ForegroundColor Green
            if ($addToGroups) {
                foreach ($grp in $p.Groups) {
                    try {
                        Add-DistributionGroupMember -Identity $grp -Member $p.External -BypassSecurityGroupManagerCheck -Confirm:$false @dc -ErrorAction Stop
                        $groupNotes += "$($p.Name) -> $grp OK"
                        Write-Host "           added to $grp" -ForegroundColor Green
                    } catch {
                        $groupNotes += "$($p.Name) -> $grp FAILED: $($_.Exception.Message)"
                        Write-Host "           could not add to $grp -- $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }
            }
            Write-CompletedLog -Action 'SmbMigrationContact' -Target (Get-LogTarget $p.Address $p.Name) -Result 'Success' -Environment 'OnPrem' -Detail "New-MailContact $($p.External) in $ou; $($p.Proxies.Count) addresses (batch $batchId)"
        } catch {
            $res = "FAILED: $($_.Exception.Message)"; $failed++
            Write-Host "  FAILED   $($p.Name)  $($_.Exception.Message)" -ForegroundColor Red
            Write-FailedLog -Action 'SmbMigrationContact' -Target (Get-LogTarget $p.Address $p.Name) -ErrorRecord $_ -Environment 'OnPrem'
        }
    }
    $rows += [PSCustomObject]@{ Batch = $batchId; PrimarySmtpAddress = $p.Address; Name = $p.Name; Alias = $p.Alias; DisplayName = $p.Display; ExternalEmailAddress = $p.External; EmailAddresses = ($p.Proxies -join ';'); OrganizationalUnit = $ou; Server = $srv; GroupsOnPrem = ($p.Groups -join ';'); GroupsAdded = $(if ($addToGroups) { 'attempted' } else { 'no' }); Result = $res; WhenUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') }
}
$outFile = Get-SmbBatchFile -Folder $basePath -BatchId $batchId -Stage 'Contacts'
$rows | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8
Write-ExportDone -Path $outFile -Rows $rows.Count
Write-Host ''
Write-Host "$ok created, $exists already existed, $failed failed." -ForegroundColor $(if ($failed) { 'Yellow' } else { 'Cyan' })
if ($groupNotes.Count) { Write-Host 'Group memberships:' -ForegroundColor Cyan; $groupNotes | ForEach-Object { Write-Host "  $_" } }
Write-Host "NEXT: wait for sync (~75 min after Step 4), restore the users in the admin center (Step 6), then Step 7." -ForegroundColor Cyan
Write-CompletedLog -Action 'SmbMigrationContact' -Target "batch $batchId" -Result $(if ($failed -and $ok) { 'Partial' } elseif ($failed) { 'Failed' } else { 'Success' }) -Detail "$ok created, $exists existed, $failed failed -> $([System.IO.Path]::GetFileName($outFile))"
