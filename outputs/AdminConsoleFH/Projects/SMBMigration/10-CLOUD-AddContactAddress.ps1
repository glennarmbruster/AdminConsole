#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# ONE-TIME follow-up to the SMBMigration project (2026-10-13). Every batch
# left an on-premises mail contact whose external address is
# <alias>@firsthorizon.onmicrosoft.com. This adds that same address as a
# proxy (smtp:) on the matching cloud mailbox, so mail the contact forwards
# into the tenant is accepted by the mailbox.
#
# Source: every SMBMigration-<batch>-Contacts.csv in this folder (Step 5),
# ExternalEmailAddress column. A batch with no Contacts file falls back to
# its Step 2 export (Alias + the onmicrosoft domain). All batches at once;
# a mailbox is processed once even if it appears in more than one file.
#
# A mailbox that already has the address, is not in Exchange Online, or is
# still directory-synced is reported and skipped. Nothing is removed and
# the primary address is never changed.
#
# Output: SMBMigration-AddContactAddress-<timestamp>.csv. Needs Exchange Online.
# TOOLKIT-HELP-END
# TOOLKIT-HELP-RUNS: Exchange Online. Adds one smtp: proxy address to each migrated mailbox.

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

# One-time script for one tenant: the domain is written in on purpose.
$onmsDomain = 'firsthorizon.onmicrosoft.com'

function exo { param([string]$n, [hashtable]$p) Invoke-EXOCommand -Name $n -Parameters $p }

# ---- gather the address for every migrated mailbox, across all batches ----
$wanted = [ordered]@{}        # PrimarySmtpAddress (lower) -> @{ Address; Add; Batch; Source }
$sources = @()
$batches = @(Get-SmbBatchList -Folder $basePath | Sort-Object)
foreach ($b in $batches) {
    $cFile = Get-SmbBatchFile -Folder $basePath -BatchId $b -Stage 'Contacts'
    $oFile = Get-SmbBatchFile -Folder $basePath -BatchId $b -Stage 'ONPREM'
    if (Test-Path -LiteralPath $cFile) {
        $n = 0
        foreach ($r in @(Import-Csv -LiteralPath $cFile)) {
            $addr = "$($r.PrimarySmtpAddress)".Trim(); $ext = "$($r.ExternalEmailAddress)".Trim() -replace '^smtp:', ''
            if (-not $addr -or -not $ext) { continue }
            $k = $addr.ToLowerInvariant()
            if (-not $wanted.Contains($k)) { $wanted[$k] = @{ Address = $addr; Add = $ext; Batch = $b; Source = 'Contacts' }; $n++ }
        }
        $sources += "$b -- Contacts file, $n mailbox(es)"
    } elseif (Test-Path -LiteralPath $oFile) {
        $n = 0
        foreach ($r in @(Import-Csv -LiteralPath $oFile | Where-Object { -not "$($_.Error)" })) {
            $addr = "$($r.PrimarySmtpAddress)".Trim(); $alias = "$($r.Alias)".Trim()
            if (-not $addr -or -not $alias) { continue }
            $k = $addr.ToLowerInvariant()
            if (-not $wanted.Contains($k)) { $wanted[$k] = @{ Address = $addr; Add = "$alias@$onmsDomain"; Batch = $b; Source = 'ONPREM export (no Contacts file)' }; $n++ }
        }
        $sources += "$b -- NO Contacts file; built from the Step 2 export, $n mailbox(es)"
    } else {
        $sources += "$b -- no Contacts or ONPREM file; skipped"
    }
}
if ($wanted.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('No SMBMigration-<batch>-Contacts.csv or -ONPREM.csv files were found in this folder. Nothing to do.', '10 - CLOUD Add Contact Address', 'OK', 'Error') | Out-Null; return }

$go = Show-StepInstructions -Title '10 - CLOUD Add Contact Address' -ScriptPath $PSCommandPath -Text @"
ONE-TIME -- add the on-premises contact's address to each migrated cloud mailbox.

* Address added: <alias>@$onmsDomain, as a secondary (smtp:) proxy
* Mailboxes: $($wanted.Count) across $($batches.Count) batch file(s)
$(($sources | ForEach-Object { "    $_" }) -join "`r`n")

Exchange Online must be connected in this window. Click OK to check each mailbox.
"@
if (-not $go) { return }

# ---- plan ----
$plan = @(); $skip = @(); $i = 0
foreach ($k in $wanted.Keys) {
    $w = $wanted[$k]; $i++
    Write-Progress -Activity 'Checking mailboxes' -Status "$i of $($wanted.Count): $($w.Address)" -PercentComplete ([int](100 * $i / $wanted.Count))
    $mb = $null
    try { $mb = exo 'Get-Mailbox' @{ Identity = $w.Address; ErrorAction = 'Stop' } } catch { }
    if (-not $mb) { $skip += "$($w.Address) -- not found in Exchange Online"; continue }
    if ("$($mb.IsDirSynced)" -eq 'True') { $skip += "$($w.Address) -- still directory-synced; the address must be added on-premises"; continue }
    $have = @($mb.EmailAddresses | ForEach-Object { "$_" })
    if ($have | Where-Object { $_ -ieq "smtp:$($w.Add)" }) { $skip += "$($w.Address) -- already has $($w.Add)"; continue }
    $plan += [PSCustomObject]@{ Address = "$($mb.PrimarySmtpAddress)"; Add = $w.Add; Batch = $w.Batch; Source = $w.Source }
}
Write-Progress -Activity 'Checking mailboxes' -Completed
if ($skip.Count) { Write-Host 'Not changed:' -ForegroundColor Yellow; $skip | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow } }
if ($plan.Count -eq 0) { Write-Host "Nothing to add -- $($skip.Count) skipped." -ForegroundColor Yellow }
else {
    $a = [System.Windows.Forms.MessageBox]::Show(
        ("Add the contact address to $($plan.Count) mailbox(es)?  ($($skip.Count) skipped -- see the console)" + "`r`n`r`n  " +
         (@($plan | Select-Object -First 15 | ForEach-Object { "$($_.Address)  +  $($_.Add)" }) -join "`r`n  ") +
         $(if ($plan.Count -gt 15) { "`r`n  ... and $($plan.Count - 15) more" } else { '' })),
        'Confirm', 'OKCancel', 'Question')
    if ($a -ne [System.Windows.Forms.DialogResult]::OK) { Write-Host 'Cancelled. Nothing was changed.' -ForegroundColor Yellow; return }
}

# ---- apply ----
$rows = @(); $ok = 0; $failed = 0; $i = 0
foreach ($p in $plan) {
    $res = ''; $i++
    Write-Progress -Activity 'Adding addresses' -Status "$i of $($plan.Count): $($p.Address)" -PercentComplete ([int](100 * $i / $plan.Count))
    try {
        # Learn (Set-Mailbox, Exchange Online): -EmailAddresses @{Add="smtp:..."} adds without touching the rest.
        exo 'Set-Mailbox' @{ Identity = $p.Address; EmailAddresses = @{ Add = "smtp:$($p.Add)" }; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
        $res = 'Added'; $ok++
        Write-Host "  added    $($p.Address)  +  $($p.Add)" -ForegroundColor Green
        Write-CompletedLog -Action 'SmbMigrationAddContactAddress' -Target (Get-LogTarget $p.Address) -Result 'Success' -Environment 'Cloud' -Detail "+ smtp:$($p.Add) (batch $($p.Batch))"
    } catch {
        $res = "FAILED: $($_.Exception.Message)"; $failed++
        Write-Host "  FAILED   $($p.Address)  $($_.Exception.Message)" -ForegroundColor Red
        Write-FailedLog -Action 'SmbMigrationAddContactAddress' -Target (Get-LogTarget $p.Address) -ErrorRecord $_ -Environment 'Cloud'
    }
    $rows += [PSCustomObject]@{ Batch = $p.Batch; PrimarySmtpAddress = $p.Address; AddressAdded = $p.Add; Source = $p.Source; Result = $res; WhenUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') }
}
Write-Progress -Activity 'Adding addresses' -Completed
foreach ($sk in $skip) {
    $addr = ($sk -split ' -- ')[0]; $w = $wanted[$addr.ToLowerInvariant()]
    $rows += [PSCustomObject]@{ Batch = "$($w.Batch)"; PrimarySmtpAddress = $addr; AddressAdded = ''; Source = "$($w.Source)"; Result = "SKIPPED: $(($sk -split ' -- ', 2)[1])"; WhenUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') }
}
$outFile = Join-Path $basePath ($script:BatchPrefix + 'AddContactAddress-' + (Get-MigrationTimestamp) + '.csv')
$rows | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8
Write-ExportDone -Path $outFile -Rows $rows.Count
Write-Host "$ok added, $failed failed, $($skip.Count) skipped." -ForegroundColor $(if ($failed) { 'Yellow' } else { 'Cyan' })
Write-CompletedLog -Action 'SmbMigrationAddContactAddress' -Target 'all batches' -Result $(if ($failed -and $ok) { 'Partial' } elseif ($failed) { 'Failed' } else { 'Success' }) -Detail "$ok added, $failed failed, $($skip.Count) skipped -> $([System.IO.Path]::GetFileName($outFile))"
