#Requires -Version 5.1
# TOOLKIT-REQUIRES: OwnProcess  -- Remove-RemoteMailbox is on-premises only
#                                 and must not share a process with a cloud-
#                                 connected console.
# TOOLKIT-HELP-BEGIN
# STEP 4 of the SMBMigration project. THE ONE THAT STARTS THE CLOCK, AND THE
# ONLY IRREVERSIBLE ONE.
#
# Deletes the on-premises AD user object of every mailbox in the batch
# (Remove-RemoteMailbox). On the next directory sync the cloud user is
# deleted too: it goes to the Entra recycle bin and the mailbox becomes
# soft-deleted (30 days) -- which is what Step 6 restores by hand. There is
# no undo for the on-premises object (no AD Recycle Bin); the Step 2 export
# is its only record, and Step 5 builds the mail contact from it.
#
# GATES, all soft -- each warns and lets you go on:
#   * Step 2 and Step 3 exports exist for this batch
#   * every mailbox shows the Step 1 tag in the cloud (Step 3's TaggedInCloud)
#   * every mailbox still resolves on-premises as a remote mailbox
#   * Step 5's contact OU is configured (Tenant Manager, Mail contact OU)
#
# Output: SMBMigration-<batch>-Deleted.csv. Run Step 5 (contacts) right
# after; the wait for sync is for Step 6, not for Step 5.
# TOOLKIT-HELP-END
# TOOLKIT-HELP-RUNS: On-premises Exchange / Active Directory. DELETES the AD user object of each mailbox in the batch.

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

$batchId = Select-SmbBatch -Folder $basePath -Title '4 - ONPREM Delete'
if (-not $batchId) { return }
$onFile = Get-SmbBatchFile -Folder $basePath -BatchId $batchId -Stage 'ONPREM'
$clFile = Get-SmbBatchFile -Folder $basePath -BatchId $batchId -Stage 'CLOUD'

$go = Show-StepInstructions -Title '4 - ONPREM Delete' -ScriptPath $PSCommandPath -Text @"
STEP 4 of 9 -- DELETE the on-premises AD objects of batch $batchId.

* Remove-RemoteMailbox on every mailbox in the batch: the AD user object is
  deleted. The next sync soft-deletes the cloud copy (restored by hand in
  Step 6).
* THERE IS NO UNDO. The Step 2 export is the only record afterwards.
* Gates (soft): exports exist, tag synced, objects resolve, contact OU set.
* Then run Step 5 (contacts) straight away.

This window must be connected to ON-PREMISES Exchange. Click OK to check
the gates; nothing is deleted until you confirm again.
"@
if (-not $go) { return }

Import-Module ActiveDirectory -ErrorAction SilentlyContinue
Enter-MigrationOnPremScope

$names = @(Read-SmbBatch -Folder $basePath -BatchId $batchId)

# ---- gates ----
if (-not (Test-Path -LiteralPath $onFile)) { if (-not (Confirm-SmbGate -Title 'Step 2 export missing' -Problem "No Step 2 export for this batch ($([System.IO.Path]::GetFileName($onFile))). After the delete there will be NO on-premises record of these objects and Step 5 cannot build contacts.")) { return } }
if (-not (Test-Path -LiteralPath $clFile)) { if (-not (Confirm-SmbGate -Title 'Step 3 export missing' -Problem "No Step 3 export for this batch ($([System.IO.Path]::GetFileName($clFile))). Step 7 will have nothing to restamp from on the cloud side.")) { return } }
if (Test-Path -LiteralPath $clFile) {
    $cl = @(Import-Csv -LiteralPath $clFile | Where-Object { -not "$($_.Error)" })
    $notTagged = @($cl | Where-Object { "$($_.TaggedInCloud)" -ne 'True' } | ForEach-Object { $_.PrimarySmtpAddress })
    $missingCl = @($names | Where-Object { $n = $_; -not ($cl | Where-Object { "$($_.PrimarySmtpAddress)" -ieq $n }) })
    if ($notTagged.Count -or $missingCl.Count) {
        $p = ''
        if ($notTagged.Count) { $p += "$($notTagged.Count) mailbox(es) did NOT show the Step 1 tag in the cloud when Step 3 ran: $($notTagged -join ', '). The admin-center search in Step 6 will not find them by tag. " }
        if ($missingCl.Count)  { $p += "$($missingCl.Count) batch address(es) have no row in the Step 3 export: $($missingCl -join ', ')." }
        if (-not (Confirm-SmbGate -Title 'Tag not synced' -Problem $p.Trim())) { return }
    }
}
Write-Host 'Mail contact OU:' -ForegroundColor DarkGray
$ou = Get-SmbContactOu
if (-not $ou) { if (-not (Confirm-SmbGate -Title 'Contact OU not set' -Problem 'No directory row in Tenant Manager has a Mail contact OU (the console window lists each row), so Step 5 cannot create the contacts yet. Fill it in and Save, then start Step 5 fresh.')) { return } }

# ---- resolve every row first ----
$plan = @(); $bad = @()
foreach ($n in $names) {
    try {
        $rm = Get-RemoteMailbox -Identity $n -ErrorAction Stop
        $plan += [PSCustomObject]@{ Address = "$($rm.PrimarySmtpAddress)"; Name = "$($rm.Name)"; Display = "$($rm.DisplayName)"; Dn = "$($rm.DistinguishedName)"; Guid = "$($rm.Guid)"; Sam = "$($rm.SamAccountName)"; Rtd = "$($rm.RecipientTypeDetails)" }
    } catch { $bad += "$n -- $($_.Exception.Message)" }
}
if ($bad.Count) {
    Write-Host "$($bad.Count) address(es) did not resolve as a remote mailbox (already deleted?):" -ForegroundColor Yellow
    $bad | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
    if ($plan.Count -eq 0) { Write-Host 'Nothing left to delete.' -ForegroundColor Yellow; return }
    if (-not (Confirm-SmbGate -Title 'Unresolved rows' -Problem "$($bad.Count) of $($names.Count) rows did not resolve on-premises (listed in the window). Delete the $($plan.Count) that did?")) { return }
}

# A PERSON'S ACCOUNT MUST NOT GET THIS FAR. Step 1 warned; this is the last
# look before an irreversible delete, so it warns again and defaults to No.
$people = @($plan | Where-Object { "$($_.Rtd)" -notmatch '^Remote(Shared|Room|Equipment)Mailbox$' })
if ($people.Count) {
    Write-Host "$($people.Count) object(s) are NOT shared mailboxes:" -ForegroundColor Red
    $people | ForEach-Object { Write-Host "  $($_.Address)  [$($_.Rtd)]  $($_.Dn)" -ForegroundColor Red }
    if (-not (Confirm-SmbGate -Title 'NOT A SHARED MAILBOX' -Problem "$($people.Count) object(s) are not shared mailboxes: $(@($people | ForEach-Object { "$($_.Address) [$($_.Rtd)]" }) -join ', '). Deleting them deletes a PERSON'S Active Directory account. Really continue?")) { return }
}
if (Test-Path -LiteralPath $clFile) {
    $held = @($cl | Where-Object { "$($_.LitigationHoldEnabled)" -eq 'True' -or "$($_.InPlaceHolds)".Trim() })
    if ($held.Count) {
        if (-not (Confirm-SmbGate -Title 'Mailboxes on hold' -Problem "$($held.Count) mailbox(es) are on litigation or in-place hold: $(@($held | ForEach-Object { $_.PrimarySmtpAddress }) -join ', '). The delete is retried with -IgnoreLegalHold if Exchange refuses; the cloud mailbox and its contents are soft-deleted, not purged, and come back in Step 6 -- but confirm the hold owner is happy first.")) { return }
    }
}

Write-Host ''
Write-Host "About to DELETE $($plan.Count) on-premises object(s):" -ForegroundColor Red
$plan | ForEach-Object { Write-Host "  $($_.Address.PadRight(45)) $($_.Dn)" }
$a = [System.Windows.Forms.MessageBox]::Show(
    ("PERMANENTLY DELETE the on-premises AD user object of $($plan.Count) mailbox(es) in batch $batchId?" + "`r`n`r`n  " +
     (@($plan | Select-Object -First 15 | ForEach-Object { $_.Address }) -join "`r`n  ") +
     $(if ($plan.Count -gt 15) { "`r`n  ... and $($plan.Count - 15) more" } else { '' }) +
     "`r`n`r`nThere is no undo. The next sync soft-deletes the cloud copies."),
    'Confirm -- irreversible', 'OKCancel', 'Warning')
if ($a -ne [System.Windows.Forms.DialogResult]::OK) { Write-Host 'Cancelled. Nothing was deleted.' -ForegroundColor Yellow; return }
$b = [System.Windows.Forms.MessageBox]::Show("Last check: delete $($plan.Count) object(s) now?", 'Confirm again', 'YesNo', 'Warning')
if ($b -ne [System.Windows.Forms.DialogResult]::Yes) { Write-Host 'Cancelled. Nothing was deleted.' -ForegroundColor Yellow; return }

$rows = @(); $ok = 0; $failed = 0
foreach ($p in $plan) {
    $res = ''
    # Learn (Remove-RemoteMailbox): -Identity, -DomainController, -Confirm.
    # Removes the on-premises mail-enabled user AND, after sync, the cloud
    # mailbox. Identity is the DN; the controller is the object's own.
    $rp = @{ Identity = $p.Dn; Confirm = $false; ErrorAction = 'Stop' }
    $srv = Get-MigrationAdServer -DistinguishedName $p.Dn
    if ($srv) { $rp['DomainController'] = $srv }
    try {
        try {
            Remove-RemoteMailbox @rp
        } catch {
            # Learn (Remove-RemoteMailbox): -IgnoreLegalHold. The on-premises
            # side refuses a held mailbox; the cloud copy is soft-deleted and
            # restored intact, so the hold loses nothing here.
            if ("$($_.Exception.Message)" -match '(?i)hold') {
                Write-Host "  (held -- retrying with -IgnoreLegalHold)" -ForegroundColor Yellow
                Remove-RemoteMailbox @rp -IgnoreLegalHold
            } else { throw }
        }
        $res = 'Deleted'; $ok++
        Write-Host "  deleted  $($p.Address)" -ForegroundColor Green
        Write-CompletedLog -Action 'SmbMigrationDelete' -Target (Get-LogTarget $p.Address $p.Name) -Result 'Success' -Environment 'OnPrem' -Detail "Remove-RemoteMailbox $($p.Dn) (batch $batchId)"
    } catch {
        $res = "FAILED: $($_.Exception.Message)"; $failed++
        Write-Host "  FAILED   $($p.Address)  $($_.Exception.Message)" -ForegroundColor Red
        Write-FailedLog -Action 'SmbMigrationDelete' -Target (Get-LogTarget $p.Address $p.Name) -ErrorRecord $_ -Environment 'OnPrem'
    }
    $rows += [PSCustomObject]@{ Batch = $batchId; PrimarySmtpAddress = $p.Address; Name = $p.Name; DisplayName = $p.Display; Sam = $p.Sam; Guid = $p.Guid; Dn = $p.Dn; Server = $srv; Result = $res; DeletedUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') }
}
$outFile = Get-SmbBatchFile -Folder $basePath -BatchId $batchId -Stage 'Deleted'
$rows | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8
Write-ExportDone -Path $outFile -Rows $rows.Count
Write-Host ''
Write-Host "$ok deleted, $failed failed." -ForegroundColor $(if ($failed) { 'Yellow' } else { 'Cyan' })
Write-Host "NEXT: run Step 5 (contacts) now. Sync runs every 30 minutes -- allow ~75 minutes, then Step 6: admin center > Users > Deleted users > search '$(Get-SmbTag -BatchId $batchId)' > select all > Restore." -ForegroundColor Cyan
Write-CompletedLog -Action 'SmbMigrationDelete' -Target "batch $batchId" -Result $(if ($failed -and $ok) { 'Partial' } elseif ($failed) { 'Failed' } else { 'Success' }) -Detail "$ok deleted, $failed failed -> $([System.IO.Path]::GetFileName($outFile))"
