#Requires -Version 5.1
# TOOLKIT-REQUIRES: OwnProcess  -- Set-RemoteMailbox is on-premises only and
#                                 must not share a process with a cloud-
#                                 connected console.
# TOOLKIT-HELP-BEGIN
# STEP 1 of the SMBMigration project. Starts a batch and tags the display
# names.
#
# Asks for a BATCH NUMBER (001, 002 ...), creates SMBMigration-<M-d>-<NNN>.csv
# -- today's month and day, so 10 October batch 1 is SMBMigration-10-10-001.csv
# -- opens it for you, and waits. Put one primary SMTP address per row under
# the PrimarySmtpAddress column, save, press Enter.
#
# Then each mailbox's DisplayName is prefixed ON-PREMISES with 'SMB-<M-d>-'.
# Only DisplayName: alias, addresses and UPN are never touched. Why: the
# cloud restore (Step 6) is done by hand in the admin center, whose Deleted
# users page searches display names -- the prefix turns a batch of hundreds
# into one search, one select-all, one Restore. Step 8 takes it off again.
# A mailbox already carrying the prefix is left alone, so re-running is safe.
#
# Output: SMBMigration-<batch>-Renamed.csv (address, old and new display name).
#
# NEXT: wait for a directory sync (~30 min) so the tag reaches the cloud,
# then Step 2.
# TOOLKIT-HELP-END
# TOOLKIT-HELP-RUNS: On-premises Exchange. Writes DisplayName on each mailbox in the batch.

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

# ---- batch number ----
$num = ''
while ($true) {
    $num = Show-InputDialog -Title '1 - ONPREM Rename -- new batch' -Prompt "Batch number for today ($(Get-Date -Format 'M-d')), 1-999:" -Default ''
    $num = "$num".Trim()
    if ($num -eq '') { Write-Host 'Cancelled.' -ForegroundColor Yellow; return }
    if ($num -match '^\d{1,3}$') { break }
    [System.Windows.Forms.MessageBox]::Show('Enter a number from 1 to 999.', 'Batch number', 'OK', 'Warning') | Out-Null
}
$batchId = (Get-Date -Format 'M-d') + '-' + ('{0:000}' -f [int]$num)
$batchFile = Get-SmbBatchFile -Folder $basePath -BatchId $batchId
$tag = Get-SmbTag -BatchId $batchId

$go = Show-StepInstructions -Title '1 - ONPREM Rename' -ScriptPath $PSCommandPath -Text @"
STEP 1 of 9 -- start batch $batchId and tag the display names.

* Batch file: $([System.IO.Path]::GetFileName($batchFile))
  It opens next. One primary SMTP address per row under PrimarySmtpAddress.
  Save it, come back here, press Enter.
* Each mailbox's DisplayName is then prefixed on-premises with '$tag'.
  Alias, addresses and UPN are not touched.
* Then WAIT for a directory sync (~30 min) before Step 2.

This window must be connected to ON-PREMISES Exchange. Click OK.
"@
if (-not $go) { return }

if (-not (Test-Path -LiteralPath $batchFile)) {
    'PrimarySmtpAddress' | Set-Content -LiteralPath $batchFile -Encoding UTF8
    Write-Host "Created $batchFile" -ForegroundColor Cyan
} else {
    Write-Host "Batch file already exists -- re-using $batchFile" -ForegroundColor Yellow
}
Invoke-Item -LiteralPath $batchFile
Write-Host ''
Read-Host "Fill in $([System.IO.Path]::GetFileName($batchFile)) -- one primary SMTP address per row -- save, then press ENTER" | Out-Null

$names = @(Read-SmbBatch -Folder $basePath -BatchId $batchId)
Write-Host "$($names.Count) mailbox(es) in batch $batchId." -ForegroundColor Cyan
if ($names.Count -ge 500) {
    if (-not (Confirm-SmbGate -Title 'Batch size' -Problem "$($names.Count) mailboxes in one batch. The Entra Connect export deletion threshold is 500 by default (other deletions in the same cycle count too) -- Step 4's deletes would be refused by the sync unless the threshold is disabled first.")) { return }
}

Import-Module ActiveDirectory -ErrorAction SilentlyContinue
Enter-MigrationOnPremScope

# RESOLVE EVERY ROW FIRST, WRITE NOTHING UNTIL THE WHOLE LIST RESOLVES. A
# typo on row 12 found after rows 1-11 were tagged is a half-tagged batch.
$plan = @(); $bad = @()
foreach ($n in $names) {
    try {
        $rm = Get-RemoteMailbox -Identity $n -ErrorAction Stop
        $plan += [PSCustomObject]@{ Address = "$($rm.PrimarySmtpAddress)"; Name = "$($rm.Name)"; Display = "$($rm.DisplayName)"; Dn = "$($rm.DistinguishedName)"; Rtd = "$($rm.RecipientTypeDetails)" }
    } catch { $bad += "$n -- $($_.Exception.Message)" }
}
# A PERSON'S MAILBOX IN THE LIST IS A TYPO, AND STEP 4 WOULD DELETE THEIR
# ACCOUNT. Caught here, at the first step, where nothing has happened yet.
$people = @($plan | Where-Object { "$($_.Rtd)" -notmatch '^Remote(Shared|Room|Equipment)Mailbox$' })
if ($people.Count) {
    Write-Host "$($people.Count) row(s) are NOT shared mailboxes:" -ForegroundColor Red
    $people | ForEach-Object { Write-Host "  $($_.Address)  [$($_.Rtd)]" -ForegroundColor Red }
    if (-not (Confirm-SmbGate -Title 'Not a shared mailbox' -Problem "$($people.Count) row(s) are not shared mailboxes (listed in the window): $(@($people | ForEach-Object { "$($_.Address) [$($_.Rtd)]" }) -join ', '). Step 4 would DELETE those accounts. Tag them anyway?")) { return }
}
if ($bad.Count) {
    Write-Host "$($bad.Count) address(es) did not resolve as a remote mailbox:" -ForegroundColor Red
    $bad | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    if (-not (Confirm-SmbGate -Title 'Unresolved rows' -Problem "$($bad.Count) of $($names.Count) rows did not resolve (listed in the window). Tag the $($plan.Count) that did?")) { return }
}
if ($plan.Count -eq 0) { Write-Host 'Nothing to tag.' -ForegroundColor Yellow; return }

$a = [System.Windows.Forms.MessageBox]::Show(
    ("Prefix the DisplayName of $($plan.Count) mailbox(es) with '$tag'?" + "`r`n`r`n  " +
     (@($plan | Select-Object -First 15 | ForEach-Object { "$($_.Address)   ($($_.Display))" }) -join "`r`n  ") +
     $(if ($plan.Count -gt 15) { "`r`n  ... and $($plan.Count - 15) more" } else { '' })),
    'Confirm', 'OKCancel', 'Question')
if ($a -ne [System.Windows.Forms.DialogResult]::OK) { Write-Host 'Cancelled. Nothing was changed.' -ForegroundColor Yellow; return }

$rows = @(); $ok = 0; $already = 0; $failed = 0
foreach ($p in $plan) {
    $res = ''; $new = $p.Display
    if (Test-SmbTagged $p.Display) {
        $res = 'AlreadyTagged'; $already++
        Write-Host "  already  $($p.Address)  ($($p.Display))" -ForegroundColor DarkGray
    } else {
        $new = $tag + $p.Display
        $sp = @{ Identity = $p.Dn; DisplayName = $new; Confirm = $false; ErrorAction = 'Stop' }
        $srv = Get-MigrationAdServer -DistinguishedName $p.Dn
        if ($srv) { $sp['DomainController'] = $srv }
        try {
            Set-RemoteMailbox @sp
            $res = 'Tagged'; $ok++
            Write-Host "  tagged   $($p.Address)  ->  $new" -ForegroundColor Green
        } catch {
            $res = "FAILED: $($_.Exception.Message)"; $failed++; $new = $p.Display
            Write-Host "  FAILED   $($p.Address)  $($_.Exception.Message)" -ForegroundColor Red
            Write-FailedLog -Action 'SmbMigrationTag' -Target (Get-LogTarget $p.Address $p.Name) -ErrorRecord $_
        }
    }
    $rows += [PSCustomObject]@{ Batch = $batchId; PrimarySmtpAddress = $p.Address; Name = $p.Name; OriginalDisplayName = (Remove-SmbTag $p.Display); TaggedDisplayName = $new; Tag = $tag; Dn = $p.Dn; Result = $res; WhenUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') }
}
$outFile = Get-SmbBatchFile -Folder $basePath -BatchId $batchId -Stage 'Renamed'
$rows | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8
Write-ExportDone -Path $outFile -Rows $rows.Count
Write-Host "Tagged $ok, already tagged $already, failed $failed." -ForegroundColor $(if ($failed) { 'Yellow' } else { 'Cyan' })
Write-Host "NEXT: wait for a directory sync (~30 min), then run Step 2 (on-premises export) and Step 3 (cloud export -- it checks the tag reached the cloud)." -ForegroundColor Cyan
Write-CompletedLog -Action 'SmbMigrationTag' -Target "batch $batchId" -Result $(if ($failed -and $ok) { 'Partial' } elseif ($failed) { 'Failed' } else { 'Success' }) -Detail "$ok tagged, $already already, $failed failed; prefix $tag -> $([System.IO.Path]::GetFileName($outFile))"
