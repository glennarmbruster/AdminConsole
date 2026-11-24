#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# STEP 8 of the SMBMigration project. Takes the Step 1 tag ('SMB-<M-d>-')
# off the DisplayName of every restored mailbox in the batch, in Exchange
# Online. DisplayName only; nothing else is touched.
#
# The clean name comes from the Step 2 export where it exists (that is the
# name the mailbox had before any of this), otherwise from stripping the
# prefix off the current cloud name. A mailbox whose name already has no
# tag is left alone. A mailbox not yet restored is reported and skipped.
#
# Output: SMBMigration-<batch>-RenamedCloud.csv. Needs Exchange Online.
# TOOLKIT-HELP-END
# TOOLKIT-HELP-RUNS: Exchange Online. Writes DisplayName on each restored mailbox in the batch.

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

function exo { param([string]$n, [hashtable]$p) Invoke-EXOCommand -Name $n -Parameters $p }

$batchId = Select-SmbBatch -Folder $basePath -Title '8 - CLOUD Rename'
if (-not $batchId) { return }
$onFile = Get-SmbBatchFile -Folder $basePath -BatchId $batchId -Stage 'ONPREM'
$export = @()
if (Test-Path -LiteralPath $onFile) { $export = @(Import-Csv -LiteralPath $onFile | Where-Object { -not "$($_.Error)" }) }

$go = Show-StepInstructions -Title '8 - CLOUD Rename' -ScriptPath $PSCommandPath -Text @"
STEP 8 of 9 -- remove the '$(Get-SmbTag -BatchId $batchId)' tag from the display names of batch $batchId.

* DisplayName only, in Exchange Online.
* Clean name from the Step 2 export ($(if ($export.Count) { "$($export.Count) rows" } else { 'MISSING -- the prefix is stripped instead' })).

Exchange Online must be connected in this window. Click OK to run.
"@
if (-not $go) { return }

$names = @(Read-SmbBatch -Folder $basePath -BatchId $batchId)
$plan = @(); $skip = @()
foreach ($n in $names) {
    $mb = $null
    try { $mb = exo 'Get-Mailbox' @{ Identity = $n; ErrorAction = 'Stop' } } catch { }
    if (-not $mb) { $skip += "$n -- not active in Exchange Online yet (restore it in the admin center, Step 6)"; continue }
    if ("$($mb.IsDirSynced)" -eq 'True') { $skip += "$n -- still directory-synced; a cloud rename would be overwritten"; continue }
    $cur = "$($mb.DisplayName)"
    if (-not (Test-SmbTagged $cur)) { $skip += "$n -- no tag on '$cur'"; continue }
    $row = $export | Where-Object { "$($_.PrimarySmtpAddress)" -ieq $n } | Select-Object -First 1
    $clean = if ($row -and "$($row.DisplayName)".Trim()) { "$($row.DisplayName)".Trim() } else { Remove-SmbTag $cur }
    $plan += [PSCustomObject]@{ Address = "$($mb.PrimarySmtpAddress)"; Current = $cur; Clean = $clean }
}
if ($skip.Count) { Write-Host 'Not renamed:' -ForegroundColor Yellow; $skip | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow } }
if ($plan.Count -eq 0) { Write-Host 'Nothing to rename.' -ForegroundColor Yellow; return }

$a = [System.Windows.Forms.MessageBox]::Show(
    ("Set the display name of $($plan.Count) mailbox(es)?" + "`r`n`r`n  " +
     (@($plan | Select-Object -First 15 | ForEach-Object { "$($_.Current)  ->  $($_.Clean)" }) -join "`r`n  ") +
     $(if ($plan.Count -gt 15) { "`r`n  ... and $($plan.Count - 15) more" } else { '' })),
    'Confirm', 'OKCancel', 'Question')
if ($a -ne [System.Windows.Forms.DialogResult]::OK) { Write-Host 'Cancelled. Nothing was changed.' -ForegroundColor Yellow; return }

$rows = @(); $ok = 0; $failed = 0
foreach ($p in $plan) {
    $res = ''
    try {
        exo 'Set-Mailbox' @{ Identity = $p.Address; DisplayName = $p.Clean; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
        $res = 'Renamed'; $ok++
        Write-Host "  renamed  $($p.Address)  ->  $($p.Clean)" -ForegroundColor Green
        Write-CompletedLog -Action 'SmbMigrationRenameCloud' -Target (Get-LogTarget $p.Address) -Result 'Success' -Environment 'Cloud' -Detail "'$($p.Current)' -> '$($p.Clean)' (batch $batchId)"
    } catch {
        $res = "FAILED: $($_.Exception.Message)"; $failed++
        Write-Host "  FAILED   $($p.Address)  $($_.Exception.Message)" -ForegroundColor Red
        Write-FailedLog -Action 'SmbMigrationRenameCloud' -Target (Get-LogTarget $p.Address) -ErrorRecord $_ -Environment 'Cloud'
    }
    $rows += [PSCustomObject]@{ Batch = $batchId; PrimarySmtpAddress = $p.Address; Before = $p.Current; After = $p.Clean; Result = $res; WhenUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') }
}
foreach ($sk in $skip) { $rows += [PSCustomObject]@{ Batch = $batchId; PrimarySmtpAddress = ($sk -split ' -- ')[0]; Before = ''; After = ''; Result = "SKIPPED: $(($sk -split ' -- ', 2)[1])"; WhenUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') } }
$outFile = Get-SmbBatchFile -Folder $basePath -BatchId $batchId -Stage 'RenamedCloud'
$rows | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8
Write-ExportDone -Path $outFile -Rows $rows.Count
Write-Host "$ok renamed, $failed failed, $($skip.Count) skipped. NEXT: Step 9 (verify)." -ForegroundColor $(if ($failed) { 'Yellow' } else { 'Cyan' })
Write-CompletedLog -Action 'SmbMigrationRenameCloud' -Target "batch $batchId" -Result $(if ($failed -and $ok) { 'Partial' } elseif ($failed) { 'Failed' } else { 'Success' }) -Detail "$ok renamed, $failed failed, $($skip.Count) skipped -> $([System.IO.Path]::GetFileName($outFile))"
