#Requires -Version 5.1
# TOOLKIT-REQUIRES: OwnProcess  -- writes to on-premises AD; must not share
#                                 a process with a cloud-connected console.
# TOOLKIT-HELP-BEGIN
# UNDO for STEP 3 of the Migration project. Reads the newest
# Migration-Moved_*.csv and moves every object that was Moved back to the OU
# it came from.
#
# Run before a directory sync and nothing happened at all. Run after one and
# the objects re-enter sync scope on the next cycle: mailboxes re-match
# their (recycle-bin or restored) cloud users by ImmutableId; lists are
# recreated as NEW cloud objects, so anything Step 5 stamped on a recreated
# list is lost and the on-premises copy wins again.
#
# Rows whose Result was not 'Moved' are skipped -- they never left.
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
$moved = Get-NewestExport -Folder $basePath -Pattern 'Migration-Moved_*.csv'
if (-not $moved) { Write-Host 'No Migration-Moved_*.csv found. Nothing to undo.' -ForegroundColor Yellow; return }

$go = Show-StepInstructions -Title '3R - ONPREM Move Back (undo)' -ScriptPath $PSCommandPath -Text @"
UNDO STEP 3 -- move the objects back to their original OUs.

* Reads: $($moved.Name)
* Every row with Result = Moved goes back to its OriginalOu.
* After the next sync the cloud copies re-match (mailboxes) or are
  recreated (lists).

This window must be connected to ON-PREMISES Exchange. Click OK to continue.
"@
if (-not $go) { return }

Import-Module ActiveDirectory -ErrorAction Stop
Enter-MigrationOnPremScope
$rows = @(Import-Csv -LiteralPath $moved.FullName | Where-Object { $_.Result -eq 'Moved' })
if ($rows.Count -eq 0) { Write-Host 'No rows with Result = Moved. Nothing to undo.' -ForegroundColor Yellow; return }

$a = [System.Windows.Forms.MessageBox]::Show("Move $($rows.Count) object(s) back to their original OUs?", 'Confirm', 'OKCancel', 'Question')
if ($a -ne [System.Windows.Forms.DialogResult]::OK) { Write-Host 'Cancelled.' -ForegroundColor Yellow; return }

$ok = 0; $failed = 0
foreach ($r in $rows) {
    try {
        # Same controller Step 3 used (Server column); older Moved files
        # without it are resolved from the original DN.
        $srv = "$($r.Server)"; if (-not $srv) { $srv = Get-MigrationAdServer -DistinguishedName "$($r.OriginalDn)" }
        $sp = @{}; if ($srv) { $sp['Server'] = $srv }
        Move-ADObject -Identity $r.Guid -TargetPath $r.OriginalOu @sp -ErrorAction Stop
        $ok++; Write-Host "  back  $($r.PrimarySmtpAddress)  ->  $($r.OriginalOu)" -ForegroundColor Green
    } catch {
        $failed++; Write-Host "  FAIL  $($r.PrimarySmtpAddress)  $($_.Exception.Message)" -ForegroundColor Red
        Write-FailedLog -Action 'MigrationMoveBack' -Target (Get-LogTarget $r.PrimarySmtpAddress $r.Name) -ErrorRecord $_
    }
}
Write-Host ''
Write-Host "$ok moved back, $failed failed." -ForegroundColor Cyan
Write-CompletedLog -Action 'MigrationMoveBack' -Target "$($rows.Count) objects" -Result $(if ($failed -and $ok) { 'Partial' } elseif ($failed) { 'Failed' } else { 'Success' }) `
    -Detail "$ok moved back, $failed failed (from $($moved.Name))"
