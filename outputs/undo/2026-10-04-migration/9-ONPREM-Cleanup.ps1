#Requires -Version 5.1
# TOOLKIT-REQUIRES: OwnProcess  -- deletes on-premises AD objects; must not
#                                 share a process with a cloud-connected
#                                 console.
# TOOLKIT-HELP-BEGIN
# STEP 9 of the Migration project. The last one. Deletes the on-premises AD
# objects that Step 5 parked in the NonSync OU, once the cloud side is
# verified.
#
# WHY DELETE AT ALL. An object left in that OU is a loaded gun. Anyone who
# drags it back into a synced OU -- tidying up, "fixing" a broken user --
# re-matches the cloud-only object by ImmutableId on the next sync, and the
# on-premises copy takes control again: attributes overwritten, and if it
# is a list, a second copy fighting the first over the same addresses.
#
# THE GATE. It refuses to run unless a Migration-Verify_*.csv exists that is
# NEWER than the move file and has no MISSING or DIFFERENT rows. The verify
# is the evidence that there is nothing left on-premises worth keeping.
#
# WHAT IT DOES, PER ROW OF THE MOVED FILE WITH Result = Moved:
#   * confirms the object still sits in the NonSync OU (if somebody already
#     moved it back, it is reported and left alone -- that is a decision,
#     not a cleanup)
#   * Remove-ADObject. Users and groups alike. Objects flagged "protect from
#     accidental deletion" fail and are reported; clear the flag by hand.
#
# IRREVERSIBLE. The AD recycle bin, if enabled, is the only way back, and
# the cloud-only object neither needs nor wants it back. This asks twice.
#
# Output lands next to this script:
#     Migration-Cleaned_<date>_<time>.csv
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

$movedFile = Get-NewestExport -Folder $basePath -Pattern 'Migration-Moved_*.csv'
if (-not $movedFile) { Write-Host 'No Migration-Moved_*.csv found. Nothing to clean up.' -ForegroundColor Yellow; return }

# ---- the gate ----
$verify = Get-NewestExport -Folder $basePath -Pattern 'Migration-Verify_*.csv'
$why = ''
if (-not $verify) { $why = 'No Migration-Verify_*.csv exists. Run Step 8 first.' }
elseif ($verify.LastWriteTime -lt $movedFile.LastWriteTime) { $why = "$($verify.Name) is OLDER than $($movedFile.Name). Run Step 8 again after the move." }
else {
    $bad = @(Import-Csv -LiteralPath $verify.FullName | Where-Object { $_.Result -in 'MISSING', 'DIFFERENT' })
    if ($bad.Count) { $why = "$($verify.Name) has $($bad.Count) MISSING/DIFFERENT row(s). Fix them (or re-run Step 7 and Step 8) first." }
}
if ($why) {
    [System.Windows.Forms.MessageBox]::Show(("REFUSED -- the cloud side is not verified." + "`r`n`r`n" + $why + "`r`n`r`nNothing was deleted."), '9 - ONPREM Cleanup', 'OK', 'Warning') | Out-Null
    Write-Host "STOPPED: $why" -ForegroundColor Yellow
    return
}

$rows = @(Import-Csv -LiteralPath $movedFile.FullName | Where-Object { $_.Result -eq 'Moved' })
if ($rows.Count -eq 0) { Write-Host "$($movedFile.Name) has no Moved rows. Nothing to clean up." -ForegroundColor Yellow; return }

$go = Show-StepInstructions -Title '9 - ONPREM Cleanup' -ScriptPath $PSCommandPath -Text @"
STEP 9 -- DELETE the parked on-premises objects. Irreversible.

* Reads:    $($movedFile.Name)  ($($rows.Count) moved object(s))
* Verified: $($verify.Name)  (no MISSING / DIFFERENT rows)
* Each object is checked to still be in the NonSync OU, then removed
  from Active Directory. Objects that were moved back are left alone.

This window must be connected to ON-PREMISES Exchange. Click OK to see the
list; nothing is deleted until you confirm again.
"@
if (-not $go) { return }

Import-Module ActiveDirectory -ErrorAction Stop

# ---- check before touching ----
$plan = @()
foreach ($r in $rows) {
    $obj = $null
    try { $obj = Get-ADObject -Identity $r.Guid -Properties distinguishedName, objectClass, ProtectedFromAccidentalDeletion -ErrorAction Stop } catch { }
    if (-not $obj) { $plan += [PSCustomObject]@{ Row = $r; State = 'already gone'; Do = $false; Dn = '' }; continue }
    $parent = ($obj.DistinguishedName -split '(?<!\\),', 2)[1]
    if ($parent -ine $r.TargetOu) { $plan += [PSCustomObject]@{ Row = $r; State = "MOVED BACK to $parent -- left alone"; Do = $false; Dn = $obj.DistinguishedName }; continue }
    $st = $(if ($obj.ProtectedFromAccidentalDeletion) { 'protected -- will fail unless cleared' } else { 'in NonSync OU -- will delete' })
    $plan += [PSCustomObject]@{ Row = $r; State = $st; Do = $true; Dn = $obj.DistinguishedName; Protected = [bool]$obj.ProtectedFromAccidentalDeletion }
}
Write-Host ''
foreach ($p in $plan) { Write-Host ("  {0,-4} {1,-45} {2}" -f $p.Row.Type, $p.Row.PrimarySmtpAddress, $p.State) -ForegroundColor $(if ($p.Do) { 'White' } else { 'Yellow' }) }
$todo = @($plan | Where-Object { $_.Do })
if ($todo.Count -eq 0) { Write-Host 'Nothing to delete.' -ForegroundColor Yellow; return }

$a = [System.Windows.Forms.MessageBox]::Show(
    ("PERMANENTLY DELETE $($todo.Count) on-premises object(s) from Active Directory?" + "`r`n`r`n  " +
     (($todo | ForEach-Object { "$($_.Row.Type)  $($_.Row.PrimarySmtpAddress)" }) -join "`r`n  ") +
     "`r`n`r`nThe cloud-only copies are verified and are not affected. There is no undo."),
    'Confirm -- irreversible', 'OKCancel', 'Warning')
if ($a -ne [System.Windows.Forms.DialogResult]::OK) { Write-Host 'Cancelled. Nothing was deleted.' -ForegroundColor Yellow; return }
$b = [System.Windows.Forms.MessageBox]::Show("Last check: delete $($todo.Count) object(s) now?", 'Confirm again', 'YesNo', 'Warning')
if ($b -ne [System.Windows.Forms.DialogResult]::Yes) { Write-Host 'Cancelled. Nothing was deleted.' -ForegroundColor Yellow; return }

$out = @(); $ok = 0; $failed = 0
foreach ($p in $plan) {
    $res = $p.State
    if ($p.Do) {
        try {
            # Learn (Remove-ADObject): -Identity <GUID/DN>, -Recursive for
            # containers only -- not used, these are leaf objects.
            Remove-ADObject -Identity $p.Row.Guid -Confirm:$false -ErrorAction Stop
            $res = 'Deleted'; $ok++
            Write-Host "  deleted  $($p.Row.PrimarySmtpAddress)" -ForegroundColor Green
            Write-CompletedLog -Action 'MigrationCleanup' -Target (Get-LogTarget $p.Row.PrimarySmtpAddress $p.Row.Name) -Result 'Success' -Detail "Remove-ADObject $($p.Dn)"
        } catch {
            $res = "FAILED: $($_.Exception.Message)"; $failed++
            Write-Host "  FAIL     $($p.Row.PrimarySmtpAddress)  $($_.Exception.Message)" -ForegroundColor Red
            Write-FailedLog -Action 'MigrationCleanup' -Target (Get-LogTarget $p.Row.PrimarySmtpAddress $p.Row.Name) -ErrorRecord $_
        }
    }
    $out += [PSCustomObject]@{ Type = $p.Row.Type; Name = $p.Row.Name; PrimarySmtpAddress = $p.Row.PrimarySmtpAddress; Guid = $p.Row.Guid; Dn = $p.Dn; Result = $res; CleanedUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') }
}
$outFile = Join-Path $basePath "Migration-Cleaned_$(Get-MigrationTimestamp).csv"
$out | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding UTF8
Write-ExportDone -Path $outFile -Rows $out.Count
Write-Host "$ok deleted, $failed failed. Move this batch's exports to Exports-Complete\ -- the project is done for these objects." -ForegroundColor Cyan
Write-CompletedLog -Action 'MigrationCleanup' -Target "$($todo.Count) objects" -Result $(if ($failed -and $ok) { 'Partial' } elseif ($failed) { 'Failed' } else { 'Success' }) -Detail "$ok deleted, $failed failed -> $([System.IO.Path]::GetFileName($outFile))"
