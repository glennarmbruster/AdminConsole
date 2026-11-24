#Requires -Version 5.1
# TOOLKIT-REQUIRES: OwnProcess  -- Get-DistributionGroup,
#                                 Get-DistributionGroupMember and
#                                 Set-DistributionGroup all exist in BOTH
#                                 Exchanges. This script means the on-premises
#                                 one, and must not share a process with a
#                                 cloud-connected console. The Set- matters
#                                 most: hiding the wrong tenant's lists from
#                                 the address book would be a live incident.
# TOOLKIT-HELP-BEGIN
# STEP 1 of the DL cloud migration. Exports the distribution lists named in
# DL-Input.csv, on-premises, so the later steps have something to work from.
#
# Before running, DL-Input.csv must have two columns -- Name,NewName. Every row
# needs the comma; NewName may be blank. The script opens the file for you to
# review and waits for you to finish.
#
#     Jobs.ColoradoSpringsCO,
#     Payroll_F26,Payroll.ColoradoSpringsCO
#
# It exports TWICE and compares the two files. If the comparison fails, the
# directory changed underneath you mid-export and the result is not safe to
# build on -- do not carry on to Step 2.
#
# THEN IT MAKES ONE CHANGE. Every row with a NewName -- the Payroll lists --
# is hidden from the address book. Nothing else is touched: no mail attributes
# are removed and nothing is deleted. That is all Step 2.
#
# The hide is why this step exists before Step 2 rather than beside it. Mail-
# disabling does not remove the cloud copy of a synced group; it keeps syncing
# and Exchange Online falls back to its @...onmicrosoft.com address, leaving it
# in the address book as a recipient that quietly swallows mail. Hiding first
# and letting a sync carry it avoids that.
#
# So: run this, WAIT ABOUT 30 MINUTES for a directory sync, then run Step 2.
#
# The hide only runs if the export comparison passed. A directory that was
# shifting mid-export is not one to start writing to.
#
# Re-running is safe. Groups already hidden are reported and skipped.
#
# Output lands next to this script:
#     dgexport_<date>_<time>.csv          the export
#     dgexport_<date>_<time>-compare.csv  the second export, for the check
#     payrollhide_<date>_<time>.csv       what was hidden, before and after
# TOOLKIT-HELP-END
# Script 1 - Export DLs for conversion (on-prem)
# Input:  DL-Input.csv  (Name,NewName)
# Output: dgexport_M-D-YYYY_HHmm.csv + compare file + validation check
# Includes:
#   - Updated instruction popup (consistent with Scripts 3 & 4)
#   - Input file review step
#   - Writes to E:\scripts\m365\adminconsole\Projects\dlconversion

Add-Type -AssemblyName System.Windows.Forms

# Dot-sourced from two levels up: these live in Projects\DLtoSharedMbx, the
# toolkit root is AdminConsole. Brings in the theme, the audit log, and
# Invoke-EXOCommand.
# Common.ps1 lives at the TOOLKIT ROOT and there is exactly one of it. This
# walks up to find it rather than counting folders.
#
# It used to be two Split-Path -Parent calls, which resolved to
# Projects\Common.ps1 -- one level short of the root -- so every script here
# failed on launch with "Common.ps1 is not recognized", followed by a cascade of
# Set-ConsoleTheme / Add-ToolHelpButton not found.
#
# Counting is the wrong tool: these scripts sit three levels down, Set-Managers\
# sits four, and the next project may sit somewhere else again. Walking up works
# at any depth and keeps a single shared copy -- a Common.ps1 dropped into a
# project folder would be a fork nobody reviews.
$toolkitRoot = $PSScriptRoot
while ($toolkitRoot -and -not (Test-Path -LiteralPath (Join-Path $toolkitRoot 'Common.ps1'))) {
    $toolkitRoot = Split-Path -Parent $toolkitRoot
}
if (-not $toolkitRoot) {
    throw "Common.ps1 was not found in any folder above '$PSScriptRoot'. This script must live under the toolkit folder."
}
. (Join-Path $toolkitRoot 'Common.ps1')

# ---------------------------------------------------------
# Instruction Popup
# ---------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Script 1 Instructions"
$form.Width = 520
# 350, not the original 260: the instruction text grew from 11 lines to 15
# when the hide was explained, and at 15px a line that reaches y=235 -- past
# the buttons. Measured rather than eyeballed; see ARCHITECTURE 14l.
$form.Height = 350
$form.StartPosition = "CenterScreen"

$label = New-Object System.Windows.Forms.Label
$label.Text = @"
Before running Script 1, please confirm the following:

* DL-Input.csv must contain two columns: Name,NewName
* Every row must include a comma (NewName may be blank)
* Example:
    Jobs.ColoradoSpringsCO,
    Payroll_F26,Payroll.ColoradoSpringsCO

This exports the lists, then HIDES every row that has a NewName
from the address book. Nothing is mail-disabled or deleted here.

Afterwards, wait about 30 minutes for a directory sync, then
run Script 2.

Click OK to continue or Cancel to stop.
"@
$label.AutoSize = $true
$label.Top = 10
$label.Left = 10
$form.Controls.Add($label)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Width = 80
$okButton.Top = 265
$okButton.Left = 300
$okButton.Add_Click({ $form.Tag = "OK"; $form.Close() })
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Width = 80
$cancelButton.Top = 265
$cancelButton.Left = 390
$cancelButton.Add_Click({ $form.Tag = "Cancel"; $form.Close() })
$form.Controls.Add($cancelButton)

# Theme + the standard '?' in the top-right, same as every other tool.
Set-ConsoleTheme -Form $form
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title '1-Onprem-Export' | Out-Null

$form.ShowDialog() | Out-Null

if ($form.Tag -eq "Cancel") {
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
    return
}

# ---------------------------------------------------------
# Begin Script 1 Logic
# ---------------------------------------------------------

# The scripts live in the folder they work in, so $PSScriptRoot IS the base
# path -- same location as before, and correct on someone else's machine
# without editing anything. Was hard-coded to E:\scripts\... which only ever
# worked on one box.
$basePath  = $PSScriptRoot
$inputFile = Join-Path $basePath "DL-Input.csv"

Write-Host "Input file: $inputFile"
Write-Host "Opening DL-Input.csv for review..."

# Open the input file for editing
ii $inputFile

# Pause until user confirms they are done editing
Write-Host ""
Read-Host "Review DL-Input.csv, save changes, then press ENTER to continue"

# Load input (NO validation, NO filtering)
$dlList = Import-Csv $inputFile

# Timestamp for filenames (M-d-yyyy_HHmm)
$timestamp = (Get-Date -Format "M-d-yyyy_HHmm")

# Output filenames
$export1 = Join-Path $basePath "dgexport_$timestamp.csv"
$export2 = Join-Path $basePath "dgexport_$timestamp-compare.csv"

Write-Host "Starting DL export for $timestamp"
Write-Host "Output directory: $basePath"

function Export-DGs {
    param([string]$path)

    $results = foreach ($row in $dlList) {
        $name    = $row.Name
        $newName = $row.NewName

        $dg = Get-DistributionGroup -Identity $name -ErrorAction Stop 

        $owners = $dg.ManagedBy -join ";"

        $members = Get-DistributionGroupMember -Identity $dg.Identity -ErrorAction SilentlyContinue |
                   Where-Object { $_.PrimarySmtpAddress } |
                   Select-Object -ExpandProperty PrimarySmtpAddress

        $emails = $dg.EmailAddresses | ForEach-Object { $_.ToString() }

        [PSCustomObject]@{
            DGName     = $dg.Name
            DGIdentity = $dg.Identity
            Owners     = $owners
            Members    = $members -join ";"
            Emails     = $emails  -join ";"
            NewName    = $newName
        }
    }

    $results | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
}

# Run export twice
Write-Host "Running first export: $export1"
Export-DGs -path $export1

Write-Host "Running second export: $export2"
Export-DGs -path $export2

# Compare the two files
Write-Host "Comparing exports..."
$diff = Compare-Object -ReferenceObject (Get-Content $export1) -DifferenceObject (Get-Content $export2)

if ($diff) {
    Write-Host "Export Validation FAILED - manually check before proceeding" -ForegroundColor Red
} else {
    Write-Host "Check Complete - export is valid" -ForegroundColor Green
}

# ---------------------------------------------------------
# Hide the payroll lists from the address book
#
# WHY: mail-disabling in Step 2 does not remove the cloud object. The AD group
# keeps syncing, and Exchange Online promotes the @...onmicrosoft.com address
# it appended at first sync to primary -- so the list stays in the GAL as a
# recipient that accepts mail and silently drops it. That is KB 4096314, not a
# fault in Step 2. Hiding here, then letting a sync carry it before Step 2
# runs, is the cheap mitigation.
#
# WHY IT RUNS LAST, AND ONLY ON A CLEAN COMPARE: the two exports above exist to
# prove the directory is not shifting underneath this run. Writing before that
# check would make this script the thing that changes it -- and the compare
# would then be reporting our own edit back to us. A failed compare means the
# export is not safe to build on, so it is not safe to act on either.
#
# WHICH GROUPS: NewName non-blank -- the SAME test Step 2 uses to decide what
# to keep. Matching on a 'payroll' name prefix instead would create a second
# definition of payroll that could disagree with Step 2's, and any group the
# two disagreed about would get the wrong treatment.
#
# NOT VERIFIED, and the CSV below exists to settle it: Disable-DistributionGroup
# removes "the email attributes required by Exchange", and
# msExchHideFromAddressLists may be one of them. If it is cleared in Step 2, the
# next sync un-hides the group and this mitigation regresses -- silently, and a
# sync cycle later. Compare HiddenAfter in this file against the groups once
# Step 2 has run and synced. If it regressed, the fallback is deleting the
# cloud object in the Entra portal and re-syncing.
# ---------------------------------------------------------
if ($diff) {
    Write-Host ""
    Write-Host "Skipping the address-book hide -- the export comparison failed." -ForegroundColor Yellow
    Write-Host "Nothing in the directory was changed. Resolve the export first." -ForegroundColor Yellow
} else {
    $payrollRows = @($dlList | Where-Object { $_.NewName -and "$($_.NewName)".Trim() -ne '' })

    if ($payrollRows.Count -eq 0) {
        Write-Host ""
        Write-Host "No rows with a NewName, so nothing to hide."
    } else {
        Write-Host ""
        Write-Host "Hiding $($payrollRows.Count) payroll list(s) from the address book..."

        $hideLog = foreach ($row in $payrollRows) {
            $name = $row.Name
            $rec  = [PSCustomObject]@{
                Name         = $name
                Guid         = ''
                HiddenBefore = ''
                HiddenAfter  = ''
                Result       = ''
            }

            $dg = Get-DistributionGroup -Identity $name -ErrorAction SilentlyContinue
            if (-not $dg) {
                Write-Host "  $name -- NOT FOUND on-prem." -ForegroundColor Red
                $rec.Result = 'NotFound'
                $rec
                continue
            }

            $rec.Guid         = "$($dg.Guid)"
            $rec.HiddenBefore = "$($dg.HiddenFromAddressListsEnabled)"

            if ($dg.HiddenFromAddressListsEnabled) {
                Write-Host "  $name -- already hidden."
                $rec.HiddenAfter = 'True'
                $rec.Result      = 'AlreadyHidden'
                $rec
                continue
            }

            try {
                # Bound on Guid, not on the name that was just looked up: a name
                # is a display value and can be ambiguous, a Guid cannot.
                #
                # Quoted, and that is not cosmetic. $dg.Guid is a [System.Guid]
                # OBJECT, and DistributionGroupIdParameter refuses to convert
                # from that type -- "Cannot process argument transformation".
                # It takes the STRING form happily. Interpolation forces it.
                Set-DistributionGroup -Identity "$($dg.Guid)" -HiddenFromAddressListsEnabled $true -ErrorAction Stop
                $after = Get-DistributionGroup -Identity "$($dg.Guid)" -ErrorAction SilentlyContinue
                $rec.HiddenAfter = "$($after.HiddenFromAddressListsEnabled)"
                $rec.Result      = 'Hidden'
                Write-Host "  $name -- hidden." -ForegroundColor Green
            } catch {
                $rec.Result = "Error: $($_.Exception.Message)"
                Write-Host "  $name -- ERROR: $($_.Exception.Message)" -ForegroundColor Red
            }
            $rec
        }

        $hideFile = Join-Path $basePath "payrollhide_$timestamp.csv"
        $hideLog | Export-Csv -Path $hideFile -NoTypeInformation -Encoding UTF8

        $done = @($hideLog | Where-Object { $_.Result -eq 'Hidden' }).Count
        $already = @($hideLog | Where-Object { $_.Result -eq 'AlreadyHidden' }).Count
        $bad = @($hideLog | Where-Object { $_.Result -notin @('Hidden', 'AlreadyHidden') }).Count

        Write-Host ""
        Write-Host "Hidden: $done   Already hidden: $already   Problems: $bad"
        Write-Host "Hide record: $hideFile"
        if ($bad -gt 0) {
            Write-Host "Check the problems above before running Step 2." -ForegroundColor Red
        }

        Write-CompletedLog -Task "Hid $done payroll DL(s) from address lists (DLtoSharedMbx step 1)"
    }
}

Write-Host ""
Write-Host "Export complete. Files saved to: $basePath"
Write-Host "Primary export: $export1"
Write-Host "Compare file:  $export2"
Write-Host ""
Write-Host "NEXT: wait for a directory sync to carry the hide to the cloud" -ForegroundColor Cyan
Write-Host "      (allow ~30 minutes), then run Step 2." -ForegroundColor Cyan
