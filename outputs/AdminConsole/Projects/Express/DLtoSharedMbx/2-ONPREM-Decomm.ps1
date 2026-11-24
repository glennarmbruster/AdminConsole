#Requires -Version 5.1
# TOOLKIT-REQUIRES: OwnProcess  -- Disable- and Remove-DistributionGroup exist
#                                 in BOTH Exchanges. This script means the
#                                 on-premises one, and running it against the
#                                 cloud by accident would be unrecoverable.
# TOOLKIT-HELP-BEGIN
# STEP 2 of the DL cloud migration, and the only irreversible one.
#
# Reads the DGExport file from Step 1 and, for each list:
#
#   * mail-disables it on-premises, always
#   * DELETES it, but ONLY where NewName is blank
#
# A row WITH a NewName -- the Payroll lists -- is mail-disabled and kept. That
# distinction is the whole point of the NewName column, and it is what stops a
# payroll group being destroyed.
#
# Deletion cannot be undone from here. The DGExport file is the only record of
# what those groups contained, so keep it.
#
# Afterwards, wait for directory sync to finish before running Step 3.
# TOOLKIT-HELP-END
# Script 2 - Decommission DLs on-prem (SAFE VERSION)
# Input: Most recent dgexport_*.csv (auto-selected with override option)

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

# The scripts live in the folder they work in, so $PSScriptRoot IS the base
# path -- same location as before, and correct on someone else's machine
# without editing anything. Was hard-coded to E:\scripts\... which only ever
# worked on one box.
$basePath = $PSScriptRoot

# ---------------------------------------------------------
# Find most recent PRIMARY DGExport file (exclude compare files)
# ---------------------------------------------------------
$files = Get-ChildItem -Path $basePath -Filter "dgexport_*.csv" |
         Where-Object { $_.Name -notmatch "compare" } |
         Sort-Object LastWriteTime -Descending

if (-not $files) {
    Write-Host "No primary dgexport_*.csv files found in $basePath" -ForegroundColor Red
    return
}

$autoFile = $files[0].FullName
$autoName = $files[0].Name

# ---------------------------------------------------------
# Instruction Popup with 3 Options
# ---------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Script 2 - Select DGExport File"
$form.Width = 520
$form.Height = 260
$form.StartPosition = "CenterScreen"

$label = New-Object System.Windows.Forms.Label
$label.Text = @"
Script 2 will decommission on-prem DLs based on the DGExport file.

The most recent DGExport file detected is:

    **$autoName**

Choose an option:
* OK -- Use this file
* Other File -- Select a different DGExport file
* Cancel -- Stop the script
"@
$label.AutoSize = $true
$label.Top = 10
$label.Left = 10
$form.Controls.Add($label)

# OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Width = 100
$okButton.Top = 170
$okButton.Left = 120
$okButton.Add_Click({ $form.Tag = "OK"; $form.Close() })
$form.Controls.Add($okButton)

# Other File button
$otherButton = New-Object System.Windows.Forms.Button
$otherButton.Text = "Other File"
$otherButton.Width = 100
$otherButton.Top = 170
$otherButton.Left = 230
$otherButton.Add_Click({ $form.Tag = "Other"; $form.Close() })
$form.Controls.Add($otherButton)

# Cancel button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Width = 100
$cancelButton.Top = 170
$cancelButton.Left = 340
$cancelButton.Add_Click({ $form.Tag = "Cancel"; $form.Close() })
$form.Controls.Add($cancelButton)

# Theme + the standard '?' in the top-right, same as every other tool.
Set-ConsoleTheme -Form $form
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title '2-Onprem-Decomm' | Out-Null

$form.ShowDialog() | Out-Null

switch ($form.Tag) {
    "Cancel" {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        return
    }
    "OK" {
        $fullPath = $autoFile
    }
    "Other" {
        # File-select dialog
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.InitialDirectory = $basePath
        $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
        $openFileDialog.Title = "Select DGExport file"

        if ($openFileDialog.ShowDialog() -ne "OK") {
            Write-Host "No file selected. Script cancelled." -ForegroundColor Yellow
            return
        }

        $fullPath = $openFileDialog.FileName
    }
}

Write-Host "Using DGExport file: $fullPath"

# ---------------------------------------------------------
# Begin Script 2 Logic
# ---------------------------------------------------------

$csv = Import-Csv $fullPath

foreach ($row in $csv) {
    $dgIdentity = $row.DGIdentity
    $newName    = $row.NewName

    Write-Host ""
    Write-Host "Processing DL: $($row.DGName)"

    $dg = Get-DistributionGroup -Identity $dgIdentity -ErrorAction SilentlyContinue
    if (-not $dg) {
        Write-Host "  SKIPPED: DL not found on-prem."
        continue
    }

    try {
        Write-Host "  Mail-disabling DL..."
        Disable-DistributionGroup -Identity $dgIdentity -Confirm:$false -ErrorAction Stop

        if ($newName -eq "") {
            Write-Host "  Deleting NON-PAYROLL DL..."
            Remove-DistributionGroup -Identity $dgIdentity -Confirm:$false -ErrorAction Stop
        }
        else {
            Write-Host "  PAYROLL DL detected -- NOT deleting."
        }

        Write-Host "  DL decommissioned."
    }
    catch {
        Write-Host "  ERROR: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "On-prem decommission complete. Wait for AAD Connect sync before running Script 3."
