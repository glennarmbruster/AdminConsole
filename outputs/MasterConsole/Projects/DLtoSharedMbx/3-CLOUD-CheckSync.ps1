#Requires -Version 5.1
# TOOLKIT-HELP-BEGIN
# STEP 3 of the DL cloud migration. Says whether Step 4 will succeed yet.
#
# THIS IS THE CLOCK THAT MATTERS. From the moment Step 2 runs until Step 4
# finishes, these addresses do not exist and mail to them bounces. Listservs
# unsubscribe addresses that bounce. Run this, and the moment it says ALL CLEAR,
# run Step 4 -- do not leave it sitting.
#
# It asks a different question for each kind of row, because they are different
# questions:
#
#   * DELETED rows (no NewName) -- has the object gone from the cloud yet?
#
#   * PAYROLL rows (NewName set) -- are the ADDRESSES free yet?
#
#     These groups are kept on-premises on purpose, so the AD object keeps
#     syncing and the cloud copy NEVER disappears -- it lingers holding an
#     @...onmicrosoft.com address. Asking "has the object gone?" therefore
#     answers no for ever. What Step 4 needs is the addresses, which are
#     released when the group is mail-disabled, and which come free earlier
#     than the object would have vanished.
#
# Read-only. Nothing is created or changed.
#
# Needs Exchange Online connected in this window.
# TOOLKIT-HELP-END
# Script 3 - Verify Sync Completion (Exchange Online)
# Input:  DGExport CSV file produced by Script 1
# Output: Confirms whether original DLs still exist in EXO
# Includes:
#   - Single combined instruction + file selection popup
#   - Auto-select most recent DGExport file
#   - 3-button dialog (OK / Other File / Cancel)

Add-Type -AssemblyName System.Windows.Forms

# Dot-sourced from two levels up: these live in Projects\DLtoSharedMbx, the
# toolkit root is AdminConsole. Brings in the theme, the audit log, and
# Invoke-EXOCommand.
. (Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'Common.ps1')

# The scripts live in the folder they work in, so $PSScriptRoot IS the base
# path -- same location as before, and correct on someone else's machine
# without editing anything. Was hard-coded to E:\scripts\... which only ever
# worked on one box.
$basePath = $PSScriptRoot

# ---------------------------------------------------------
# Auto-select most recent PRIMARY DGExport file (exclude compare files)
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
# Combined Instruction + File Selection Dialog (3-button)
# ---------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Script 3 - Verify Sync Completion"
$form.Width = 520
$form.Height = 300
$form.StartPosition = "CenterScreen"

$label = New-Object System.Windows.Forms.Label
$label.Text = @"
Before running Script 3, please confirm the following:

* You must be connected to Exchange Online PowerShell
* This script ONLY checks whether Azure AD Connect sync has completed
* No objects are created or modified by this script

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
$okButton.Top = 210
$okButton.Left = 120
$okButton.Add_Click({ $form.Tag = "OK"; $form.Close() })
$form.Controls.Add($okButton)

# Other File button
$otherButton = New-Object System.Windows.Forms.Button
$otherButton.Text = "Other File"
$otherButton.Width = 100
$otherButton.Top = 210
$otherButton.Left = 230
$otherButton.Add_Click({ $form.Tag = "Other"; $form.Close() })
$form.Controls.Add($otherButton)

# Cancel button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Width = 100
$cancelButton.Top = 210
$cancelButton.Left = 340
$cancelButton.Add_Click({ $form.Tag = "Cancel"; $form.Close() })
$form.Controls.Add($cancelButton)

# Theme + the standard '?' in the top-right, same as every other tool.
Set-ConsoleTheme -Form $form
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title '3-Cloud-CheckSync' | Out-Null

$form.ShowDialog() | Out-Null

switch ($form.Tag) {
    "Cancel" {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        return
    }
    "OK" {
        $csvFile = $autoFile
    }
    "Other" {
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.InitialDirectory = $basePath
        $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
        $openFileDialog.Title = "Select the desired DGExport file"

        if ($openFileDialog.ShowDialog() -ne "OK") {
            Write-Host "No file selected. Script cancelled." -ForegroundColor Yellow
            return
        }

        $csvFile = $openFileDialog.FileName
    }
}

Write-Host "Selected file: $csvFile"

# ---------------------------------------------------------
# Begin Script 3 Logic
# ---------------------------------------------------------

# Load CSV and ignore blank DGName rows
$dlList = Import-Csv $csvFile | Where-Object {
    $_.DGName -and $_.DGName.Trim() -ne ""
}

Write-Host ""
Write-Host "Checking sync status in Exchange Online..." -ForegroundColor Cyan
Write-Host ""

# Mirrors Script 4's Normalize-Domain exactly. This check has to ask the same
# question Step 4 will ask -- a gate that tests something different from the
# thing it guards is not a gate.
function Normalize-Domain {
    param($Address)
    if ($Address -match "@") {
        $local, $domain = $Address.Split("@")
        return "$local@ExpressPros.com"
    }
    return $Address
}

function Test-AddressFree {
    <#
        Is this SMTP address unheld in Exchange Online?

        Asked by address, not by object. Those are the same question for a
        deleted group and different questions for a kept one.
    #>
    param([string]$Address)
    $hit = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{
        Filter = "EmailAddresses -eq 'smtp:$Address'"; ErrorAction = 'SilentlyContinue'
    }
    return @{ Free = (-not $hit); HeldBy = $(if ($hit) { "$($hit[0].Name)" } else { '' }) }
}

$notReady = 0
$ready    = 0

foreach ($row in $dlList) {

    $dgName  = $row.DGName
    $newName = "$($row.NewName)".Trim()

    if ($newName -eq '') {
        # ---- Row will be DELETED on-prem by Step 2 -------------------------
        # Object-gone is the right test here: nothing keeps it alive, so its
        # disappearance from the cloud is exactly what "sync finished" means.
        Write-Host "Checking: $dgName  (to be deleted)"
        try {
            # Cloud, named explicitly. Get-Recipient exists on-premises too, and
            # this check is meaningless against the wrong directory -- it would
            # report "sync complete" for objects the cloud never had.
            $null = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{ Identity = $dgName; ErrorAction = 'Stop' }
            Write-Host "  STILL EXISTS in EXO -- Sync NOT complete" -ForegroundColor Yellow
            $notReady++
        }
        catch {
            Write-Host "  NOT FOUND in EXO -- Sync complete" -ForegroundColor Green
            $ready++
        }
        continue
    }

    # ---- PAYROLL row: kept on-prem, so the object NEVER goes away ----------
    #
    # Step 2 mail-disables these but deliberately does not delete them. The AD
    # group therefore keeps syncing and the cloud object stays forever, holding
    # an @...onmicrosoft.com address (KB 4096314).
    #
    # The old check asked "is the object gone?" and so answered STILL EXISTS on
    # every run, for ever, on exactly the lists that carry the listserv risk --
    # a permanent red light that had to be ignored, which is the same as having
    # no light at all.
    #
    # What Step 4 actually needs is the ADDRESS, because it re-adds the original
    # addresses as aliases on the new shared mailbox. Mail-disabling releases
    # the vanity address while leaving the object in place, so the address is
    # usually free well before the object would have vanished. Asking about the
    # address is both correct and earlier -- and earlier is less time bouncing.
    Write-Host "Checking: $dgName  (payroll -- kept on-prem)"

    $wanted = @()
    $wanted += "$newName@ExpressPros.com"                      # Step 4's new primary
    foreach ($e in ("$($row.Emails)" -split ';')) {             # Step 4's aliases
        if ($e -match '^(?i)smtp:') {
            $wanted += (Normalize-Domain ($e.Split(':')[1]))
        }
    }
    $wanted = @($wanted | Where-Object { $_ } | Select-Object -Unique)

    $blocked = @()
    foreach ($addr in $wanted) {
        $r = Test-AddressFree -Address $addr
        if ($r.Free) {
            Write-Host "  $addr -- free" -ForegroundColor Green
        } else {
            Write-Host "  $addr -- STILL HELD by '$($r.HeldBy)'" -ForegroundColor Yellow
            $blocked += $addr
        }
    }

    if ($blocked.Count -eq 0) {
        Write-Host "  READY -- every address Step 4 needs is free." -ForegroundColor Green
        $ready++
    } else {
        Write-Host "  NOT READY -- $($blocked.Count) address(es) still held. Wait for sync." -ForegroundColor Yellow
        $notReady++
    }
}

Write-Host ""
Write-Host "Ready: $ready    Not ready: $notReady"
if ($notReady -eq 0) {
    Write-Host "All clear -- run Step 4 NOW." -ForegroundColor Green
    Write-Host "Every minute between Step 2 and Step 4 is a minute these addresses bounce." -ForegroundColor Cyan
} else {
    Write-Host "Do not run Step 4 yet. Re-run this in a few minutes." -ForegroundColor Yellow
}
