#Requires -Version 5.1
# TOOLKIT-HELP-BEGIN
# STEP 4 of the DL cloud migration. Builds the replacement for each list that
# was decommissioned on-premises.
#
# For every row in the DGExport file it creates:
#
#   * a shared mailbox -- named from NewName if there is one, otherwise the
#     original name -- carrying the old list's addresses as aliases
#   * a mail-enabled security group, SMB-Access-<name>, holding the old list's
#     members
#   * Full Access and Send As for that group over the mailbox
#   * owners on the group, taken from Managers.csv
#
# Sent items are configured to copy to the mailbox for both Send As and Send on
# Behalf, so replies do not vanish into an individual's Sent Items.
#
# Safe to re-run: a mailbox or group that already exists is left alone rather
# than recreated, so a partial run can be finished by running it again.
#
# Needs Exchange Online connected in this window.
# TOOLKIT-HELP-END
# Script 4 -- Create Shared Mailboxes + SMB-Access Groups (Exchange Online)
# Uses DGExport file from Script 1 and managers.csv for ManagedBy owners

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
$form.Text = "Script 4 - Create Shared Mailboxes"
$form.Width = 520
$form.Height = 300
$form.StartPosition = "CenterScreen"

$label = New-Object System.Windows.Forms.Label
$label.Text = @"
Before running Script 4, please confirm the following:

* You must be connected to Exchange Online PowerShell
* This script creates shared mailboxes and SMB-Access groups
* It assigns permissions and owners from managers.csv

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
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title '4-Cloud-Create' | Out-Null

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
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.InitialDirectory = $basePath
        $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
        $openFileDialog.Title = "Select the desired DGExport file"

        if ($openFileDialog.ShowDialog() -ne "OK") {
            Write-Host "No file selected. Script cancelled." -ForegroundColor Yellow
            return
        }

        $fullPath = $openFileDialog.FileName
    }
}

Write-Host "Using DGExport file: $fullPath"

# ---------------------------------------------------------
# Begin Script 4 Logic
# ---------------------------------------------------------

$CSV = Import-Csv $fullPath
$Managers = Import-Csv "$basePath\managers.csv"

# Resolve managers
$ResolvedManagers = @()
foreach ($m in $Managers.email) {
    $m = $m.Trim()
    if ($m -eq "") { continue }

    $Resolved = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{ Identity = $m; ErrorAction = 'SilentlyContinue' }
    if ($Resolved) {
        $ResolvedManagers += $Resolved.PrimarySmtpAddress
    }
}

function Normalize-Domain {
    param($Address)
    if ($Address -match "@") {
        $local, $domain = $Address.Split("@")
        return "$local@ExpressPros.com"
    }
    return $Address
}

foreach ($row in $CSV) {

    $OriginalName = $row.DGName
    $NewName      = $row.NewName

    # Determine mailbox name
    if ($NewName -ne "") {
        $MailboxName = $NewName
    }
    else {
        $MailboxName = $OriginalName
    }

    # Extract all SMTP addresses (ignore X500)
    $EmailsRaw = $row.Emails -split ";"
    $SMTPs = $EmailsRaw | Where-Object { $_ -match "^smtp:" -or $_ -match "^SMTP:" }

    # Normalize domain + strip "smtp:" prefix
    $SMTPs = $SMTPs | ForEach-Object {
        $addr = $_.Split(":")[1]
        Normalize-Domain $addr
    }

    # Determine primary SMTP
    if ($NewName -ne "") {
        $Primary = "$NewName@ExpressPros.com"
    }
    else {
        $Primary = ($SMTPs | Where-Object { $_ -cmatch "@" } | Select-Object -First 1)
    }

    # ---------------------------------------------------------
    # SAFER MAILBOX CREATION PATTERN
    # ---------------------------------------------------------
    if (-not (Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $MailboxName; ErrorAction = 'SilentlyContinue' })) {
        Invoke-EXOCommand -Name 'New-Mailbox' -Parameters @{ Shared = $true; Name = $MailboxName; PrimarySmtpAddress = $Primary }

        # Apply Sent Items settings AFTER creation (works in all modules)
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{ Identity = $MailboxName; MessageCopyForSentAsEnabled = $true }
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{ Identity = $MailboxName; MessageCopyForSendOnBehalfEnabled = $true }
    }

    # Add aliases (all SMTPs except primary)
    foreach ($alias in $SMTPs) {
        if ($alias -ne $Primary) {
            Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{ Identity = $MailboxName; EmailAddresses = @{Add=$alias} }
        }
    }

    # Create access group (direct mail-enabled security)
    $AccessGroup = "SMB-Access-$MailboxName"
    if (-not (Invoke-EXOCommand -Name 'Get-DistributionGroup' -Parameters @{ Identity = $AccessGroup; ErrorAction = 'SilentlyContinue' })) {

        Invoke-EXOCommand -Name 'New-DistributionGroup' -Parameters @{
            Name               = $AccessGroup
            Alias              = $AccessGroup
            PrimarySmtpAddress = "$AccessGroup@ExpressPros.com"
            Type               = 'Security'
            ErrorAction        = 'Stop'
        }

        Invoke-EXOCommand -Name 'Set-DistributionGroup' -Parameters @{ Identity = $AccessGroup; Description = '.' }
    }

    # Assign permissions (NO PROMPTS)
    Invoke-EXOCommand -Name 'Add-MailboxPermission' -Parameters @{ Identity = $MailboxName; User = $AccessGroup; AccessRights = 'FullAccess'; AutoMapping = $false; Confirm = $false }
    Invoke-EXOCommand -Name 'Add-RecipientPermission' -Parameters @{ Identity = $MailboxName; Trustee = $AccessGroup; AccessRights = 'SendAs'; Confirm = $false }

    # Assign owners from managers.csv
    Invoke-EXOCommand -Name 'Set-DistributionGroup' -Parameters @{ Identity = $AccessGroup; ManagedBy = $ResolvedManagers; BypassSecurityGroupManagerCheck = $true }

    # Populate access group
    $Members = $row.Members -split ";"

    foreach ($m in $Members) {
        $m = $m.Trim()
        if ($m -eq "") { continue }

        # Resolve user
        $Resolved = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{ Identity = $m; ErrorAction = 'SilentlyContinue' }
        if (-not $Resolved) {
            $Resolved = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{ Filter = "PrimarySmtpAddress -eq '$m'"; ErrorAction = 'SilentlyContinue' }
        }
        if (-not $Resolved) {
            $Resolved = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{ Filter = "EmailAddresses -eq 'smtp:$m'"; ErrorAction = 'SilentlyContinue' }
        }

        if ($Resolved) {
            Write-Host "Adding $($Resolved.PrimarySmtpAddress) to group $AccessGroup"
            Invoke-EXOCommand -Name 'Add-DistributionGroupMember' -Parameters @{ Identity = $AccessGroup; Member = $Resolved.PrimarySmtpAddress; BypassSecurityGroupManagerCheck = $true; ErrorAction = 'SilentlyContinue' }
        }
    }
}
