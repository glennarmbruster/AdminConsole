# Script 4A — Create New Cloud-Only Shared Mailboxes + SMB-Access Groups
# No DGExport required. User enters 1–6 mailbox names manually.
# Uses managers.csv for ManagedBy owners.

Add-Type -AssemblyName System.Windows.Forms

$basePath = "E:\scripts\m365\dlconversion"
$managersFile = Join-Path $basePath "managers.csv"

# ---------------------------------------------------------
# Load managers.csv and resolve recipients
# ---------------------------------------------------------
if (-not (Test-Path $managersFile)) {
    Write-Host "ERROR: managers.csv not found at $managersFile" -ForegroundColor Red
    return
}

$Managers = Import-Csv $managersFile

$ResolvedManagers = @()
foreach ($m in $Managers.email) {
    if ($m -eq "") { continue }

    $Resolved = Get-Recipient -Identity $m -ErrorAction SilentlyContinue
    if ($Resolved) {
        $ResolvedManagers += $Resolved.PrimarySmtpAddress
    }
}

# ---------------------------------------------------------
# Build Input Dialog (6 mailbox names)
# ---------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Create New Cloud Mailboxes"
$form.Width = 520
$form.Height = 500     # Taller dialog box
$form.StartPosition = "CenterScreen"

$label = New-Object System.Windows.Forms.Label
$label.Text = @"
This script will create NEW shared mailboxes directly in Exchange Online.

For each mailbox entered:
• A shared mailbox will be created
• Sent Items settings will be enabled
• An SMB-Access-<MailboxName> security group will be created
• FullAccess and SendAs permissions will be assigned
• Owners will be assigned from managers.csv
• No members will be added to the SMB-Access group

Enter 1–6 mailbox names exactly as you want them created.
"@
$label.AutoSize = $true
$label.Top = 10
$label.Left = 10
$form.Controls.Add($label)

# Textboxes
$textBoxes = @()
$top = 180
for ($i = 1; $i -le 6; $i++) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "MB${i}:"
    $lbl.Top = $top
    $lbl.Left = 20
    $lbl.Width = 40
    $form.Controls.Add($lbl)

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Top = $top - 3
    $tb.Left = 70
    $tb.Width = 400
    $form.Controls.Add($tb)

    $textBoxes += $tb
    $top += 35
}

# OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Width = 100
$okButton.Top = 420
$okButton.Left = 200
$okButton.Add_Click({ $form.Tag = "OK"; $form.Close() })
$form.Controls.Add($okButton)

# Cancel button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Width = 100
$cancelButton.Top = 420
$cancelButton.Left = 310
$cancelButton.Add_Click({ $form.Tag = "Cancel"; $form.Close() })
$form.Controls.Add($cancelButton)

$form.ShowDialog() | Out-Null

if ($form.Tag -eq "Cancel") {
    Write-Host "Operation cancelled by user." -ForegroundColor Yellow
    return
}

# ---------------------------------------------------------
# Collect mailbox names
# ---------------------------------------------------------
$MailboxNames = @()
foreach ($tb in $textBoxes) {
    if ($tb.Text -ne "") {
        $MailboxNames += $tb.Text
    }
}

if ($MailboxNames.Count -eq 0) {
    Write-Host "No mailbox names entered. Exiting." -ForegroundColor Yellow
    return
}

Write-Host ""
Write-Host "Mailboxes to create:" -ForegroundColor Cyan
$MailboxNames | ForEach-Object { Write-Host "  $_" }
Write-Host ""

# ---------------------------------------------------------
# Process each mailbox
# ---------------------------------------------------------
foreach ($MailboxName in $MailboxNames) {

    Write-Host ""
    Write-Host "---------------------------------------------"
    Write-Host "Processing mailbox: $MailboxName"
    Write-Host "---------------------------------------------"

    $Primary = "$MailboxName@ExpressPros.com"

    # ---------------------------------------------------------
    # Create shared mailbox (if not exists)
    # ---------------------------------------------------------
    $existingMB = Get-Mailbox -Identity $MailboxName -ErrorAction SilentlyContinue

    if (-not $existingMB) {
        Write-Host "Creating shared mailbox: $MailboxName"
        New-Mailbox -Shared -Name $MailboxName -PrimarySmtpAddress $Primary

        # Sent Items settings
        Set-Mailbox -Identity $MailboxName -MessageCopyForSentAsEnabled $true
        Set-Mailbox -Identity $MailboxName -MessageCopyForSendOnBehalfEnabled $true
    }
    else {
        Write-Host "Shared mailbox already exists — skipping creation."
        Write-Host "Ensuring Sent Items settings are applied..."
        Set-Mailbox -Identity $MailboxName -MessageCopyForSentAsEnabled $true
        Set-Mailbox -Identity $MailboxName -MessageCopyForSendOnBehalfEnabled $true
    }

    # ---------------------------------------------------------
    # Create SMB-Access group
    # ---------------------------------------------------------
    $AccessGroup = "SMB-Access-$MailboxName"
    $existingDG = Get-DistributionGroup -Identity $AccessGroup -ErrorAction SilentlyContinue

    if (-not $existingDG) {
        Write-Host "Creating SMB-Access group: $AccessGroup"

        New-DistributionGroup `
            -Name $AccessGroup `
            -Alias $AccessGroup `
            -PrimarySmtpAddress "$AccessGroup@ExpressPros.com" `
            -Type Security `
            -ErrorAction Stop

        Set-DistributionGroup -Identity $AccessGroup -Description "."
    }
    else {
        Write-Host "SMB-Access group already exists — skipping creation."
    }

    # ---------------------------------------------------------
    # Delay to allow AAD DualWrite to complete
    # ---------------------------------------------------------
    Write-Host "Waiting 15 seconds for Azure AD to finish provisioning..." -ForegroundColor Yellow
    Start-Sleep -Seconds 15

    # ---------------------------------------------------------
    # Assign permissions
    # ---------------------------------------------------------
    Write-Host "Assigning FullAccess and SendAs permissions..."

    Add-MailboxPermission -Identity $MailboxName -User $AccessGroup -AccessRights FullAccess -AutoMapping:$false -Confirm:$false
    Add-RecipientPermission -Identity $MailboxName -Trustee $AccessGroup -AccessRights SendAs -Confirm:$false

    # ---------------------------------------------------------
    # Assign owners
    # ---------------------------------------------------------
    Write-Host "Assigning owners from managers.csv..."
    Set-DistributionGroup -Identity $AccessGroup -ManagedBy $ResolvedManagers -BypassSecurityGroupManagerCheck

    Write-Host "Completed: $MailboxName" -ForegroundColor Green
}

Write-Host ""
Write-Host "All requested mailboxes processed." -ForegroundColor Green
