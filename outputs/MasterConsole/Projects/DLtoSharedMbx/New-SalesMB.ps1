# Script — Create Sales Shared Mailbox + SMB-Access Group
# Uses NewSalesInput.csv (email column) + managers.csv for ManagedBy owners
# Mirrors Script 4 behavior and naming conventions

Add-Type -AssemblyName System.Windows.Forms

$basePath = "E:\scripts\m365\dlconversion"

# ---------------------------------------------------------
# Load managers.csv
# ---------------------------------------------------------
$Managers = Import-Csv "$basePath\managers.csv"

$ResolvedManagers = @()
foreach ($m in $Managers.email) {
    $m = $m.Trim()
    if ($m -eq "") { continue }

    $Resolved = Get-Recipient -Identity $m -ErrorAction SilentlyContinue
    if ($Resolved) {
        $ResolvedManagers += $Resolved.PrimarySmtpAddress
    }
}

# ---------------------------------------------------------
# Popup dialog — instructions + mailbox name input
# ---------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Create Sales Shared Mailbox"
$form.Width = 520
$form.Height = 320
$form.StartPosition = "CenterScreen"

$label = New-Object System.Windows.Forms.Label
$label.Text = @"
Before continuing, please confirm the following:

• You are connected to Exchange Online PowerShell
• This script will create a new Sales shared mailbox
• It will also create the SMB-Access group for this mailbox
• Owners will be assigned from managers.csv
• Members will be added from NewSalesInput.csv

Enter the name of the new Sales mailbox (example: Sales.ChandlerAZ):
"@
$label.AutoSize = $true
$label.Top = 10
$label.Left = 10
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Width = 350
$textBox.Top = 160
$textBox.Left = 20
$form.Controls.Add($textBox)

# OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Width = 100
$okButton.Top = 210
$okButton.Left = 120
$okButton.Add_Click({
    if ($textBox.Text.Trim() -ne "") {
        $form.Tag = $textBox.Text.Trim()
        $form.Close()
    }
})
$form.Controls.Add($okButton)

# Cancel button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Width = 100
$cancelButton.Top = 210
$cancelButton.Left = 240
$cancelButton.Add_Click({
    $form.Tag = $null
    $form.Close()
})
$form.Controls.Add($cancelButton)

$form.ShowDialog() | Out-Null

if (-not $form.Tag) {
    Write-Host "Operation cancelled." -ForegroundColor Yellow
    return
}

$MailboxName = $form.Tag
$Primary = "$MailboxName@ExpressPros.com"

Write-Host "Creating mailbox: $MailboxName" -ForegroundColor Cyan

# ---------------------------------------------------------
# Load NewSalesInput.csv
# ---------------------------------------------------------
$InputFile = "$basePath\NewSalesInput.csv"

if (-not (Test-Path $InputFile)) {
    Write-Host "Input file not found: $InputFile" -ForegroundColor Red
    return
}

$SalesMembers = Import-Csv $InputFile

# ---------------------------------------------------------
# Create shared mailbox
# ---------------------------------------------------------
if (-not (Get-Mailbox -Identity $MailboxName -ErrorAction SilentlyContinue)) {

    New-Mailbox -Shared -Name $MailboxName -PrimarySmtpAddress $Primary

    # Sent Items settings
    Set-Mailbox -Identity $MailboxName -MessageCopyForSentAsEnabled $true
    Set-Mailbox -Identity $MailboxName -MessageCopyForSendOnBehalfEnabled $true
}

# ---------------------------------------------------------
# Create SMB-Access group (naming: SMB-Access-MailboxName)
# ---------------------------------------------------------
$AccessGroup = "SMB-Access-$MailboxName"

if (-not (Get-DistributionGroup -Identity $AccessGroup -ErrorAction SilentlyContinue)) {

    New-DistributionGroup `
        -Name $AccessGroup `
        -Alias $AccessGroup `
        -PrimarySmtpAddress "$AccessGroup@ExpressPros.com" `
        -Type Security `
        -ErrorAction Stop

    Set-DistributionGroup -Identity $AccessGroup -Description "."
}

# ---------------------------------------------------------
# Assign permissions
# ---------------------------------------------------------
Add-MailboxPermission -Identity $MailboxName -User $AccessGroup -AccessRights FullAccess -AutoMapping:$false -Confirm:$false
Add-RecipientPermission -Identity $MailboxName -Trustee $AccessGroup -AccessRights SendAs -Confirm:$false

# ---------------------------------------------------------
# Assign owners from managers.csv
# ---------------------------------------------------------
if ($ResolvedManagers.Count -gt 0) {
    Set-DistributionGroup -Identity $AccessGroup -ManagedBy $ResolvedManagers -BypassSecurityGroupManagerCheck
}
else {
    Write-Host "No managers resolved from managers.csv; skipping ManagedBy assignment." -ForegroundColor Yellow
}

# ---------------------------------------------------------
# Add members from NewSalesInput.csv
# ---------------------------------------------------------
foreach ($entry in $SalesMembers) {

    $email = $entry.email
    if ($null -eq $email) { continue }
    $email = $email.Trim()
    if ($email -eq "") { continue }

    $Resolved = Get-Recipient -Identity $email -ErrorAction SilentlyContinue
    if (-not $Resolved) {
        $Resolved = Get-Recipient -Filter "PrimarySmtpAddress -eq '$email'" -ErrorAction SilentlyContinue
    }
    if (-not $Resolved) {
        $Resolved = Get-Recipient -Filter "EmailAddresses -eq 'smtp:$email'" -ErrorAction SilentlyContinue
    }

    if ($Resolved) {
        Write-Host "Adding $($Resolved.PrimarySmtpAddress) to $AccessGroup"
        Add-DistributionGroupMember -Identity $AccessGroup -Member $Resolved.PrimarySmtpAddress -BypassSecurityGroupManagerCheck -ErrorAction SilentlyContinue
    }
    else {
        Write-Host "Could not resolve $email; skipping." -ForegroundColor Yellow
    }
}
