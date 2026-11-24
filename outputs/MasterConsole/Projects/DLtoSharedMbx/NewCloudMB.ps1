# Script — Create Cloud Shared Mailbox + SMB-Access Group
# Uses NewCloudMBInput.csv (email column) + managers.csv for ManagedBy owners

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
# Unified Popup — Mailbox Name + Members Input
# ---------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Create Cloud Shared Mailbox"
$form.Width = 750
$form.Height = 650
$form.StartPosition = "CenterScreen"

$label = New-Object System.Windows.Forms.Label
$label.Text = @"
Before continuing, please confirm the following:

• You are connected to Exchange Online PowerShell
• This script will create a new Cloud shared mailbox
• It will also create the SMB-Access group for this mailbox
• Owners will be assigned from managers.csv
• Members will be added from your pasted list

Enter the mailbox name and member email addresses below.
"@
$label.AutoSize = $true
$label.Top = 10
$label.Left = 10
$form.Controls.Add($label)

# Mailbox Name
$nameLabel = New-Object System.Windows.Forms.Label
$nameLabel.Text = "Mailbox Name:"
$nameLabel.Top = 150
$nameLabel.Left = 10
$nameLabel.AutoSize = $true
$form.Controls.Add($nameLabel)

$nameBox = New-Object System.Windows.Forms.TextBox
$nameBox.Width = 400
$nameBox.Top = 170
$nameBox.Left = 10
$form.Controls.Add($nameBox)

# Members
$membersLabel = New-Object System.Windows.Forms.Label
$membersLabel.Text = "Members (paste anything containing .com addresses):"
$membersLabel.Top = 210
$membersLabel.Left = 10
$membersLabel.AutoSize = $true
$form.Controls.Add($membersLabel)

$membersBox = New-Object System.Windows.Forms.TextBox
$membersBox.Multiline = $true
$membersBox.ScrollBars = "Vertical"
$membersBox.Width = 700
$membersBox.Height = 300
$membersBox.Top = 240
$membersBox.Left = 10
$form.Controls.Add($membersBox)

# OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "BUILD CSV"
$okButton.Width = 120
$okButton.Top = 560
$okButton.Left = 450
$okButton.Add_Click({
    if ($nameBox.Text.Trim() -ne "") {
        $form.Tag = "OK"
        $form.Close()
    }
})
$form.Controls.Add($okButton)

# Cancel button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Width = 120
$cancelButton.Top = 560
$cancelButton.Left = 580
$cancelButton.Add_Click({
    $form.Tag = "Cancel"
    $form.Close()
})
$form.Controls.Add($cancelButton)

$form.ShowDialog() | Out-Null

if ($form.Tag -eq "Cancel") {
    Write-Host "Operation cancelled." -ForegroundColor Yellow
    return
}

$MailboxName = $nameBox.Text.Trim()
$Primary = "$MailboxName@ExpressPros.com"

# ---------------------------------------------------------
# Extract .com email addresses from pasted text
# ---------------------------------------------------------
$rawMembers = $membersBox.Text

# Split on anything that is NOT part of an email
$tokens = $rawMembers -split "[^A-Za-z0-9@\.\-_]"

$emails = foreach ($t in $tokens) {
    $clean = $t.Trim()
    if ($clean -match "\.com$") {
        $clean
    }
}

# Remove blanks + duplicates
$emails = $emails | Where-Object { $_ -ne "" } | Sort-Object -Unique

# ---------------------------------------------------------
# Build NewCloudMBInput.csv
# ---------------------------------------------------------
$inputFile = "$basePath\NewCloudMBInput.csv"

$emails | ForEach-Object {
    [PSCustomObject]@{
        email = $_
    }
} | Export-Csv -Path $inputFile -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "NewCloudMBInput.csv has been created:" -ForegroundColor Cyan
Get-Content $inputFile | Write-Host

Write-Host ""
Read-Host "Review the above CSV. Press ENTER to continue"

# ---------------------------------------------------------
# Load NewCloudMBInput.csv
# ---------------------------------------------------------
$MailboxMembers = Import-Csv $inputFile

Write-Host "Creating mailbox: $MailboxName" -ForegroundColor Cyan

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
# Create SMB-Access group
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
# Add members from NewCloudMBInput.csv
# ---------------------------------------------------------
foreach ($entry in $MailboxMembers) {

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
