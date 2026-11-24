#Requires -Version 5.1
<#
    SRGDisable.ps1
    SRG (cloud-only, srgpros.com) termination/disable tool -- hides and
    converts one or more mailboxes to shared, and optionally grants a
    manager Full Access (mailbox) and/or Site Collection Admin (OneDrive).

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

param(
    [string]$SmtpAddress
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'SRGCommon.ps1')

#region SRG-specific constants ---------------------------------------------------
$script:SpoTenantPrefix   = 'srgpros'
$script:OneDriveTenantHost = "$($script:SpoTenantPrefix)-my.sharepoint.com"
$script:SpoAdminUrl        = "https://$($script:SpoTenantPrefix)-admin.sharepoint.com"
#endregion

#region Parsing ----------------------------------------------------------------
function ConvertTo-NormalizedUserList {
    param([string]$Text)
    if (-not $Text) { return @() }
    $raw = @($Text -split '[\r\n,;]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
    $seen = New-Object System.Collections.Generic.HashSet[string]
    $result = @()
    foreach ($u in $raw) {
        if ($seen.Add($u.ToLowerInvariant())) { $result += $u }
    }
    return $result
}

function Get-OneDrivePersonalSiteUrl {
    param([string]$Smtp)
    $slug = ($Smtp -replace '\.', '_') -replace '@', '_'
    return "https://$script:OneDriveTenantHost/personal/$slug"
}
#endregion

#region Actions ------------------------------------------------------------------
function Disable-SrgUser {
    param(
        [string]$Smtp,
        [string]$ManagerSmtp,
        [bool]$GrantMailboxAccess,
        [bool]$GrantOneDriveAccess
    )
    $out = New-Object System.Text.StringBuilder
    [void]$out.AppendLine("=== $Smtp ===")

    $statusHideShared = 'FAILED'
    try {
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
            Identity = $Smtp; HiddenFromAddressListsEnabled = $true; Type = 'Shared'; ErrorAction = 'Stop'
        } | Out-Null
        $statusHideShared = 'OK'
        [void]$out.AppendLine("  Hide + convert to shared mailbox: OK")
    } catch {
        [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context 'Hide + convert to shared'))
    }

    $statusFullAccess = 'SKIPPED'
    if ($ManagerSmtp -and $GrantMailboxAccess) {
        try {
            Invoke-EXOCommand -Name 'Add-MailboxPermission' -Parameters @{
                Identity = $Smtp; User = $ManagerSmtp; AccessRights = 'FullAccess'; InheritanceType = 'All'; ErrorAction = 'Stop'
            } | Out-Null
            $statusFullAccess = 'OK'
            [void]$out.AppendLine("  Manager Full Access ($ManagerSmtp): OK")
        } catch {
            $statusFullAccess = 'FAILED'
            [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context 'Manager Full Access'))
        }
    } else {
        $reason = if (-not $ManagerSmtp) { 'no manager entered' } else { 'Manager Access set to OneDrive Only' }
        [void]$out.AppendLine("  Manager Full Access: SKIPPED ($reason)")
    }

    $statusOneDrive = 'SKIPPED'
    if ($ManagerSmtp -and $GrantOneDriveAccess) {
        $odUrl = Get-OneDrivePersonalSiteUrl -Smtp $Smtp
        try {
            Set-SPOUser -Site $odUrl -LoginName $ManagerSmtp -IsSiteCollectionAdmin $true -ErrorAction Stop | Out-Null
            $statusOneDrive = 'OK'
            [void]$out.AppendLine("  Manager OneDrive access ($ManagerSmtp -> $odUrl): OK")
        } catch {
            $statusOneDrive = 'FAILED'
            [void]$out.AppendLine((Format-ErrorDetail -ErrorRecord $_ -Context "Manager OneDrive access ($odUrl)"))
        }
    } else {
        $reason = if (-not $ManagerSmtp) { 'no manager entered' } else { 'Manager Access set to Mailbox Only' }
        [void]$out.AppendLine("  Manager OneDrive access: SKIPPED ($reason)")
    }

    [void]$out.AppendLine("  License removal: NOT ATTEMPTED -- handle manually (see completion popup).")
    [void]$out.AppendLine("  Sign-in block: NOT ATTEMPTED -- handle manually (see completion popup).")

    $managerLabel = if ($ManagerSmtp) { $ManagerSmtp } else { '(none)' }
    Write-CompletedLog -Task "Disabled $Smtp (manager: $managerLabel; HideShared=$statusHideShared; FullAccess=$statusFullAccess; OneDrive=$statusOneDrive; LicenseRemoved=NOT ATTEMPTED; SignInBlocked=NOT ATTEMPTED)"

    return $out.ToString()
}
#endregion

#region WinForms UI --------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "SRG Disable"
$form.Size = New-Object System.Drawing.Size(900, 574)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text = "Connect to Exchange Online and SharePoint Online (Connect-SPOService -Url $script:SpoAdminUrl) before running. License removal and sign-in block are NOT performed by this tool -- handle those manually (a popup will remind you which users when each run finishes)."
$lblWarn.Location = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size = New-Object System.Drawing.Size(860, 54)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

$lblUsers = New-Object System.Windows.Forms.Label
$lblUsers.Text = "User(s) to disable -- one per line (commas/semicolons also tolerated):"
$lblUsers.Location = New-Object System.Drawing.Point(12, 68)
$lblUsers.AutoSize = $true
$form.Controls.Add($lblUsers)
$txtUsers = New-Object System.Windows.Forms.TextBox
$txtUsers.Multiline = $true
$txtUsers.ScrollBars = "Vertical"
$txtUsers.Location = New-Object System.Drawing.Point(12, 88)
$txtUsers.Size = New-Object System.Drawing.Size(860, 110)
$txtUsers.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtUsers)

$lblManager = New-Object System.Windows.Forms.Label
$lblManager.Text = "Manager (optional):"
$lblManager.Location = New-Object System.Drawing.Point(12, 206)
$lblManager.AutoSize = $true
$form.Controls.Add($lblManager)
$txtManager = New-Object System.Windows.Forms.TextBox
$txtManager.Location = New-Object System.Drawing.Point(12, 226)
$txtManager.Size = New-Object System.Drawing.Size(260, 22)
$txtManager.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtManager)

$grpManagerAccess = New-Object System.Windows.Forms.GroupBox
$grpManagerAccess.Text = "Manager Access (only applies if a manager is entered above)"
$grpManagerAccess.Location = New-Object System.Drawing.Point(286, 214)
$grpManagerAccess.Size = New-Object System.Drawing.Size(586, 50)
$form.Controls.Add($grpManagerAccess)
$radMailboxOneDrive = New-Object System.Windows.Forms.RadioButton
$radMailboxOneDrive.Text = "Mailbox & OneDrive"
$radMailboxOneDrive.Location = New-Object System.Drawing.Point(10, 20)
$radMailboxOneDrive.AutoSize = $true
$radMailboxOneDrive.Checked = $true
$grpManagerAccess.Controls.Add($radMailboxOneDrive)
$radMailboxOnly = New-Object System.Windows.Forms.RadioButton
$radMailboxOnly.Text = "Mailbox Only"
$radMailboxOnly.Location = New-Object System.Drawing.Point(190, 20)
$radMailboxOnly.AutoSize = $true
$grpManagerAccess.Controls.Add($radMailboxOnly)
$radOneDriveOnly = New-Object System.Windows.Forms.RadioButton
$radOneDriveOnly.Text = "OneDrive Only"
$radOneDriveOnly.Location = New-Object System.Drawing.Point(340, 20)
$radOneDriveOnly.AutoSize = $true
$grpManagerAccess.Controls.Add($radOneDriveOnly)

function Update-ManagerAccessEnabled {
    $hasManager = -not [string]::IsNullOrWhiteSpace($txtManager.Text)
    $grpManagerAccess.Enabled = $hasManager
}
$txtManager.Add_TextChanged({ Update-ManagerAccessEnabled })

$btnDisable = New-Object System.Windows.Forms.Button
$btnDisable.Text = "Disable"
$btnDisable.Location = New-Object System.Drawing.Point(12, 286)
$btnDisable.Size = New-Object System.Drawing.Size(110, 30)
$btnDisable.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnDisable)
$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear"
$btnClear.Location = New-Object System.Drawing.Point(130, 286)
$btnClear.Size = New-Object System.Drawing.Size(90, 30)
$form.Controls.Add($btnClear)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(228, 286)
$btnClose.Size = New-Object System.Drawing.Size(90, 30)
$form.Controls.Add($btnClose)
$form.CancelButton = $btnClose

$lblResult = New-Object System.Windows.Forms.Label
$lblResult.Text = "Result:"
$lblResult.Location = New-Object System.Drawing.Point(12, 326)
$lblResult.Size = New-Object System.Drawing.Size(860, 20)
$lblResult.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblResult)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = "Vertical"
$txtOut.Location = New-Object System.Drawing.Point(12, 350)
$txtOut.Size = New-Object System.Drawing.Size(860, 140)
$txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtOut.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.Controls.Add($txtOut)

if ($SmtpAddress) { $txtUsers.Text = $SmtpAddress }
Update-ManagerAccessEnabled

$btnDisable.Add_Click({
    $btnDisable.Enabled = $false
    try {
        $users = ConvertTo-NormalizedUserList -Text $txtUsers.Text
        if ($users.Count -eq 0) { $txtOut.Text = "Enter at least one user to disable first."; return }
        $userSmtps = @($users | ForEach-Object { Get-TargetSmtp -Text $_ })

        $managerSmtp = $null
        $grantMailbox = $false
        $grantOneDrive = $false
        if (-not [string]::IsNullOrWhiteSpace($txtManager.Text)) {
            $managerSmtp = Get-TargetSmtp -Text $txtManager.Text
            $grantMailbox  = $radMailboxOneDrive.Checked -or $radMailboxOnly.Checked
            $grantOneDrive = $radMailboxOneDrive.Checked -or $radOneDriveOnly.Checked
        }

        $managerNote = if ($managerSmtp) {
            $accessBits = @()
            if ($grantMailbox)  { $accessBits += 'mailbox Full Access (no Send-As)' }
            if ($grantOneDrive) { $accessBits += 'OneDrive Site Collection Admin' }
            "Manager '$managerSmtp' will be granted: $($accessBits -join ' and ')."
        } else {
            "No manager entered -- no Full Access or OneDrive access will be granted to anyone."
        }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Disable $($userSmtps.Count) user(s)?`n`n$($userSmtps -join "`n")`n`nFor each: hide from address lists, convert to shared mailbox.`n`nNOTE: license removal and sign-in block are NOT performed by this tool -- handle those manually.`n`n$managerNote",
            "Confirm Disable ($($userSmtps.Count) user(s))",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }

        $out = New-Object System.Text.StringBuilder
        $n = 0
        foreach ($smtp in $userSmtps) {
            $n++
            $lblResult.Text = "Disabling $n of $($userSmtps.Count): $smtp ..."
            $form.Refresh()
            [void]$out.Append((Disable-SrgUser -Smtp $smtp -ManagerSmtp $managerSmtp -GrantMailboxAccess $grantMailbox -GrantOneDriveAccess $grantOneDrive))
            [void]$out.AppendLine("")
        }
        $lblResult.Text = "Done -- processed $($userSmtps.Count) user(s)."
        $txtOut.Text = $out.ToString()

        [System.Windows.Forms.MessageBox]::Show(
            "Complete:`n`nManually remove from license group and block sign-in for these users.`n`n$($userSmtps -join "`n")",
            "Manual Steps Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'SRG Disable'
    } finally {
        $btnDisable.Enabled = $true
    }
})

$btnClear.Add_Click({
    $txtUsers.Text = ''
    $txtManager.Text = ''
    $radMailboxOneDrive.Checked = $true
    Update-ManagerAccessEnabled
    $lblResult.Text = "Result:"
    $txtOut.Text = ''
})

$btnClose.Add_Click({ $form.Close() })

[void]$form.ShowDialog()
#endregion
