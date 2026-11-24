#Requires -Version 5.1
<#
.SYNOPSIS
    Standalone Calendar Permissions tool -- Layer 2 tool script.
.DESCRIPTION
    Fully functional on its own:
        .\CalendarPermissions.ps1
    prompts for mailbox / Environment / Domain Controller via Common.ps1's
    Show-TargetPromptDialog, then opens the Calendar Permissions window.

    Fully functional from AdminConsole.ps1 too:
        .\CalendarPermissions.ps1 -SmtpAddress user@domain.com -Environment Franchise -DomainController fran-vm-dc01...
    skips the prompt entirely and opens straight to the Calendar Permissions
    window for that mailbox.

    Get-CalendarFolderIdentity, Get-CalendarPermissions, Set-CalendarPermission,
    Remove-CalendarPermission, and Show-CalendarPermissionsForm below are
    MOVED, UNCHANGED, from the original monolithic Mailbox-Group-Inspector.ps1
    (original lines 1272-1455, region "Actions - calendar permissions"). The
    only edit anywhere in this file that touches original tool behavior is
    mechanical: Show-CalendarPermissionsForm's closing call changed from
        [void]$calForm.ShowDialog($form)
    to
        # ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin
        # Console, which opens top-left. Helper is in Common.ps1.
        Set-FormScreenCorner -Form $calForm -Corner 'BottomLeft'
        [void]$calForm.ShowDialog()
    because there is no longer a parent $form to own it -- this script IS
    the top-level window now. Every other line, including all comments, is
    byte-for-byte the original.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-10 (extracted from Mailbox-Group-Inspector.ps1 -- see
              REFACTOR-DESIGN.md)
#>
param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

$target = Resolve-ToolTarget -SmtpAddress $SmtpAddress -Environment $Environment -DomainController $DomainController -ToolTitle 'Calendar Permissions -- Enter Target'
if (-not $target) { return }   # operator cancelled the prompt

#region Actions - calendar permissions (moved unchanged from Mailbox-Group-Inspector.ps1 lines 1272-1455)
function Get-CalendarFolderIdentity {
    # Resolves the actual calendar folder rather than hardcoding ":\Calendar",
    # since the folder can be renamed or (rarely) localized. Falls back to the
    # first Calendar-scope folder if none matches the default root path.
    param([string]$Smtp)
    $folders = @(Invoke-EXOCommand -Name 'Get-MailboxFolderStatistics' -Parameters @{ Identity = $Smtp; FolderScope = 'Calendar'; ErrorAction = 'Stop' })
    if ($folders.Count -eq 0) { throw "No Calendar-scope folder found for $Smtp." }
    $primary = $folders | Where-Object { $_.FolderPath -eq '/Calendar' } | Select-Object -First 1
    if (-not $primary) { $primary = $folders | Select-Object -First 1 }
    return "$($Smtp):\$($primary.Name)"
}
function Get-CalendarPermissions {
    param([string]$Smtp)
    try {
        $calId = Get-CalendarFolderIdentity -Smtp $Smtp
        $perms = @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' -Parameters @{ Identity = $calId; ErrorAction = 'Stop' })
        return [PSCustomObject]@{ CalendarIdentity = $calId; Permissions = $perms; Error = $null }
    } catch {
        return [PSCustomObject]@{ CalendarIdentity = $null; Permissions = @(); Error = (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxFolderPermission (Calendar)') }
    }
}
function Set-CalendarPermission {
    # Add-MailboxFolderPermission fails if the user already has an entry, so
    # check first (best-effort) rather than parsing the resulting error text.
    # If the check itself errors for any reason, fall through to Add-; a
    # genuine problem will surface clearly from that call instead.
    param([string]$CalendarIdentity, [string]$Grantee, [string]$Role)
    $existing = @()
    try {
        $existing = @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' -Parameters @{ Identity = $CalendarIdentity; User = $Grantee; ErrorAction = 'Stop' })
    } catch {
        $existing = @()
    }
    if ($existing.Count -gt 0) {
        Invoke-EXOCommand -Name 'Set-MailboxFolderPermission' -Parameters @{
            Identity = $CalendarIdentity; User = $Grantee; AccessRights = $Role; Confirm = $false; ErrorAction = 'Stop'
        } | Out-Null
        return 'Updated'
    } else {
        Invoke-EXOCommand -Name 'Add-MailboxFolderPermission' -Parameters @{
            Identity = $CalendarIdentity; User = $Grantee; AccessRights = $Role; Confirm = $false; ErrorAction = 'Stop'
        } | Out-Null
        return 'Added'
    }
}
function Remove-CalendarPermission {
    # Remove-MailboxFolderPermission removes ALL permissions for this user on
    # the folder -- there's no selective/partial removal, per Microsoft Learn.
    param([string]$CalendarIdentity, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-MailboxFolderPermission' -Parameters @{
        Identity = $CalendarIdentity; User = $Grantee; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Show-CalendarPermissionsForm {
    param([string]$Smtp)
    $calForm = New-Object System.Windows.Forms.Form
    $calForm.Text = "Calendar Permissions: $Smtp"
    $calForm.Size = New-Object System.Drawing.Size(620, 480)
    $calForm.StartPosition = "CenterScreen"
    $calForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblCalId = New-Object System.Windows.Forms.Label
    $lblCalId.Text = "Calendar folder:"
    $lblCalId.Location = New-Object System.Drawing.Point(12, 12)
    $lblCalId.AutoSize = $true
    $calForm.Controls.Add($lblCalId)
    $grdCal = New-Object System.Windows.Forms.DataGridView
    $grdCal.Location = New-Object System.Drawing.Point(12, 34)
    $grdCal.Size = New-Object System.Drawing.Size(580, 220)
    $grdCal.ReadOnly = $true
    $grdCal.AllowUserToAddRows = $false
    $grdCal.AllowUserToDeleteRows = $false
    $grdCal.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdCal.MultiSelect = $false
    $grdCal.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grdCal.Font = New-Object System.Drawing.Font("Consolas", 9)
    [void]$grdCal.Columns.Add('User', 'User')
    [void]$grdCal.Columns.Add('AccessRights', 'Access Rights')
    foreach ($c in $grdCal.Columns) { $c.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable }
    $calForm.Controls.Add($grdCal)
    $lblGrantee = New-Object System.Windows.Forms.Label
    $lblGrantee.Text = "User (email or logon):"
    $lblGrantee.Location = New-Object System.Drawing.Point(12, 266)
    $lblGrantee.AutoSize = $true
    $calForm.Controls.Add($lblGrantee)
    $txtGrantee = New-Object System.Windows.Forms.TextBox
    $txtGrantee.Location = New-Object System.Drawing.Point(160, 263)
    $txtGrantee.Size = New-Object System.Drawing.Size(220, 20)
    $txtGrantee.Font = New-Object System.Drawing.Font("Consolas", 9)
    $calForm.Controls.Add($txtGrantee)
    $lblRole = New-Object System.Windows.Forms.Label
    $lblRole.Text = "Role:"
    $lblRole.Location = New-Object System.Drawing.Point(390, 266)
    $lblRole.AutoSize = $true
    $calForm.Controls.Add($lblRole)
    $cboRole = New-Object System.Windows.Forms.ComboBox
    $cboRole.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('Reviewer', 'Editor', 'Owner') | ForEach-Object { [void]$cboRole.Items.Add($_) }
    $cboRole.SelectedIndex = 0
    $cboRole.Location = New-Object System.Drawing.Point(430, 263)
    $cboRole.Size = New-Object System.Drawing.Size(120, 20)
    $calForm.Controls.Add($cboRole)
    $lblCalStatus = New-Object System.Windows.Forms.Label
    $lblCalStatus.Text = ""
    $lblCalStatus.Location = New-Object System.Drawing.Point(12, 296)
    $lblCalStatus.Size = New-Object System.Drawing.Size(580, 40)
    $calForm.Controls.Add($lblCalStatus)
    $btnAddPerm = New-Object System.Windows.Forms.Button
    $btnAddPerm.Text = "Add / Update Permission"
    $btnAddPerm.Location = New-Object System.Drawing.Point(12, 340)
    $btnAddPerm.Size = New-Object System.Drawing.Size(170, 28)
    $calForm.Controls.Add($btnAddPerm)
    $btnRemovePerm = New-Object System.Windows.Forms.Button
    $btnRemovePerm.Text = "Remove Selected"
    $btnRemovePerm.Location = New-Object System.Drawing.Point(190, 340)
    $btnRemovePerm.Size = New-Object System.Drawing.Size(150, 28)
    $calForm.Controls.Add($btnRemovePerm)
    $btnCalClose = New-Object System.Windows.Forms.Button
    $btnCalClose.Text = "Close"
    $btnCalClose.Location = New-Object System.Drawing.Point(500, 340)
    $btnCalClose.Size = New-Object System.Drawing.Size(90, 28)
    $calForm.Controls.Add($btnCalClose)
    $script:CalendarIdentityForForm = $null
    function Update-CalGridFromCurrent {
        $grdCal.Rows.Clear()
        $result = Get-CalendarPermissions -Smtp $Smtp
        if ($result.Error) {
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Red
            $lblCalStatus.Text = $result.Error
            $lblCalId.Text = "Calendar folder: (unresolved)"
            return
        }
        $script:CalendarIdentityForForm = $result.CalendarIdentity
        $lblCalId.Text = "Calendar folder: $($result.CalendarIdentity)"
        foreach ($p in $result.Permissions) {
            [void]$grdCal.Rows.Add($p.User.ToString(), ($p.AccessRights -join ', '))
        }
    }
    Update-CalGridFromCurrent
    $btnAddPerm.Add_Click({
        $grantee = $txtGrantee.Text.Trim()
        if (-not $grantee) { $lblCalStatus.Text = "Enter a user first."; return }
        if (-not $script:CalendarIdentityForForm) { $lblCalStatus.Text = "Calendar folder not resolved -- see status above."; return }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Grant '$($cboRole.SelectedItem)' calendar access to '$grantee' on '$Smtp'?",
            "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnAddPerm.Enabled = $false; $btnRemovePerm.Enabled = $false
        try {
            $action = Set-CalendarPermission -CalendarIdentity $script:CalendarIdentityForForm -Grantee $grantee -Role $cboRole.SelectedItem
            Write-CompletedLog -Task "$action CalendarPermission $Smtp (EXO; user: $grantee; role: $($cboRole.SelectedItem))"
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Black
            $lblCalStatus.Text = "$action permission for $grantee."
            Update-CalGridFromCurrent
        } catch {
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Red
            $lblCalStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Calendar permission grant'
        } finally {
            $btnAddPerm.Enabled = $true; $btnRemovePerm.Enabled = $true
        }
    })
    $btnRemovePerm.Add_Click({
        if ($grdCal.SelectedRows.Count -eq 0) { $lblCalStatus.Text = "Select a row first."; return }
        $grantee = $grdCal.SelectedRows[0].Cells['User'].Value
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Remove all calendar permissions for '$grantee' on '$Smtp'?",
            "Confirm Remove", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnAddPerm.Enabled = $false; $btnRemovePerm.Enabled = $false
        try {
            Remove-CalendarPermission -CalendarIdentity $script:CalendarIdentityForForm -Grantee $grantee
            Write-CompletedLog -Task "Removed CalendarPermission $Smtp (EXO; user: $grantee)"
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Black
            $lblCalStatus.Text = "Removed permission for $grantee."
            Update-CalGridFromCurrent
        } catch {
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Red
            $lblCalStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove-MailboxFolderPermission (Calendar)'
        } finally {
            $btnAddPerm.Enabled = $true; $btnRemovePerm.Enabled = $true
        }
    })
    $btnCalClose.Add_Click({ $calForm.Close() })
    # CHANGED FROM ORIGINAL: was "ShowDialog($form)" -- $form was the
    # monolith's shared parent window. This script has no parent window, so
    # it shows itself as the top-level dialog instead. Nothing else in this
    # function changed.
    [void]$calForm.ShowDialog()
}
#endregion

Show-CalendarPermissionsForm -Smtp $target.Smtp
