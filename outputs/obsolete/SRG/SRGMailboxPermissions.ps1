#Requires -Version 5.1
<#
    SRGMailboxPermissions.ps1
    SRG (cloud-only, srgpros.com) Mailbox Permissions (Full Access / Send As
    delegation) tool.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

param(
    [string]$SmtpAddress
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'SRGCommon.ps1')

$target = Resolve-ToolTarget -SmtpAddress $SmtpAddress -ToolTitle 'SRG Mailbox Permissions -- Enter Target'
if (-not $target) { return }

#region Actions - mailbox (ported unchanged from MailboxPermissions.ps1 -- EXO only)
function Get-MailboxState {
    param([string]$Smtp)
    $state = [ordered]@{
        DisplayName          = $null
        PrimarySmtpAddress   = $null
        Alias                = $null
        RecipientTypeDetails = $null
        Hidden               = $null
        LitigationHold       = $null
        ArchiveEnabled       = $null
        ArchiveStatus        = $null
        MailboxSize          = $null
        MailboxItemCount     = $null
        ArchiveSize          = $null
        ArchiveItemCount     = $null
        FullAccess           = @()
        SendAs               = @()
        Errors               = @()
    }
    try {
        $mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $state.DisplayName          = $mbx.DisplayName
        $state.PrimarySmtpAddress   = $mbx.PrimarySmtpAddress
        $state.Alias                = $mbx.Alias
        $state.RecipientTypeDetails = $mbx.RecipientTypeDetails
        $state.Hidden               = $mbx.HiddenFromAddressListsEnabled
        $state.LitigationHold       = $mbx.LitigationHoldEnabled
        $state.ArchiveEnabled = ($mbx.ArchiveGuid -and $mbx.ArchiveGuid -ne [Guid]::Empty)
        $state.ArchiveStatus  = $mbx.ArchiveStatus
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox (core properties)')
    }
    try {
        $stats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $state.MailboxSize      = $stats.TotalItemSize
        $state.MailboxItemCount = $stats.ItemCount
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics (mailbox size)')
    }
    if ($state.ArchiveEnabled) {
        try {
            $astats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; Archive = $true; ErrorAction = 'Stop' }
            $state.ArchiveSize      = $astats.TotalItemSize
            $state.ArchiveItemCount = $astats.ItemCount
        } catch {
            $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics -Archive (archive size)')
        }
    }
    try {
        $fa = Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object {
                $_.User -notlike 'NT AUTHORITY\SELF' -and
                $_.IsInherited -eq $false -and
                $_.Deny -eq $false -and
                ($_.AccessRights -contains 'FullAccess')
            }
        $state.FullAccess = @($fa | ForEach-Object { $_.User.ToString() })
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxPermission (Full Access)')
    }
    try {
        $sa = Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object {
                $_.Trustee -notlike 'NT AUTHORITY\SELF' -and
                $_.AccessControlType -eq 'Allow'
            }
        $state.SendAs = @($sa | ForEach-Object { $_.Trustee.ToString() })
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-RecipientPermission (Send As)')
    }
    return [PSCustomObject]$state
}
function Format-MailboxState {
    param([string]$Heading, $State)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("--- $Heading (Mailbox) ---")
    [void]$sb.AppendLine("  Display Name        : $($State.DisplayName)")
    [void]$sb.AppendLine("  Primary SMTP        : $($State.PrimarySmtpAddress)")
    [void]$sb.AppendLine("  Alias               : $($State.Alias)")
    [void]$sb.AppendLine("  Mailbox Type        : $($State.RecipientTypeDetails)")
    [void]$sb.AppendLine("  Hidden From GAL     : $($State.Hidden)")
    [void]$sb.AppendLine("  Litigation Hold     : $($State.LitigationHold)")
    [void]$sb.AppendLine("  Mailbox Size        : $($State.MailboxSize)  ($($State.MailboxItemCount) items)")
    [void]$sb.AppendLine("  Archive Enabled     : $($State.ArchiveEnabled)  (ArchiveStatus: $($State.ArchiveStatus))")
    if ($State.ArchiveEnabled) {
        [void]$sb.AppendLine("  Archive Size        : $($State.ArchiveSize)  ($($State.ArchiveItemCount) items)")
    }
    [void]$sb.AppendLine("  Full Access grantees:")
    if ($State.FullAccess.Count -eq 0) { [void]$sb.AppendLine("    (none)") }
    else { $State.FullAccess | ForEach-Object { [void]$sb.AppendLine("    $_") } }
    [void]$sb.AppendLine("  Send As grantees    :")
    if ($State.SendAs.Count -eq 0) { [void]$sb.AppendLine("    (none)") }
    else { $State.SendAs | ForEach-Object { [void]$sb.AppendLine("    $_") } }
    if ($State.Errors.Count -gt 0) {
        [void]$sb.AppendLine("  --- Field errors (partial data above is still valid) ---")
        $State.Errors | ForEach-Object { [void]$sb.AppendLine("  $_") }
    }
    return $sb.ToString()
}
function Grant-FullAccessPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Add-MailboxPermission' -Parameters @{
        Identity = $Smtp; User = $Grantee; AccessRights = 'FullAccess'; InheritanceType = 'All'; ErrorAction = 'Stop'
    } | Out-Null
}
function Revoke-FullAccessPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-MailboxPermission' -Parameters @{
        Identity = $Smtp; User = $Grantee; AccessRights = 'FullAccess'; InheritanceType = 'All'; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Grant-SendAsPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Add-RecipientPermission' -Parameters @{
        Identity = $Smtp; Trustee = $Grantee; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Revoke-SendAsPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-RecipientPermission' -Parameters @{
        Identity = $Smtp; Trustee = $Grantee; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
#endregion

#region Wrapper window for mailbox delegation (see .DESCRIPTION above) -----------
function Show-MailboxPermissionsForm {
    param([string]$Smtp)
    $permForm = New-Object System.Windows.Forms.Form
    $permForm.Text = "SRG Mailbox Permissions: $Smtp"
    $permForm.Size = New-Object System.Drawing.Size(780, 420)
    $permForm.StartPosition = "CenterScreen"
    $permForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Multiline = $true
    $txtOut.ReadOnly = $true
    $txtOut.ScrollBars = "Vertical"
    $txtOut.Location = New-Object System.Drawing.Point(12, 12)
    $txtOut.Size = New-Object System.Drawing.Size(740, 240)
    $txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
    $permForm.Controls.Add($txtOut)

    $grpMbx = New-Object System.Windows.Forms.GroupBox
    $grpMbx.Text = "Mailbox Permissions"
    $grpMbx.Location = New-Object System.Drawing.Point(12, 262)
    $grpMbx.Size = New-Object System.Drawing.Size(740, 90)
    $permForm.Controls.Add($grpMbx)
    $lblMbxGrantee = New-Object System.Windows.Forms.Label
    $lblMbxGrantee.Text = "Grantee (email or logon):"
    $lblMbxGrantee.Location = New-Object System.Drawing.Point(10, 22)
    $lblMbxGrantee.AutoSize = $true
    $grpMbx.Controls.Add($lblMbxGrantee)
    $txtMbxGrantee = New-Object System.Windows.Forms.TextBox
    $txtMbxGrantee.Location = New-Object System.Drawing.Point(170, 19)
    $txtMbxGrantee.Size = New-Object System.Drawing.Size(260, 22)
    $txtMbxGrantee.Font = New-Object System.Drawing.Font("Consolas", 9)
    $grpMbx.Controls.Add($txtMbxGrantee)
    $btnFAGrant = New-Object System.Windows.Forms.Button
    $btnFAGrant.Text = "Grant Full Access"
    $btnFAGrant.Location = New-Object System.Drawing.Point(10, 54)
    $btnFAGrant.Size = New-Object System.Drawing.Size(140, 28)
    $grpMbx.Controls.Add($btnFAGrant)
    $btnFARevoke = New-Object System.Windows.Forms.Button
    $btnFARevoke.Text = "Revoke Full Access"
    $btnFARevoke.Location = New-Object System.Drawing.Point(156, 54)
    $btnFARevoke.Size = New-Object System.Drawing.Size(140, 28)
    $grpMbx.Controls.Add($btnFARevoke)
    $btnSAGrant = New-Object System.Windows.Forms.Button
    $btnSAGrant.Text = "Grant Send As"
    $btnSAGrant.Location = New-Object System.Drawing.Point(302, 54)
    $btnSAGrant.Size = New-Object System.Drawing.Size(140, 28)
    $grpMbx.Controls.Add($btnSAGrant)
    $btnSARevoke = New-Object System.Windows.Forms.Button
    $btnSARevoke.Text = "Revoke Send As"
    $btnSARevoke.Location = New-Object System.Drawing.Point(448, 54)
    $btnSARevoke.Size = New-Object System.Drawing.Size(140, 28)
    $grpMbx.Controls.Add($btnSARevoke)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh"
    $btnRefresh.Location = New-Object System.Drawing.Point(12, 360)
    $btnRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $permForm.Controls.Add($btnRefresh)
    $btnPermClose = New-Object System.Windows.Forms.Button
    $btnPermClose.Text = "Close"
    $btnPermClose.Location = New-Object System.Drawing.Point(652, 360)
    $btnPermClose.Size = New-Object System.Drawing.Size(100, 28)
    $permForm.Controls.Add($btnPermClose)

    function Update-StateDisplay {
        try {
            $state = Get-MailboxState -Smtp $Smtp
            $txtOut.Text = Format-MailboxState -Heading "CURRENT STATE" -State $state
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mailbox Permissions'
        }
    }

    function Invoke-MailboxEdit {
        param([string]$ActionLabel, [string]$Verb, [scriptblock]$Action)
        $grantee = $txtMbxGrantee.Text.Trim()
        if (-not $grantee) { $txtOut.Text = "Enter a grantee (email or logon) first."; return }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "$ActionLabel for '$grantee' on mailbox '$Smtp'?",
            "Confirm $ActionLabel",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }
        $grpMbx.Enabled = $false
        try {
            $before = Get-MailboxState -Smtp $Smtp
            $out = New-Object System.Text.StringBuilder
            [void]$out.Append((Format-MailboxState -Heading "BEFORE" -State $before))
            & $Action
            [void]$out.AppendLine("$ActionLabel : SUCCESS ($grantee)")
            $after = Get-MailboxState -Smtp $Smtp
            [void]$out.Append((Format-MailboxState -Heading "AFTER" -State $after))
            $txtOut.Text = $out.ToString()
            Write-CompletedLog -Task "$Verb $Smtp (grantee: $grantee)"
        }
        catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context $ActionLabel
        }
        finally {
            $grpMbx.Enabled = $true
        }
    }

    Update-StateDisplay
    $btnRefresh.Add_Click({ Update-StateDisplay })
    $btnFAGrant.Add_Click({
        Invoke-MailboxEdit -ActionLabel "Grant Full Access" -Verb "Granted FullAccess" -Action {
            Grant-FullAccessPermission -Smtp $Smtp -Grantee $txtMbxGrantee.Text.Trim()
        }
    })
    $btnFARevoke.Add_Click({
        Invoke-MailboxEdit -ActionLabel "Revoke Full Access" -Verb "Revoked FullAccess" -Action {
            Revoke-FullAccessPermission -Smtp $Smtp -Grantee $txtMbxGrantee.Text.Trim()
        }
    })
    $btnSAGrant.Add_Click({
        Invoke-MailboxEdit -ActionLabel "Grant Send As" -Verb "Granted SendAs" -Action {
            Grant-SendAsPermission -Smtp $Smtp -Grantee $txtMbxGrantee.Text.Trim()
        }
    })
    $btnSARevoke.Add_Click({
        Invoke-MailboxEdit -ActionLabel "Revoke Send As" -Verb "Revoked SendAs" -Action {
            Revoke-SendAsPermission -Smtp $Smtp -Grantee $txtMbxGrantee.Text.Trim()
        }
    })
    $btnPermClose.Add_Click({ $permForm.Close() })
    [void]$permForm.ShowDialog()
}
#endregion

Show-MailboxPermissionsForm -Smtp $target.Smtp
