#Requires -Version 5.1
<#
    MailboxPermissions.ps1
    Standalone Mailbox Permissions (Full Access / Send As delegation) tool
    -- Layer 2 tool script.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-BEGIN
# Grant, view or remove access to somebody else's mailbox.
#
# Two different things, often confused, both handled here:
#
#   * Full Access -- open and read the mailbox
#   * Send As -- send mail that appears to come from the mailbox
#
# They are granted separately. "I can see the shared mailbox but can't send from
# it" is Full Access without Send As, and is the most common ticket in this area.
#
# Changes are written to the audit log. Access granted here does not expire --
# remove it when it is no longer needed.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
#
# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost: the window flashes up, work may have started, and
# the operator has seen a tool they were not meant to be in.
if (-not (Test-ConsoleLaunch)) { return }

# NO PROMPT BEFORE THE WINDOW. The address box is ON the form -- see
# Add-ToolTargetBar in Common.ps1. Resolve-ToolTarget is not used any more:
# its job was to decide which DIRECTORY an address belongs to, and every read
# and write below goes through Invoke-EXOCommand to Exchange Online, which
# does not need to be told.
#
# An empty address is fine. The window opens and asks.
$script:Target = Get-TargetSmtp -Text $SmtpAddress

#region Actions - mailbox (moved unchanged from Mailbox-Group-Inspector.ps1 lines 671-795)
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
        Write-ToolStep 'Reading mailbox ...'
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
        Write-ToolStep 'Reading statistics ...'
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
        Write-ToolStep 'Reading Full Access ...'
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
        Write-ToolStep 'Reading Send As ...'
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

#region NEW - wrapper window for mailbox delegation (see .DESCRIPTION above)
function Show-MailboxPermissionsForm {
    # NO PARAMETERS. The mailbox can change while the window is open, so it
    # lives in $script:Target rather than in a parameter fixed at build time.
    # $Environment is still read below; it is the script-level parameter, and
    # it only ever labelled the audit line.
    $permForm = New-Object System.Windows.Forms.Form
    $permForm.Text = 'Mailbox Permissions'
    $permForm.ClientSize = New-Object System.Drawing.Size(764, 400)
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

    # SENT-ITEMS BEHAVIOUR, added 2026-08-27.
    #
    # Every time Send As is granted on a shared mailbox, the next ticket is
    # "my sent mail is in MY Sent Items, not the shared one". These two
    # booleans are the fix, and they belong on the same form as the grant so
    # the question gets asked at the same time as the permission is given.
    # Learn 2026-08-27 (Set-Mailbox): -MessageCopyForSentAsEnabled,
    # -MessageCopyForSendOnBehalfEnabled, both Boolean, Exchange Online.
    $chkCopySA = New-Object System.Windows.Forms.CheckBox
    $chkCopySA.Text = 'Send As copies to shared Sent Items'
    $chkCopySA.Location = New-Object System.Drawing.Point(124, 363)
    $chkCopySA.Size = New-Object System.Drawing.Size(230, 22)
    $permForm.Controls.Add($chkCopySA)

    $chkCopySOB = New-Object System.Windows.Forms.CheckBox
    $chkCopySOB.Text = 'Send on Behalf copies too'
    $chkCopySOB.Location = New-Object System.Drawing.Point(358, 363)
    $chkCopySOB.Size = New-Object System.Drawing.Size(180, 22)
    $permForm.Controls.Add($chkCopySOB)

    $btnCopies = New-Object System.Windows.Forms.Button
    $btnCopies.Text = 'Apply copies'
    $btnCopies.Location = New-Object System.Drawing.Point(544, 360)
    $btnCopies.Size = New-Object System.Drawing.Size(100, 28)
    $permForm.Controls.Add($btnCopies)

    function Update-StateDisplay {
        if (-not "$script:Target".Trim()) {
            $txtOut.Text = 'Enter a mailbox address above and press Look Up.'
            $permForm.Text = 'Mailbox Permissions'
            return
        }
        $permForm.Text = "Mailbox Permissions: $script:Target"
        try {
            $state = Get-MailboxState -Smtp $script:Target
            $txtOut.Text = Format-MailboxState -Heading "CURRENT STATE" -State $state
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mailbox Permissions'
        }
        # The two copy flags, read on their own so a failure here costs the
        # checkboxes and not the permission display above them.
        try {
            $mb = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $script:Target; ErrorAction = 'Stop' }
            $chkCopySA.Checked  = [bool]$mb.MessageCopyForSentAsEnabled
            $chkCopySOB.Checked = [bool]$mb.MessageCopyForSendOnBehalfEnabled
            $chkCopySA.Enabled = $true; $chkCopySOB.Enabled = $true; $btnCopies.Enabled = $true
        } catch {
            $chkCopySA.Enabled = $false; $chkCopySOB.Enabled = $false; $btnCopies.Enabled = $false
        }
    }

    $btnCopies.Add_Click({
        try {
            if (-not "$script:Target".Trim()) { return }
            $confirm = Show-ToolDialog(
                ("Set sent-items behaviour on '$script:Target'?" + "`r`n`r`n" +
                 "  Send As copies to shared Sent Items      : $($chkCopySA.Checked)" + "`r`n" +
                 "  Send on Behalf copies to shared Sent Items: $($chkCopySOB.Checked)"),
                'Confirm', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
                Identity                          = $script:Target
                MessageCopyForSentAsEnabled       = $chkCopySA.Checked
                MessageCopyForSendOnBehalfEnabled = $chkCopySOB.Checked
                ErrorAction                       = 'Stop' } | Out-Null
            Write-CompletedLog -Action 'SetSentItemsCopy' -Target (Get-LogTarget $script:Target) -Result 'Success' `
                -Detail "SentAs copy=$($chkCopySA.Checked); SendOnBehalf copy=$($chkCopySOB.Checked)"
            Update-StateDisplay
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Set-Mailbox (sent-items copies)'
            Write-FailedLog -Action 'SetSentItemsCopy' -Target $script:Target -ErrorRecord $_
        }
    })

    function Invoke-MailboxEdit {
        param([string]$ActionLabel, [string]$Verb, [scriptblock]$Action)
        $grantee = $txtMbxGrantee.Text.Trim()
        if (-not $grantee) { $txtOut.Text = "Enter a grantee (email or logon) first."; return }
        $confirm = Show-ToolDialog(
            "$ActionLabel for '$grantee' on mailbox '$script:Target'?",
            "Confirm $ActionLabel",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }
        $grpMbx.Enabled = $false
        try {
            $before = Get-MailboxState -Smtp $script:Target
            $out = New-Object System.Text.StringBuilder
            [void]$out.Append((Format-MailboxState -Heading "BEFORE" -State $before))
            & $Action
            [void]$out.AppendLine("$ActionLabel : SUCCESS ($grantee)")
            $after = Get-MailboxState -Smtp $script:Target
            [void]$out.Append((Format-MailboxState -Heading "AFTER" -State $after))
            $txtOut.Text = $out.ToString()
            # 'Cloud' rather than an empty pair of brackets.
            #
            # $Environment comes from Resolve-ToolTarget and is now NULL for a
            # cloud-only object, which -CloudOk made possible. Left as-is the
            # audit line read "(; grantee: ...)" -- a gap that looks like a bug
            # in the logger rather than the true statement "this object has no
            # on-premises directory".
            $envLabel = if ("$Environment".Trim()) { "$Environment" } else { 'Cloud' }
            Write-CompletedLog -Task "$Verb $script:Target ($envLabel; grantee: $grantee)"
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
            Grant-FullAccessPermission -Smtp $script:Target -Grantee $txtMbxGrantee.Text.Trim()
        }
    })
    $btnFARevoke.Add_Click({
        Invoke-MailboxEdit -ActionLabel "Revoke Full Access" -Verb "Revoked FullAccess" -Action {
            Revoke-FullAccessPermission -Smtp $script:Target -Grantee $txtMbxGrantee.Text.Trim()
        }
    })
    $btnSAGrant.Add_Click({
        Invoke-MailboxEdit -ActionLabel "Grant Send As" -Verb "Granted SendAs" -Action {
            Grant-SendAsPermission -Smtp $script:Target -Grantee $txtMbxGrantee.Text.Trim()
        }
    })
    $btnSARevoke.Add_Click({
        Invoke-MailboxEdit -ActionLabel "Revoke Send As" -Verb "Revoked SendAs" -Action {
            Revoke-SendAsPermission -Smtp $script:Target -Grantee $txtMbxGrantee.Text.Trim()
        }
    })
    $btnPermClose.Add_Click({ $permForm.Close() })

    # BEFORE Add-ToolHelpButton: the bar shifts every existing control down and
    # grows the form, and the help button places itself from ClientSize at run
    # time. Before Set-FormScreenCorner too, which positions from the size.
    $bar = Add-ToolTargetBar -Form $permForm -InitialValue $script:Target

    # WIRED HERE, NOT PASSED INTO THE HELPER. Update-StateDisplay is nested in
    # this function, so the handler has to be built in this scope to see it.
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        Update-StateDisplay
    })

    Set-ConsoleTheme -Form $permForm
    Set-FormScreenCorner -Form $permForm -Corner 'BottomLeft'
    # Consistent '?' in the top-right of every tool. Added after the theme so
    # it is not restyled, and BringToFront'd inside the helper because
    # WinForms draws EARLIER-added controls on top.
    Add-ToolHelpButton -Form $permForm -ScriptPath $PSCommandPath -Title 'Mailbox Permissions' | Out-Null

    [void]$permForm.ShowDialog()
}
#endregion

Show-MailboxPermissionsForm
