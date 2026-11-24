#Requires -Version 5.1
<#
.SYNOPSIS
    Standalone Out of Office tool -- Layer 2 tool script.
.DESCRIPTION
    Fully functional on its own:
        .\OOFEditor.ps1
    prompts for mailbox / Environment / Domain Controller via Common.ps1's
    Show-TargetPromptDialog, then opens the Out of Office window.

    Fully functional from AdminConsole.ps1 too:
        .\OOFEditor.ps1 -SmtpAddress user@domain.com -Environment Franchise -DomainController fran-vm-dc01...
    skips the prompt entirely and opens straight to the Out of Office window
    for that mailbox.

    Get-OutOfOfficeConfig, Set-OutOfOfficeConfig, and Show-OutOfOfficeForm
    below are MOVED, UNCHANGED, from the original monolithic
    Mailbox-Group-Inspector.ps1 (original lines 1108-1269, region "Actions -
    out of office"). The only edit anywhere in this file that touches
    original tool behavior is mechanical: Show-OutOfOfficeForm's closing call
    changed from
        [void]$oofForm.ShowDialog($form)
    to
        # ADDED 2026-07-24 -- open bottom-left so a tool never covers the Admin
        # Console, which opens top-left. Helper is in Common.ps1.
        Set-FormScreenCorner -Form $oofForm -Corner 'BottomLeft'
        [void]$oofForm.ShowDialog()
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

$target = Resolve-ToolTarget -SmtpAddress $SmtpAddress -Environment $Environment -DomainController $DomainController -ToolTitle 'Out of Office -- Enter Target'
if (-not $target) { return }   # operator cancelled the prompt

#region Actions - out of office (moved unchanged from Mailbox-Group-Inspector.ps1 lines 1108-1269)
function Get-OutOfOfficeConfig {
    param([string]$Smtp)
    Invoke-EXOCommand -Name 'Get-MailboxAutoReplyConfiguration' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
}
function Set-OutOfOfficeConfig {
    param(
        [string]$Smtp,
        [string]$AutoReplyState,
        [string]$ExternalAudience,
        [string]$InternalMessage,
        [string]$ExternalMessage,
        [Nullable[DateTime]]$StartTime,
        [Nullable[DateTime]]$EndTime
    )
    $params = @{
        Identity         = $Smtp
        AutoReplyState   = $AutoReplyState
        ExternalAudience = $ExternalAudience
        InternalMessage  = $InternalMessage
        ExternalMessage  = $ExternalMessage
        Confirm          = $false
        ErrorAction      = 'Stop'
    }
    if ($AutoReplyState -eq 'Scheduled') {
        if ($StartTime) { $params.StartTime = $StartTime }
        if ($EndTime)   { $params.EndTime   = $EndTime }
    }
    Invoke-EXOCommand -Name 'Set-MailboxAutoReplyConfiguration' -Parameters $params | Out-Null
}
function Show-OutOfOfficeForm {
    param([string]$Smtp)
    $oofForm = New-Object System.Windows.Forms.Form
    $oofForm.Text = "Out of Office: $Smtp"
    $oofForm.Size = New-Object System.Drawing.Size(560, 560)
    $oofForm.StartPosition = "CenterScreen"
    $oofForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblState = New-Object System.Windows.Forms.Label
    $lblState.Text = "State:"
    $lblState.Location = New-Object System.Drawing.Point(12, 15)
    $lblState.AutoSize = $true
    $oofForm.Controls.Add($lblState)
    $cboState = New-Object System.Windows.Forms.ComboBox
    $cboState.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('Disabled', 'Enabled', 'Scheduled') | ForEach-Object { [void]$cboState.Items.Add($_) }
    $cboState.Location = New-Object System.Drawing.Point(140, 12)
    $cboState.Size = New-Object System.Drawing.Size(150, 20)
    $oofForm.Controls.Add($cboState)
    $lblAudience = New-Object System.Windows.Forms.Label
    $lblAudience.Text = "External audience:"
    $lblAudience.Location = New-Object System.Drawing.Point(300, 15)
    $lblAudience.AutoSize = $true
    $oofForm.Controls.Add($lblAudience)
    $cboAudience = New-Object System.Windows.Forms.ComboBox
    $cboAudience.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('None', 'Known', 'All') | ForEach-Object { [void]$cboAudience.Items.Add($_) }
    $cboAudience.Location = New-Object System.Drawing.Point(420, 12)
    $cboAudience.Size = New-Object System.Drawing.Size(100, 20)
    $oofForm.Controls.Add($cboAudience)
    $lblStart = New-Object System.Windows.Forms.Label
    $lblStart.Text = "Start (Scheduled only, e.g. 2026-07-10 08:00):"
    $lblStart.Location = New-Object System.Drawing.Point(12, 45)
    $lblStart.AutoSize = $true
    $oofForm.Controls.Add($lblStart)
    $txtStart = New-Object System.Windows.Forms.TextBox
    $txtStart.Location = New-Object System.Drawing.Point(280, 42)
    $txtStart.Size = New-Object System.Drawing.Size(240, 20)
    $txtStart.Font = New-Object System.Drawing.Font("Consolas", 9)
    $oofForm.Controls.Add($txtStart)
    $lblEnd = New-Object System.Windows.Forms.Label
    $lblEnd.Text = "End (Scheduled only):"
    $lblEnd.Location = New-Object System.Drawing.Point(12, 70)
    $lblEnd.AutoSize = $true
    $oofForm.Controls.Add($lblEnd)
    $txtEnd = New-Object System.Windows.Forms.TextBox
    $txtEnd.Location = New-Object System.Drawing.Point(280, 67)
    $txtEnd.Size = New-Object System.Drawing.Size(240, 20)
    $txtEnd.Font = New-Object System.Drawing.Font("Consolas", 9)
    $oofForm.Controls.Add($txtEnd)
    $lblInt = New-Object System.Windows.Forms.Label
    $lblInt.Text = "Internal message:"
    $lblInt.Location = New-Object System.Drawing.Point(12, 98)
    $lblInt.AutoSize = $true
    $oofForm.Controls.Add($lblInt)
    $txtInt = New-Object System.Windows.Forms.TextBox
    $txtInt.Multiline = $true
    $txtInt.ScrollBars = "Vertical"
    $txtInt.Location = New-Object System.Drawing.Point(12, 118)
    $txtInt.Size = New-Object System.Drawing.Size(520, 140)
    $txtInt.Font = New-Object System.Drawing.Font("Consolas", 9)
    $oofForm.Controls.Add($txtInt)
    $lblExt = New-Object System.Windows.Forms.Label
    $lblExt.Text = "External message:"
    $lblExt.Location = New-Object System.Drawing.Point(12, 264)
    $lblExt.AutoSize = $true
    $oofForm.Controls.Add($lblExt)
    $txtExt = New-Object System.Windows.Forms.TextBox
    $txtExt.Multiline = $true
    $txtExt.ScrollBars = "Vertical"
    $txtExt.Location = New-Object System.Drawing.Point(12, 284)
    $txtExt.Size = New-Object System.Drawing.Size(520, 140)
    $txtExt.Font = New-Object System.Drawing.Font("Consolas", 9)
    $oofForm.Controls.Add($txtExt)
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = ""
    $lblStatus.Location = New-Object System.Drawing.Point(12, 430)
    $lblStatus.Size = New-Object System.Drawing.Size(520, 40)
    $oofForm.Controls.Add($lblStatus)
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Save"
    $btnSave.Location = New-Object System.Drawing.Point(12, 474)
    $btnSave.Size = New-Object System.Drawing.Size(100, 28)
    $oofForm.Controls.Add($btnSave)
    $btnOofClose = New-Object System.Windows.Forms.Button
    $btnOofClose.Text = "Close"
    $btnOofClose.Location = New-Object System.Drawing.Point(120, 474)
    $btnOofClose.Size = New-Object System.Drawing.Size(100, 28)
    $oofForm.Controls.Add($btnOofClose)
    function Update-OofFormFromCurrent {
        try {
            $cfg = Get-OutOfOfficeConfig -Smtp $Smtp
            $cboState.SelectedItem    = $cfg.AutoReplyState.ToString()
            $cboAudience.SelectedItem = $cfg.ExternalAudience.ToString()
            $txtStart.Text = if ($cfg.StartTime) { $cfg.StartTime.ToString() } else { '' }
            $txtEnd.Text   = if ($cfg.EndTime)   { $cfg.EndTime.ToString() }   else { '' }
            $txtInt.Text   = $cfg.InternalMessage
            $txtExt.Text   = $cfg.ExternalMessage
            $lblStatus.ForeColor = [System.Drawing.Color]::Black
            $lblStatus.Text = "Loaded current settings."
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxAutoReplyConfiguration'
        }
    }
    Update-OofFormFromCurrent
    $btnSave.Add_Click({
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Save Out of Office settings for '$Smtp'? State: $($cboState.SelectedItem)",
            "Confirm Save", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnSave.Enabled = $false
        try {
            $st = $null; $et = $null
            if ($cboState.SelectedItem -eq 'Scheduled') {
                if ($txtStart.Text.Trim()) { $st = [DateTime]::Parse($txtStart.Text.Trim()) }
                if ($txtEnd.Text.Trim())   { $et = [DateTime]::Parse($txtEnd.Text.Trim()) }
            }
            Set-OutOfOfficeConfig -Smtp $Smtp -AutoReplyState $cboState.SelectedItem -ExternalAudience $cboAudience.SelectedItem `
                -InternalMessage $txtInt.Text -ExternalMessage $txtExt.Text -StartTime $st -EndTime $et
            Write-CompletedLog -Task "Set OutOfOffice $Smtp (EXO; state: $($cboState.SelectedItem))"
            $lblStatus.ForeColor = [System.Drawing.Color]::Black
            $lblStatus.Text = "Saved successfully."
            Update-OofFormFromCurrent
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Set-MailboxAutoReplyConfiguration'
        } finally {
            $btnSave.Enabled = $true
        }
    })
    $btnOofClose.Add_Click({ $oofForm.Close() })
    # CHANGED FROM ORIGINAL: was "ShowDialog($form)" -- $form was the
    # monolith's shared parent window. This script has no parent window, so
    # it shows itself as the top-level dialog instead. Nothing else in this
    # function changed.
    [void]$oofForm.ShowDialog()
}
#endregion

Show-OutOfOfficeForm -Smtp $target.Smtp
