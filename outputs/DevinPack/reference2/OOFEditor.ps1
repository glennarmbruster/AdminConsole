#Requires -Version 5.1
<#
    OOFEditor.ps1
    Standalone Out of Office tool -- Layer 2 tool script.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-BEGIN
# Read and set someone's out-of-office message without needing their mailbox
# open.
#
# Shows whether auto-reply is on, the scheduled window if there is one, and both
# messages -- the internal one and the one external senders receive, which are
# often different and only one of which is usually wrong.
#
# For the classic ticket: somebody left, or is off sick, and their auto-reply
# needs setting or clearing on their behalf.
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
# Add-ToolTargetBar in Common.ps1. An empty address is fine; the window
# opens and asks for one.
$script:Target = Get-TargetSmtp -Text $SmtpAddress

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
    # NO -Smtp PARAMETER. The address can change while the window is open, so
    # it lives in $script:Target rather than a parameter fixed at build time.
    $oofForm = New-Object System.Windows.Forms.Form
    $oofForm.Text = "Out of Office: $script:Target"
    $oofForm.ClientSize = New-Object System.Drawing.Size(544, 514)
    $oofForm.StartPosition = 'Manual'
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
        # OPENED WITH NO ADDRESS IS A NORMAL STATE NOW -- the window appears
        # before anything has been typed, so the refresh must not run empty.
        if (-not "$script:Target".Trim()) {
            $oofForm.Text = 'Out of Office'
            return
        }
        $oofForm.Text = 'Out of Office: ' + $script:Target
        try {
            $cfg = Get-OutOfOfficeConfig -Smtp $script:Target
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
            "Save Out of Office settings for '$script:Target'? State: $($cboState.SelectedItem)",
            "Confirm Save", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnSave.Enabled = $false
        try {
            $st = $null; $et = $null
            if ($cboState.SelectedItem -eq 'Scheduled') {
                if ($txtStart.Text.Trim()) { $st = [DateTime]::Parse($txtStart.Text.Trim()) }
                if ($txtEnd.Text.Trim())   { $et = [DateTime]::Parse($txtEnd.Text.Trim()) }
            }
            Set-OutOfOfficeConfig -Smtp $script:Target -AutoReplyState $cboState.SelectedItem -ExternalAudience $cboAudience.SelectedItem `
                -InternalMessage $txtInt.Text -ExternalMessage $txtExt.Text -StartTime $st -EndTime $et
            Write-CompletedLog -Task "Set OutOfOffice $script:Target (EXO; state: $($cboState.SelectedItem))"
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
    $oofForm.StartPosition = 'Manual'
    Set-ConsoleTheme -Form $oofForm
    # Consistent '?' in the top-right of every tool. Added after the theme so
    # it is not restyled, and BringToFront'd inside the helper because
    # WinForms draws EARLIER-added controls on top.
    # BEFORE Add-ToolHelpButton: the bar shifts every existing control down
    # and grows the form, and the help button places itself from ClientSize at
    # run time.
    $bar = Add-ToolTargetBar -Form $oofForm -Label 'Mailbox or address:' -InitialValue $script:Target
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        Update-OofFormFromCurrent
    })

    Add-ToolHelpButton -Form $oofForm -ScriptPath $PSCommandPath -Title 'Out of Office' | Out-Null

    # AFTER the bar, not before. Add-ToolTargetBar grows the form by the
    # height of the address row, so a corner pinned first leaves the window
    # hanging that far past the edge of the working area -- which on the
    # bottom-left corner means underneath the taskbar.
    Set-FormScreenCorner -Form $oofForm -Corner 'BottomLeft'

    [void]$oofForm.ShowDialog()
}
#endregion

Show-OutOfOfficeForm
