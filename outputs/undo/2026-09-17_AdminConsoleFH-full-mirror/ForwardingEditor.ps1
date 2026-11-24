#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
<#
    ForwardingEditor.ps1
    Set, change or clear where a mailbox's incoming mail goes. Author : Glenn Armbruster

    Author   : Glenn Armbruster
    Revised  : 2026-08-27
#>
# TOOLKIT-HELP-RUNS: reads one mailbox's forwarding. Apply changes it and asks first.
# TOOLKIT-HELP-BEGIN
# "Forward Bob's mail to Alice while he's out."  "Stop forwarding it."
#
# Mailbox Health SHOWS forwarding. Returning User CLEARS it. This SETS it.
#
# THERE ARE TWO FORWARDING FIELDS AND THEY BEHAVE DIFFERENTLY:
#
#   Internal (ForwardingAddress)      a recipient in OUR directory. Honours
#                                     transport rules and is what you want for
#                                     one colleague covering another.
#   External (ForwardingSmtpAddress)  any address. This is the one that trips
#                                     the outbound-spam policy: Exchange Online
#                                     BLOCKS automatic forwarding to external
#                                     addresses by default, and a forward that
#                                     is set but blocked looks exactly like a
#                                     forward that works until somebody asks
#                                     where their mail went.
#
# This tool picks the field for you: an address that resolves in the directory
# goes in Internal, anything else goes in External and gets the warning.
#
# KEEP A COPY decides whether mail also stays in the original mailbox. Off
# means the mailbox stops receiving anything -- the usual mistake when the
# person comes back and finds three weeks of nothing.
#
# WHAT THIS DOES NOT TOUCH: inbox rules that forward. Those are per-user and
# live in Inbox Rules. A mailbox can be forwarding through a rule while this
# tool says it is not.
#
# Exchange Online only.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

#region UI --------------------------------------------------------------------

function Show-ForwardForm {
    param([string]$Seed)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Forwarding'
    $form.ClientSize = New-Object System.Drawing.Size(760, 400)
    $form.StartPosition = 'Manual'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $y = 12
    $lblMbx = New-Object System.Windows.Forms.Label
    $lblMbx.Text = 'Mailbox:'; $lblMbx.Location = New-Object System.Drawing.Point(12, ($y + 3)); $lblMbx.Size = New-Object System.Drawing.Size(70, 18)
    $form.Controls.Add($lblMbx)
    $txtMbx = New-Object System.Windows.Forms.TextBox
    $txtMbx.Location = New-Object System.Drawing.Point(84, $y); $txtMbx.Size = New-Object System.Drawing.Size(400, 24)
    $txtMbx.Font = New-Object System.Drawing.Font('Consolas', 9); $txtMbx.Text = "$Seed"
    $form.Controls.Add($txtMbx)
    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Text = 'Load'; $btnLoad.Location = New-Object System.Drawing.Point(496, ($y - 2)); $btnLoad.Size = New-Object System.Drawing.Size(110, 28)
    $form.Controls.Add($btnLoad)
    $y += 34

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(12, $y); $lblStatus.Size = New-Object System.Drawing.Size(736, 18)
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray; $lblStatus.Text = 'Enter a mailbox and press Load.'
    $form.Controls.Add($lblStatus)
    $y += 26

    $txtNow = New-Object System.Windows.Forms.TextBox
    $txtNow.Location = New-Object System.Drawing.Point(12, $y); $txtNow.Size = New-Object System.Drawing.Size(736, 110)
    $txtNow.Multiline = $true; $txtNow.ReadOnly = $true; $txtNow.WordWrap = $true; $txtNow.ScrollBars = 'Vertical'
    $txtNow.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtNow)
    $y += 120

    $grp = New-Object System.Windows.Forms.GroupBox
    $grp.Text = 'Set'; $grp.Location = New-Object System.Drawing.Point(12, $y); $grp.Size = New-Object System.Drawing.Size(736, 96)
    $form.Controls.Add($grp)
    $lblTo = New-Object System.Windows.Forms.Label
    $lblTo.Text = 'Forward to:'; $lblTo.Location = New-Object System.Drawing.Point(12, 26); $lblTo.Size = New-Object System.Drawing.Size(90, 18)
    $grp.Controls.Add($lblTo)
    $txtTo = New-Object System.Windows.Forms.TextBox
    $txtTo.Location = New-Object System.Drawing.Point(104, 22); $txtTo.Size = New-Object System.Drawing.Size(400, 24); $txtTo.Font = New-Object System.Drawing.Font('Consolas', 9)
    $grp.Controls.Add($txtTo)
    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Text = '(blank = clear forwarding)'; $lblHint.Location = New-Object System.Drawing.Point(512, 26); $lblHint.Size = New-Object System.Drawing.Size(200, 18); $lblHint.ForeColor = [System.Drawing.Color]::DimGray
    $grp.Controls.Add($lblHint)
    $chkKeep = New-Object System.Windows.Forms.CheckBox
    $chkKeep.Text = 'Keep a copy in this mailbox too  (DeliverToMailboxAndForward)'; $chkKeep.Checked = $true
    $chkKeep.Location = New-Object System.Drawing.Point(104, 56); $chkKeep.Size = New-Object System.Drawing.Size(460, 22)
    $grp.Controls.Add($chkKeep)
    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text = 'Apply'; $btnApply.Location = New-Object System.Drawing.Point(594, 52); $btnApply.Size = New-Object System.Drawing.Size(130, 28); $btnApply.Enabled = $false
    $grp.Controls.Add($btnApply)
    $y += 106

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Location = New-Object System.Drawing.Point(618, $y); $btnClose.Size = New-Object System.Drawing.Size(130, 28)
    $form.Controls.Add($btnClose); $btnClose.Add_Click({ $form.Close() })

    $form | Add-Member -NotePropertyName 'Busy' -NotePropertyValue $false -Force
    $form | Add-Member -NotePropertyName 'Mbx'  -NotePropertyValue $null  -Force

    $show = {
        $m = $form.Mbx
        $int = "$($m.ForwardingAddress)"; $ext = "$($m.ForwardingSmtpAddress)" -replace '^smtp:', ''
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("$($m.DisplayName)  <$($m.PrimarySmtpAddress)>")
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("  Internal forward  : $(if ($int) { $int } else { '(none)' })")
        [void]$sb.AppendLine("  External forward  : $(if ($ext) { $ext } else { '(none)' })")
        [void]$sb.AppendLine("  Keep a copy       : $($m.DeliverToMailboxAndForward)")
        if ($ext) { [void]$sb.AppendLine(''); [void]$sb.AppendLine('  EXTERNAL forwarding is set. Check the outbound spam policy allows it, or this mail is being dropped silently.') }
        $txtNow.Text = $sb.ToString()
        $txtTo.Text = if ($int) { $int } else { $ext }
        $chkKeep.Checked = [bool]$m.DeliverToMailboxAndForward
    }

    $btnLoad.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; $btnApply.Enabled = $false
            $s = Get-TargetSmtp -Text $txtMbx.Text
            if (-not $s) { $lblStatus.Text = 'Enter a mailbox address.'; return }
            $form.Mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $s; ErrorAction = 'Stop' }
            & $show
            $lblStatus.Text = 'Loaded.'; $btnApply.Enabled = $true
        } catch { $lblStatus.Text = 'Load failed -- see the error dialog.'; Show-HandlerError -ErrorRecord $_ -Context 'Load (Get-Mailbox)' }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    $btnApply.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $m = $form.Mbx; if (-not $m) { return }
            $to = "$($txtTo.Text)".Trim()
            $p = @{ Identity = "$($m.PrimarySmtpAddress)"; ErrorAction = 'Stop' }
            $desc = ''; $warn = ''

            if (-not $to) {
                # BOTH fields cleared. Clearing only the one that was set leaves
                # the other one -- and there are mailboxes with both.
                $p.ForwardingAddress = $null; $p.ForwardingSmtpAddress = $null; $p.DeliverToMailboxAndForward = $false
                $desc = 'CLEAR forwarding'
            } else {
                $rcp = $null
                try { $rcp = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{ Identity = $to; ErrorAction = 'Stop' } } catch { }
                if ($rcp) {
                    $p.ForwardingAddress = "$($rcp.PrimarySmtpAddress)"; $p.ForwardingSmtpAddress = $null
                    $desc = "forward to $($rcp.DisplayName) <$($rcp.PrimarySmtpAddress)> (internal)"
                } else {
                    if ($to -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { $lblStatus.Text = "'$to' is not in the directory and is not a valid address."; return }
                    $p.ForwardingSmtpAddress = $to; $p.ForwardingAddress = $null
                    $desc = "forward to $to (EXTERNAL)"
                    $warn = "`r`n`r`nEXTERNAL ADDRESS. Exchange Online blocks automatic external forwarding by default. Unless the outbound spam policy allows it for this mailbox, the forward will be set and the mail will be dropped silently."
                }
                $p.DeliverToMailboxAndForward = $chkKeep.Checked
                $desc += "; keep a copy = $($chkKeep.Checked)"
            }

            $a = Show-ToolDialog(("On $($m.PrimarySmtpAddress):`r`n`r`n  $desc" + $warn), 'Confirm', 'OKCancel', $(if ($warn) { 'Warning' } else { 'Question' }))
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters $p | Out-Null
            Write-CompletedLog -Action 'SetForwarding' -Target (Get-LogTarget $m.PrimarySmtpAddress) -Result 'Success' -Detail $desc
            $form.Mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = "$($m.PrimarySmtpAddress)"; ErrorAction = 'Stop' }
            & $show
            $lblStatus.Text = "Applied: $desc"
        } catch { $lblStatus.Text = 'Apply failed -- see the error dialog.'; Show-HandlerError -ErrorRecord $_ -Context 'Apply (Set-Mailbox forwarding)' }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    Set-ConsoleTheme -Form $form
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Forwarding' | Out-Null
    [void]$form.ShowDialog()
}
#endregion

Show-ForwardForm -Seed (Get-TargetSmtp -Text $SmtpAddress)
