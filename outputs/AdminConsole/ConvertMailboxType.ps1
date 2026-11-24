#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
<#
    ConvertMailboxType.ps1
    User mailbox to shared, or shared back to user, without touching anything else. Author : Glenn Armbruster

    Author   : Glenn Armbruster
    Revised  : 2026-08-27
#>
# TOOLKIT-HELP-RUNS: reads one mailbox's type. Convert changes it and asks first.
# TOOLKIT-HELP-BEGIN
# Disable-User converts a departing person's mailbox to shared as part of a
# whole offboarding. This does ONLY the conversion, for the two asks that are
# not an offboarding:
#
#   "They've left, keep the mailbox as shared for the team, don't disable
#    anything else."                                        -> To Shared
#
#   "We hired them back, make it a normal mailbox again."   -> To User
#
# THE LICENCE IS THE THING TO KNOW ABOUT.
#
#   To Shared   A shared mailbox under 50 GB needs NO licence. The person's
#               licence can be removed AFTER the conversion -- not before, or
#               the mailbox is deleted with them. This tool does not remove the
#               licence; it tells you when you can.
#
#   To User     A user mailbox MUST be licensed. Without one it works for a
#               30-day grace period and is then disabled. Assign the licence
#               first, then convert. This tool checks nothing about licences
#               because it cannot see them from Exchange alone -- it warns and
#               trusts you.
#
# Also handles Room and Equipment, because a shared mailbox that should have
# been a room is a recurring one.
#
# Exchange Online only. The conversion syncs to on-premises as an attribute
# change; nothing here writes to Active Directory.
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

function Show-ConvertForm {
    param([string]$Seed)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Convert Mailbox Type'
    $form.ClientSize = New-Object System.Drawing.Size(720, 360)
    $form.StartPosition = 'Manual'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $y = 12
    $lblMbx = New-Object System.Windows.Forms.Label
    $lblMbx.Text = 'Mailbox:'; $lblMbx.Location = New-Object System.Drawing.Point(12, ($y + 3)); $lblMbx.Size = New-Object System.Drawing.Size(70, 18)
    $form.Controls.Add($lblMbx)
    $txtMbx = New-Object System.Windows.Forms.TextBox
    $txtMbx.Location = New-Object System.Drawing.Point(84, $y); $txtMbx.Size = New-Object System.Drawing.Size(380, 24)
    $txtMbx.Font = New-Object System.Drawing.Font('Consolas', 9); $txtMbx.Text = "$Seed"
    $form.Controls.Add($txtMbx)
    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Text = 'Load'; $btnLoad.Location = New-Object System.Drawing.Point(476, ($y - 2)); $btnLoad.Size = New-Object System.Drawing.Size(110, 28)
    $form.Controls.Add($btnLoad)
    $y += 34

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(12, $y); $lblStatus.Size = New-Object System.Drawing.Size(696, 18)
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray; $lblStatus.Text = 'Enter a mailbox and press Load.'
    $form.Controls.Add($lblStatus)
    $y += 26

    $txtNow = New-Object System.Windows.Forms.TextBox
    $txtNow.Location = New-Object System.Drawing.Point(12, $y); $txtNow.Size = New-Object System.Drawing.Size(696, 140)
    $txtNow.Multiline = $true; $txtNow.ReadOnly = $true; $txtNow.WordWrap = $true; $txtNow.ScrollBars = 'Vertical'
    $txtNow.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtNow)
    $y += 150

    $lblTo = New-Object System.Windows.Forms.Label
    $lblTo.Text = 'Convert to:'; $lblTo.Location = New-Object System.Drawing.Point(12, ($y + 4)); $lblTo.Size = New-Object System.Drawing.Size(80, 18)
    $form.Controls.Add($lblTo)
    $cboTo = New-Object System.Windows.Forms.ComboBox
    $cboTo.DropDownStyle = 'DropDownList'
    # Learn 2026-08-27 (Set-Mailbox -Type): Regular | Room | Equipment | Shared.
    @('Shared', 'Regular', 'Room', 'Equipment') | ForEach-Object { [void]$cboTo.Items.Add($_) }
    $cboTo.SelectedIndex = 0
    $cboTo.Location = New-Object System.Drawing.Point(96, $y); $cboTo.Size = New-Object System.Drawing.Size(160, 24)
    $form.Controls.Add($cboTo)
    $btnGo = New-Object System.Windows.Forms.Button
    $btnGo.Text = 'Convert'; $btnGo.Location = New-Object System.Drawing.Point(272, ($y - 2)); $btnGo.Size = New-Object System.Drawing.Size(130, 28); $btnGo.Enabled = $false
    $form.Controls.Add($btnGo)
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Location = New-Object System.Drawing.Point(578, ($y - 2)); $btnClose.Size = New-Object System.Drawing.Size(130, 28)
    $form.Controls.Add($btnClose); $btnClose.Add_Click({ $form.Close() })

    $form | Add-Member -NotePropertyName 'Busy' -NotePropertyValue $false -Force
    $form | Add-Member -NotePropertyName 'Mbx'  -NotePropertyValue $null  -Force

    $load = {
        $s = Get-TargetSmtp -Text $txtMbx.Text
        if (-not $s) { $lblStatus.Text = 'Enter a mailbox address.'; return }
        $m = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $s; ErrorAction = 'Stop' }
        $form.Mbx = $m
        $st = $null
        try { $st = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = "$($m.PrimarySmtpAddress)"; ErrorAction = 'Stop' } } catch { }
        $size = if ($st) { "$($st.TotalItemSize)" } else { '(unknown)' }
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("$($m.DisplayName)  <$($m.PrimarySmtpAddress)>")
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("  Type now          : $($m.RecipientTypeDetails)")
        [void]$sb.AppendLine("  Size              : $size")
        [void]$sb.AppendLine("  Litigation hold   : $(if ($m.LitigationHoldEnabled) { 'ON' } else { 'off' })")
        [void]$sb.AppendLine("  Archive           : $(if ("$($m.ArchiveGuid)" -and "$($m.ArchiveGuid)" -ne '00000000-0000-0000-0000-000000000000') { 'yes' } else { 'no' })")
        [void]$sb.AppendLine("  Sign-in blocked   : $($m.AccountDisabled)")
        $txtNow.Text = $sb.ToString()
        $cur = "$($m.RecipientTypeDetails)"
        $idx = switch -Regex ($cur) { 'Shared' { 1 } 'Room' { 0 } 'Equipment' { 0 } default { 0 } }
        $cboTo.SelectedIndex = $idx
        $lblStatus.Text = 'Loaded.'
        $btnGo.Enabled = $true
    }

    $btnLoad.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try { $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; $btnGo.Enabled = $false; & $load }
        catch { $txtNow.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox'; $lblStatus.Text = 'Could not read the mailbox.' }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    $btnGo.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $m = $form.Mbx
            if (-not $m) { return }
            $to = "$($cboTo.SelectedItem)"
            $cur = "$($m.RecipientTypeDetails)"
            $same = ($to -eq 'Regular' -and $cur -eq 'UserMailbox') -or ($to -eq 'Shared' -and $cur -eq 'SharedMailbox') -or
                    ($to -eq 'Room' -and $cur -eq 'RoomMailbox') -or ($to -eq 'Equipment' -and $cur -eq 'EquipmentMailbox')
            if ($same) { $lblStatus.Text = "Already $cur. Nothing to do."; return }

            $warn = switch ($to) {
                'Regular' { "A USER mailbox must be licensed. If this person has no licence, it works for 30 days and is then disabled. Assign the licence FIRST." }
                'Shared'  { "A shared mailbox under 50 GB needs no licence. Remove the person's licence AFTER this completes -- removing it first deletes the mailbox with them." }
                default   { "Room and equipment mailboxes need no licence. Booking settings (Set-BookingPolicy) are separate and default to auto-accept off." }
            }
            $a = Show-ToolDialog(
                ("Convert $($m.PrimarySmtpAddress)" + "`r`n`r`n  from  $cur" + "`r`n  to    $to" + "`r`n`r`n" + $warn),
                'Confirm', 'OKCancel', 'Warning')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{ Identity = "$($m.PrimarySmtpAddress)"; Type = $to; ErrorAction = 'Stop' } | Out-Null
            Write-CompletedLog -Action 'ConvertMailboxType' -Target (Get-LogTarget $m.PrimarySmtpAddress) -Result 'Success' -Detail "$cur -> $to"
            & $load
            $lblStatus.Text = "Converted to $to. " + $(if ($to -eq 'Shared') { 'The licence can be removed now.' } elseif ($to -eq 'Regular') { 'Make sure a licence is assigned.' } else { '' })
        } catch {
            $txtNow.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Set-Mailbox -Type'
            $lblStatus.Text = 'Conversion failed.'
            Write-FailedLog -Action 'ConvertMailboxType' -Target "$($form.Mbx.PrimarySmtpAddress)" -ErrorRecord $_
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $form.Busy = $false
        }
    })

    Set-ConsoleTheme -Form $form
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Convert Mailbox Type' | Out-Null
    [void]$form.ShowDialog()
}
#endregion

Show-ConvertForm -Seed (Get-TargetSmtp -Text $SmtpAddress)
