#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
<#
    MobileDevices.ps1
    Every phone and tablet synced to a mailbox: what it is, when it last synced, and the two ways to cut it off. Author : Glenn Armbruster

    Author   : Glenn Armbruster
    Revised  : 2026-08-27
#>
# TOOLKIT-HELP-RUNS: lists the devices on one mailbox. Remove and Wipe change things and each asks first.
# TOOLKIT-HELP-BEGIN
# Phones and tablets that sync with a mailbox through ActiveSync or Outlook
# mobile, with model, OS, last sync time and status.
#
# TWO DIFFERENT ACTIONS, AND THE DIFFERENCE MATTERS:
#
#   Remove partnership   Exchange forgets the device. It stops syncing. The mail
#                        already on the phone STAYS on the phone. Use this for
#                        an old phone the user replaced, or a device that
#                        should never have been there.
#
#   Wipe (account only)  Exchange tells the device to delete THIS ACCOUNT's
#                        mail, contacts and calendar the next time it checks in.
#                        Personal photos and other apps are untouched. Use this
#                        for a departing employee's personal phone.
#
# A FULL device wipe -- factory reset of everything -- is deliberately NOT
# offered here. It is the right answer for a lost company-owned phone and the
# wrong answer for anything else, and this tool cannot tell which it is looking
# at. If you need it, that is a one-off script and a conversation first.
#
# A wipe is QUEUED, not immediate. It runs when the device next connects, which
# for a phone that has been switched off is never. The grid shows the wipe
# status so you can see whether it has actually happened.
#
# Departing-employee note: Disable-User does not do this. Run this first.
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

#region Data ------------------------------------------------------------------

function Get-MailboxDevice {
    <#
        Learn 2026-08-27 (Get-MobileDeviceStatistics): -Mailbox <identity>.
        Returns DeviceId, DeviceType, DeviceModel, DeviceOS, DeviceFriendlyName,
        LastSuccessSync, DeviceAccessState, Status, IsRemoteWipeSupported,
        DeviceWipeRequestTime, DeviceWipeSentTime, DeviceWipeAckTime, Identity.

        The statistics cmdlet rather than Get-MobileDevice, because the
        statistics carry the last-sync time and the wipe state, and those are
        the two facts that decide what to do.
    #>
    param([Parameter(Mandatory)][string]$Smtp)
    return @(Invoke-EXOCommand -Name 'Get-MobileDeviceStatistics' -Parameters @{
        Mailbox = $Smtp; ErrorAction = 'Stop'
    })
}

function Remove-MailboxDevice {
    # Learn 2026-08-27 (Remove-MobileDevice): -Identity, -Confirm.
    param([Parameter(Mandatory)][string]$Identity)
    Invoke-EXOCommand -Name 'Remove-MobileDevice' -Parameters @{
        Identity = $Identity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}

function Clear-MailboxDeviceAccount {
    <#
        Learn 2026-08-27 (Clear-MobileDevice): -Identity, -AccountOnly,
        -Confirm. -AccountOnly removes only this account's data -- mail,
        contacts, calendar -- and leaves the rest of the device alone.

        The full-wipe form of this cmdlet (no -AccountOnly) is intentionally
        not reachable from this tool. See the help block.
    #>
    param([Parameter(Mandatory)][string]$Identity)
    Invoke-EXOCommand -Name 'Clear-MobileDevice' -Parameters @{
        Identity = $Identity; AccountOnly = $true; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
#endregion

#region UI --------------------------------------------------------------------

function Show-DeviceForm {
    param([string]$Seed)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Mobile Devices'
    $form.ClientSize = New-Object System.Drawing.Size(1060, 560)
    $form.StartPosition = 'Manual'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $y = 12
    $lblMbx = New-Object System.Windows.Forms.Label
    $lblMbx.Text = 'Mailbox:'
    $lblMbx.Location = New-Object System.Drawing.Point(12, ($y + 3))
    $lblMbx.Size = New-Object System.Drawing.Size(70, 18)
    $form.Controls.Add($lblMbx)

    $txtMbx = New-Object System.Windows.Forms.TextBox
    $txtMbx.Location = New-Object System.Drawing.Point(84, $y)
    $txtMbx.Size = New-Object System.Drawing.Size(360, 24)
    $txtMbx.Font = New-Object System.Drawing.Font('Consolas', 9)
    $txtMbx.Text = "$Seed"
    $form.Controls.Add($txtMbx)

    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Text = 'Load'
    $btnLoad.Location = New-Object System.Drawing.Point(456, ($y - 2))
    $btnLoad.Size = New-Object System.Drawing.Size(110, 28)
    $form.Controls.Add($btnLoad)
    $y += 34

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(12, $y)
    $lblStatus.Size = New-Object System.Drawing.Size(1036, 18)
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
    $lblStatus.Text = 'Enter a mailbox and press Load.'
    $form.Controls.Add($lblStatus)
    $y += 22

    $pb = Add-ReportProgressBar -Form $form -X 12 -Y $y -Width 1036 -Height 14
    $y += 24

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(12, $y)
    $grid.Size = New-Object System.Drawing.Size(1036, 380)
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $true
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.RowHeadersVisible = $false
    $grid.Font = New-Object System.Drawing.Font('Consolas', 9)
    foreach ($c in @(@('Device', 18), @('Model', 16), @('OS', 16), @('Type', 10), @('LastSync', 14), @('Access', 10), @('Wipe', 16))) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c[0]; $col.HeaderText = $c[0]; $col.FillWeight = $c[1]
        [void]$grid.Columns.Add($col)
    }
    $colId = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colId.Name = 'Identity'; $colId.Visible = $false
    [void]$grid.Columns.Add($colId)
    $form.Controls.Add($grid)
    $y += 390

    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text = 'Remove partnership'
    $btnRemove.Location = New-Object System.Drawing.Point(12, $y)
    $btnRemove.Size = New-Object System.Drawing.Size(170, 28)
    $btnRemove.Enabled = $false
    $form.Controls.Add($btnRemove)

    $btnWipe = New-Object System.Windows.Forms.Button
    $btnWipe.Text = 'Wipe account data'
    $btnWipe.Location = New-Object System.Drawing.Point(190, $y)
    $btnWipe.Size = New-Object System.Drawing.Size(170, 28)
    $btnWipe.Enabled = $false
    $form.Controls.Add($btnWipe)

    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text = 'Remove = stops syncing, mail stays on the phone.   Wipe = deletes this account from the phone on its next check-in.'
    $lblNote.Location = New-Object System.Drawing.Point(372, ($y + 6))
    $lblNote.Size = New-Object System.Drawing.Size(540, 18)
    $lblNote.ForeColor = [System.Drawing.Color]::DimGray
    $form.Controls.Add($lblNote)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'
    $btnClose.Location = New-Object System.Drawing.Point(918, $y)
    $btnClose.Size = New-Object System.Drawing.Size(130, 28)
    $form.Controls.Add($btnClose)
    $btnClose.Add_Click({ $form.Close() })

    $form | Add-Member -NotePropertyName 'Busy' -NotePropertyValue $false -Force
    $form | Add-Member -NotePropertyName 'Smtp' -NotePropertyValue ''     -Force

    $load = {
        $grid.Rows.Clear()
        $btnRemove.Enabled = $false; $btnWipe.Enabled = $false
        $s = Get-TargetSmtp -Text $txtMbx.Text
        if (-not $s) { $lblStatus.Text = 'Enter a mailbox address.'; return }
        $form.Smtp = $s
        $p = New-ProgressReporter -Label $lblStatus -Bar $pb -Activity 'Reading devices'
        & $p "Reading devices on $s..." -1
        $devs = Get-MailboxDevice -Smtp $s
        foreach ($d in $devs) {
            $ls = ''
            try { if ($d.LastSuccessSync) { $ls = ([datetime]$d.LastSuccessSync).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } } catch { }
            # The wipe column says where in the lifecycle a wipe is, because
            # "I wiped it" and "the phone has actually been wiped" are different
            # facts separated by however long the phone stays off.
            $wipe = if ($d.DeviceWipeAckTime)       { 'DONE ' + ([datetime]$d.DeviceWipeAckTime).ToLocalTime().ToString('MM-dd HH:mm') }
                    elseif ($d.DeviceWipeSentTime)  { 'sent, not acknowledged' }
                    elseif ($d.DeviceWipeRequestTime) { 'requested, not sent' }
                    else { '' }
            $name = if ("$($d.DeviceFriendlyName)".Trim()) { "$($d.DeviceFriendlyName)" } else { "$($d.DeviceId)" }
            [void]$grid.Rows.Add($name, "$($d.DeviceModel)", "$($d.DeviceOS)", "$($d.DeviceType)", $ls, "$($d.DeviceAccessState)", $wipe, "$($d.Identity)")
        }
        & $p "$($devs.Count) device(s)." 100
        $btnRemove.Enabled = ($devs.Count -gt 0); $btnWipe.Enabled = ($devs.Count -gt 0)
    }

    $btnLoad.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try { $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; & $load }
        catch { $lblStatus.Text = 'Load failed -- see the error dialog.'; Show-HandlerError -ErrorRecord $_ -Context 'Load (Get-MobileDeviceStatistics)' }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    $act = {
        param([string]$Verb, [scriptblock]$Do, [string]$Action, [string]$Warn)
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $sel = @($grid.SelectedRows)
            if ($sel.Count -eq 0) { $lblStatus.Text = 'Select one or more devices first.'; return }
            $names = @($sel | ForEach-Object { "$($_.Cells['Device'].Value)  ($($_.Cells['Model'].Value))" })
            $a = Show-ToolDialog(
                ("$Verb $($sel.Count) device(s) on $($form.Smtp)?" + "`r`n`r`n  " + ($names -join "`r`n  ") + "`r`n`r`n" + $Warn),
                'Confirm', 'OKCancel', 'Warning')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $ok = 0; $failed = 0; $errs = @()
            foreach ($r in $sel) {
                try { & $Do "$($r.Cells['Identity'].Value)"; $ok++ }
                catch { $failed++; $errs += "$($r.Cells['Device'].Value) -- $($_.Exception.Message)" }
            }
            $result = if ($failed -and $ok) { 'Partial' } elseif ($failed) { 'Failed' } else { 'Success' }
            Write-CompletedLog -Action $Action -Target (Get-LogTarget $form.Smtp) -Result $result `
                -Detail "$ok ok, $failed failed: $($names -join '; ')$(if ($errs) { ' | ' + ($errs -join ' | ') })"
            & $load
            $lblStatus.Text = "$Verb`: $ok done, $failed failed."
        } catch {
            $lblStatus.Text = "$Verb failed -- see the error dialog."
            Show-HandlerError -ErrorRecord $_ -Context $Verb
            Write-FailedLog -Action $Action -Target "$($form.Smtp)" -ErrorRecord $_
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $form.Busy = $false
        }
    }

    $btnRemove.Add_Click({
        & $act 'Remove partnership for' { param($id) Remove-MailboxDevice -Identity $id } 'RemoveMobileDevice' `
            'The device stops syncing. Mail already on it stays on it.'
    })
    $btnWipe.Add_Click({
        & $act 'Wipe account data from' { param($id) Clear-MailboxDeviceAccount -Identity $id } 'WipeMobileDeviceAccount' `
            'THIS ACCOUNT''s mail, contacts and calendar are deleted from the device on its next check-in. Personal data is not touched. This cannot be recalled once the device receives it.'
    })

    Set-ConsoleTheme -Form $form
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Mobile Devices' | Out-Null
    [void]$form.ShowDialog()
}
#endregion

Show-DeviceForm -Seed (Get-TargetSmtp -Text $SmtpAddress)
