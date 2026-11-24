#Requires -Version 5.1
# TOOLKIT-VARIANT: OnPrem
<#
    EnableArchive.ps1
    Turn on the online archive for a mailbox that is filling up. Author : Glenn Armbruster

    Author   : Glenn Armbruster
    Revised  : 2026-08-27
#>
# TOOLKIT-HELP-RUNS: reads the on-premises remote-mailbox record. Enable changes it and asks first.
# TOOLKIT-HELP-BEGIN
# "Mailbox is full."  MBX-Stats says where the space went; this is the fix.
#
# An online archive is a second mailbox in the cloud that the retention policy
# moves old mail into. It is the answer to "full" that does not involve asking
# somebody to delete things, and it needs a licence that most user plans
# already include.
#
# ENABLED ON-PREMISES, NOT IN THE CLOUD, AND THAT IS NOT OPTIONAL.
#
# In a hybrid directory the on-premises object is the source of truth. Enabling
# the archive in Exchange Online directly sets msExchArchiveGuid there; the next
# directory sync writes the on-premises value -- which is empty -- back over it,
# and the archive is disconnected. New User enables archives with
# Enable-RemoteMailbox -Archive for exactly this reason, and this tool does the
# same thing for a mailbox that already exists.
#
# So: this runs on-premises, and the archive appears in the cloud after a
# directory sync cycle. Mailbox Health shows both sides' ArchiveGuid; when they
# match, it is done.
#
# Cannot be undone from here. Disable-RemoteMailbox -Archive exists, but an
# archive with mail in it is not something to switch off from a button.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

#region UI --------------------------------------------------------------------

function Show-ArchiveForm {
    param([string]$Seed)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Enable Archive -- on-premises'
    $form.ClientSize = New-Object System.Drawing.Size(720, 330)
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
    $txtNow.Location = New-Object System.Drawing.Point(12, $y); $txtNow.Size = New-Object System.Drawing.Size(696, 150)
    $txtNow.Multiline = $true; $txtNow.ReadOnly = $true; $txtNow.WordWrap = $true; $txtNow.ScrollBars = 'Vertical'
    $txtNow.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtNow)
    $y += 160

    $btnGo = New-Object System.Windows.Forms.Button
    $btnGo.Text = 'Enable archive'; $btnGo.Location = New-Object System.Drawing.Point(12, $y); $btnGo.Size = New-Object System.Drawing.Size(150, 30); $btnGo.Enabled = $false
    $form.Controls.Add($btnGo)
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Location = New-Object System.Drawing.Point(578, $y); $btnClose.Size = New-Object System.Drawing.Size(130, 30)
    $form.Controls.Add($btnClose); $btnClose.Add_Click({ $form.Close() })

    $form | Add-Member -NotePropertyName 'Busy' -NotePropertyValue $false -Force
    $form | Add-Member -NotePropertyName 'Rm'   -NotePropertyValue $null  -Force

    $load = {
        $s = Get-TargetSmtp -Text $txtMbx.Text
        if (-not $s) { $lblStatus.Text = 'Enter a mailbox address.'; return }
        # Get-RemoteMailbox is on-premises ONLY -- no collision, no wrapper
        # needed. Its absence from the cloud is the point: this is the object
        # that decides what the cloud copy looks like.
        $p = @{ Identity = $s; ErrorAction = 'Stop' }
        if ($DomainController) { $p.DomainController = $DomainController }
        $rm = Get-RemoteMailbox @p
        $form.Rm = $rm
        $empty = '00000000-0000-0000-0000-000000000000'
        $has = ("$($rm.ArchiveGuid)" -and "$($rm.ArchiveGuid)" -ne $empty)
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("$($rm.DisplayName)  <$($rm.PrimarySmtpAddress)>")
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("  Type            : $($rm.RecipientTypeDetails)")
        [void]$sb.AppendLine("  Archive (on-prem): $(if ($has) { 'ENABLED  ' + $rm.ArchiveGuid } else { 'not enabled' })")
        [void]$sb.AppendLine("  Archive state   : $($rm.ArchiveState)")
        [void]$sb.AppendLine("  Archive name    : $(if ("$($rm.ArchiveName)") { @($rm.ArchiveName) -join ', ' } else { '(default)' })")
        if ($has) { [void]$sb.AppendLine(''); [void]$sb.AppendLine('  Already enabled here. If the cloud does not show it, wait for a sync cycle, then check Mailbox Health.') }
        $txtNow.Text = $sb.ToString()
        $btnGo.Enabled = -not $has
        $lblStatus.Text = if ($has) { 'Archive already enabled on-premises.' } else { 'No archive. Enable will create one.' }
    }

    $btnLoad.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try { $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; $btnGo.Enabled = $false; & $load }
        catch { $lblStatus.Text = 'Load failed -- see the error dialog.'; Show-HandlerError -ErrorRecord $_ -Context 'Load (Get-RemoteMailbox)' }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    $btnGo.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $rm = $form.Rm; if (-not $rm) { return }
            $a = Show-ToolDialog(
                ("Enable the online archive for:`r`n`r`n  $($rm.PrimarySmtpAddress)`r`n`r`nThis runs on-premises. The archive appears in the cloud after the next directory sync. It cannot be switched off from this tool."),
                'Confirm', 'OKCancel', 'Question')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            # Learn 2026-08-27 (Enable-RemoteMailbox): -Identity, -Archive
            # switch. Same call New User makes at creation.
            $p = @{ Identity = "$($rm.PrimarySmtpAddress)"; Archive = $true; ErrorAction = 'Stop' }
            if ($DomainController) { $p.DomainController = $DomainController }
            Enable-RemoteMailbox @p | Out-Null
            Write-CompletedLog -Action 'EnableArchive' -Target (Get-LogTarget $rm.PrimarySmtpAddress) -Result 'Success' -Detail 'Enable-RemoteMailbox -Archive (on-premises)' -Environment $Environment
            & $load
            $lblStatus.Text = 'Archive enabled on-premises. Allow a sync cycle before it appears in the cloud.'
        } catch { $lblStatus.Text = 'Enable failed -- see the error dialog.'; Show-HandlerError -ErrorRecord $_ -Context 'Enable-RemoteMailbox -Archive' }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
    })

    Set-ConsoleTheme -Form $form
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Enable Archive' | Out-Null
    [void]$form.ShowDialog()
}
#endregion

Show-ArchiveForm -Seed (Get-TargetSmtp -Text $SmtpAddress)
