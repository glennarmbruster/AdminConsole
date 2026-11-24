#Requires -Version 5.1
# TOOLKIT-VARIANT: OnPrem
# TOOLKIT-REQUIRES: OwnProcess  -- New-/Set-/Remove-MailContact exist in BOTH
#   Exchanges and create or delete a different object in each. Sharing a
#   process with a connected Exchange Online session is how a contact lands in
#   the wrong directory with no error.
<#
    MailContacts.ps1
    Create, update or remove a mail contact -- an outside person who needs to be in the address book. Author : Glenn Armbruster

    Author   : Glenn Armbruster
    Revised  : 2026-08-27
#>
# TOOLKIT-HELP-RUNS: reads the contact. Apply creates, updates or removes it on-premises, and asks first.
# TOOLKIT-HELP-BEGIN
# A mail contact is an EXTERNAL address that shows up in the company address
# book -- a vendor, an auditor, outside counsel, an external contact's personal
# address. It has no mailbox and no login here; mail to it goes out to wherever
# it really lives.
#
# CREATED ON-PREMISES AND SYNCED UP, like every other object this toolkit
# makes in a hybrid directory. Allow a sync cycle before it appears in the cloud
# address book.
#
# THE EXTERNAL ADDRESS MUST BE A FULL ADDRESS WITH ITS OWN DOMAIN. A bare name
# would get the company domain appended, which is the opposite of what a
# contact is for, so this tool refuses it rather than guessing.
#
#   Add      creates it. Directory decides the OU, as in Mail Object.
#   Update   changes the display name or the external address on one that exists.
#   Remove   deletes it. Nothing is soft-deleted; a contact has no mailbox to
#            keep. It is recreated in thirty seconds if that was a mistake.
#
# Preview first. Apply does what Preview said.
# TOOLKIT-HELP-END

param([string]$SmtpAddress)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

#region Data ------------------------------------------------------------------

function Find-ExistingContact {
    # Get-Recipient rather than Get-MailContact so the answer to "what is at
    # this address" is honest when it is a mailbox and not a contact at all.
    param([string]$Address)
    try {
        $r = Invoke-OnPremCommand -Name 'Get-Recipient' -Parameters @{ Identity = $Address; ErrorAction = 'Stop' }
        return [PSCustomObject]@{ Found = $true; Kind = "$($r.RecipientTypeDetails)"; Display = "$($r.DisplayName)"; Identity = "$($r.Identity)" }
    } catch { return [PSCustomObject]@{ Found = $false; Kind = ''; Display = ''; Identity = '' } }
}
#endregion

#region UI --------------------------------------------------------------------

function Show-ContactForm {
    param([string]$Seed)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Mail Contacts -- on-premises'
    $form.ClientSize = New-Object System.Drawing.Size(820, 520)
    $form.StartPosition = 'CenterScreen'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $y = 12
    $lblAddr = New-Object System.Windows.Forms.Label
    $lblAddr.Text = 'External address (full, with its own domain):'; $lblAddr.Location = New-Object System.Drawing.Point(12, $y); $lblAddr.Size = New-Object System.Drawing.Size(400, 18)
    $form.Controls.Add($lblAddr); $y += 20
    $txtAddr = New-Object System.Windows.Forms.TextBox
    $txtAddr.Location = New-Object System.Drawing.Point(12, $y); $txtAddr.Size = New-Object System.Drawing.Size(420, 24)
    $txtAddr.Font = New-Object System.Drawing.Font('Consolas', 9); $txtAddr.Text = "$Seed"
    $form.Controls.Add($txtAddr)
    $lblFound = New-Object System.Windows.Forms.Label
    $lblFound.Location = New-Object System.Drawing.Point(444, ($y + 3)); $lblFound.Size = New-Object System.Drawing.Size(360, 18); $lblFound.ForeColor = [System.Drawing.Color]::DimGray
    $form.Controls.Add($lblFound); $y += 34

    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Text = 'Display name:'; $lblName.Location = New-Object System.Drawing.Point(12, ($y + 3)); $lblName.Size = New-Object System.Drawing.Size(100, 18)
    $form.Controls.Add($lblName)
    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Location = New-Object System.Drawing.Point(114, $y); $txtName.Size = New-Object System.Drawing.Size(318, 24)
    $form.Controls.Add($txtName); $y += 34

    $grpAct = New-Object System.Windows.Forms.GroupBox
    $grpAct.Text = 'Action'; $grpAct.Location = New-Object System.Drawing.Point(12, $y); $grpAct.Size = New-Object System.Drawing.Size(300, 52)
    $form.Controls.Add($grpAct)
    $rbAdd = New-Object System.Windows.Forms.RadioButton
    $rbAdd.Text = 'Add'; $rbAdd.Checked = $true; $rbAdd.Location = New-Object System.Drawing.Point(12, 20); $rbAdd.Size = New-Object System.Drawing.Size(70, 22)
    $grpAct.Controls.Add($rbAdd)
    $rbUpd = New-Object System.Windows.Forms.RadioButton
    $rbUpd.Text = 'Update'; $rbUpd.Location = New-Object System.Drawing.Point(90, 20); $rbUpd.Size = New-Object System.Drawing.Size(90, 22)
    $grpAct.Controls.Add($rbUpd)
    $rbRem = New-Object System.Windows.Forms.RadioButton
    $rbRem.Text = 'Remove'; $rbRem.Location = New-Object System.Drawing.Point(190, 20); $rbRem.Size = New-Object System.Drawing.Size(90, 22)
    $grpAct.Controls.Add($rbRem)

    $lblDir = New-Object System.Windows.Forms.Label
    $lblDir.Text = 'Directory:'; $lblDir.Location = New-Object System.Drawing.Point(330, ($y + 22)); $lblDir.Size = New-Object System.Drawing.Size(70, 18)
    $form.Controls.Add($lblDir)
    $cboDir = New-Object System.Windows.Forms.ComboBox
    $cboDir.DropDownStyle = 'DropDownList'; $cboDir.Location = New-Object System.Drawing.Point(402, ($y + 18)); $cboDir.Size = New-Object System.Drawing.Size(160, 24)
    foreach ($d in @(Get-TenantDirectory)) { [void]$cboDir.Items.Add("$($d.Id)") }
    if ($cboDir.Items.Count) { $def = Get-DefaultTenantDirectory; $cboDir.SelectedItem = $(if ($def) { "$($def.Id)" } else { $cboDir.Items[0] }) }
    $form.Controls.Add($cboDir)
    $y += 60

    $lblOu = New-Object System.Windows.Forms.Label
    $lblOu.Text = 'OU:'; $lblOu.Location = New-Object System.Drawing.Point(12, ($y + 3)); $lblOu.Size = New-Object System.Drawing.Size(30, 18)
    $form.Controls.Add($lblOu)
    $txtOu = New-Object System.Windows.Forms.TextBox
    $txtOu.Location = New-Object System.Drawing.Point(44, $y); $txtOu.Size = New-Object System.Drawing.Size(760, 24); $txtOu.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtOu); $y += 34

    $btnPreview = New-Object System.Windows.Forms.Button
    $btnPreview.Text = 'Preview'; $btnPreview.Location = New-Object System.Drawing.Point(12, $y); $btnPreview.Size = New-Object System.Drawing.Size(130, 30)
    $form.Controls.Add($btnPreview)
    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text = 'Apply'; $btnApply.Location = New-Object System.Drawing.Point(150, $y); $btnApply.Size = New-Object System.Drawing.Size(130, 30)
    $form.Controls.Add($btnApply)
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Location = New-Object System.Drawing.Point(674, $y); $btnClose.Size = New-Object System.Drawing.Size(130, 30)
    $form.Controls.Add($btnClose); $btnClose.Add_Click({ $form.Close() })
    $y += 40

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Location = New-Object System.Drawing.Point(12, $y); $txtOut.Size = New-Object System.Drawing.Size(792, 220)
    $txtOut.Multiline = $true; $txtOut.ReadOnly = $true; $txtOut.WordWrap = $true; $txtOut.ScrollBars = 'Vertical'
    $txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtOut)

    $form | Add-Member -NotePropertyName 'Busy' -NotePropertyValue $false -Force

    $syncOu = {
        # MailContactOu from Tenant.json, via the same helper Mail Object uses.
        # Empty means it is not configured for this directory -- type it, and
        # add MailContactOu to the directory so the next person does not have to.
        $txtOu.Text = Get-TenantDirectoryOu -DirectoryId "$($cboDir.SelectedItem)" -Kind 'MailContact'
    }
    $cboDir.Add_SelectedIndexChanged($syncOu); & $syncOu

    $syncAddr = {
        $a = "$($txtAddr.Text)".Trim()
        if ($a -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { $lblFound.Text = ''; return }
        $f = Find-ExistingContact -Address $a
        $lblFound.Text = if ($f.Found) { "exists: $($f.Kind) -- $($f.Display)" } else { 'not found -- Add will create it' }
        if ($f.Found -and $f.Kind -match 'MailContact' -and -not "$($txtName.Text)".Trim()) { $txtName.Text = $f.Display }
    }
    $txtAddr.Add_Leave($syncAddr)

    $gather = {
        $a = "$($txtAddr.Text)".Trim()
        if ($a -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { return @{ Ok = $false; Msg = 'Enter the full external address, including its domain.' } }
        $dom = Get-MailDomain
        if ($dom -and $a.ToLower().EndsWith('@' + $dom.ToLower())) { return @{ Ok = $false; Msg = "'$a' is on the company domain. A contact is for an OUTSIDE address." } }
        $name = "$($txtName.Text)".Trim()
        $act = if ($rbAdd.Checked) { 'Add' } elseif ($rbUpd.Checked) { 'Update' } else { 'Remove' }
        if ($act -ne 'Remove' -and -not $name) { return @{ Ok = $false; Msg = 'Enter a display name.' } }
        $ou = "$($txtOu.Text)".Trim()
        if ($act -eq 'Add' -and -not $ou) { return @{ Ok = $false; Msg = 'No OU. Set ContactOu or ResourceOu for this directory in Tenant Manager, or type it.' } }
        return @{ Ok = $true; Address = $a; Name = $name; Alias = (($a -split '@')[0] -replace '[^A-Za-z0-9._-]', ''); Action = $act; Dir = "$($cboDir.SelectedItem)"; Ou = $ou; Found = (Find-ExistingContact -Address $a) }
    }

    $btnPreview.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $g = & $gather
            if (-not $g.Ok) { $txtOut.Text = $g.Msg; return }
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine('PREVIEW -- nothing has been changed'); [void]$sb.AppendLine('')
            [void]$sb.AppendLine("  Action  : $($g.Action)"); [void]$sb.AppendLine("  Address : $($g.Address)")
            switch ($g.Action) {
                'Add' {
                    if ($g.Found.Found) { [void]$sb.AppendLine("  ALREADY EXISTS as $($g.Found.Kind) -- '$($g.Found.Display)'. Nothing would be created."); break }
                    [void]$sb.AppendLine("  Name    : $($g.Name)"); [void]$sb.AppendLine("  Alias   : $($g.Alias)"); [void]$sb.AppendLine("  OU      : $($g.Ou)")
                    [void]$sb.AppendLine(''); [void]$sb.AppendLine('  Would run New-MailContact on-premises.')
                }
                'Update' {
                    if (-not $g.Found.Found) { [void]$sb.AppendLine('  NOT FOUND. Nothing to update -- use Add.'); break }
                    if ($g.Found.Kind -notmatch 'MailContact') { [void]$sb.AppendLine("  REFUSED. That address is a $($g.Found.Kind), not a contact."); break }
                    [void]$sb.AppendLine("  Current name : $($g.Found.Display)"); [void]$sb.AppendLine("  New name     : $($g.Name)")
                    [void]$sb.AppendLine(''); [void]$sb.AppendLine('  Would run Set-MailContact on-premises.')
                }
                'Remove' {
                    if (-not $g.Found.Found) { [void]$sb.AppendLine('  NOT FOUND. Nothing to remove.'); break }
                    if ($g.Found.Kind -notmatch 'MailContact') { [void]$sb.AppendLine("  REFUSED. '$($g.Address)' is a $($g.Found.Kind), not a contact. This tool will not delete it."); break }
                    [void]$sb.AppendLine("  Would REMOVE contact '$($g.Found.Display)'. Not recoverable, but trivially recreated.")
                }
            }
            $txtOut.Text = $sb.ToString()
        } catch { $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Preview' }
        finally { $form.Busy = $false }
    })

    $btnApply.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $g = & $gather
            if (-not $g.Ok) { $txtOut.Text = $g.Msg; return }
            $f = $g.Found
            if ($g.Action -eq 'Add' -and $f.Found) { $txtOut.Text = "'$($g.Address)' already exists as $($f.Kind). Nothing created."; return }
            if ($g.Action -ne 'Add' -and -not $f.Found) { $txtOut.Text = "'$($g.Address)' not found. Nothing done."; return }
            if ($g.Action -ne 'Add' -and $f.Kind -notmatch 'MailContact') { $txtOut.Text = "REFUSED. '$($g.Address)' is a $($f.Kind), not a contact."; return }

            $a = Show-ToolDialog(("$($g.Action) mail contact on-premises?`r`n`r`n  $($g.Address)`r`n  $(if ($g.Action -ne 'Remove') { $g.Name } else { $f.Display })"), 'Confirm', 'OKCancel', $(if ($g.Action -eq 'Remove') { 'Warning' } else { 'Question' }))
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            switch ($g.Action) {
                'Add' {
                    # Learn 2026-08-27 (New-MailContact): -Name, -ExternalEmailAddress, -Alias, -OrganizationalUnit, -DisplayName.
                    Invoke-OnPremCommand -Name 'New-MailContact' -Parameters @{ Name = $g.Name; DisplayName = $g.Name; Alias = $g.Alias; ExternalEmailAddress = $g.Address; OrganizationalUnit = $g.Ou; ErrorAction = 'Stop' } | Out-Null
                    $txtOut.Text = "CREATED`r`n`r`n  $($g.Name)  <$($g.Address)>`r`n  OU: $($g.Ou)`r`n`r`n  Allow a sync cycle before it appears in the cloud address book."
                }
                'Update' {
                    Invoke-OnPremCommand -Name 'Set-MailContact' -Parameters @{ Identity = $f.Identity; Name = $g.Name; DisplayName = $g.Name; ErrorAction = 'Stop' } | Out-Null
                    $txtOut.Text = "UPDATED`r`n`r`n  $($f.Display)  ->  $($g.Name)"
                }
                'Remove' {
                    Invoke-OnPremCommand -Name 'Remove-MailContact' -Parameters @{ Identity = $f.Identity; Confirm = $false; ErrorAction = 'Stop' } | Out-Null
                    $txtOut.Text = "REMOVED`r`n`r`n  $($f.Display)  <$($g.Address)>"
                }
            }
            Write-CompletedLog -Action "$($g.Action)MailContact" -Target (Get-LogTarget $g.Address) -Result 'Success' -Environment $g.Dir -Detail "$($g.Action) contact '$(if ($g.Action -ne 'Remove') { $g.Name } else { $f.Display })'$(if ($g.Action -eq 'Add') { " in $($g.Ou)" })"
            & $syncAddr
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Apply'
            Write-FailedLog -Action 'MailContact' -Target "$($txtAddr.Text)" -ErrorRecord $_
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $form.Busy = $false
        }
    })

    Set-ConsoleTheme -Form $form
    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Mail Contacts' | Out-Null
    [void]$form.ShowDialog()
}
#endregion

Show-ContactForm -Seed "$SmtpAddress"
