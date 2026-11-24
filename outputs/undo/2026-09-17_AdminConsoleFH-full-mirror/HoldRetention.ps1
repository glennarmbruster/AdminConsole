#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
<#
    HoldRetention.ps1
    Litigation hold and retention policy on one mailbox: see them, set them. Author : Glenn Armbruster

    Author   : Glenn Armbruster
    Revised  : 2026-08-27
#>
# TOOLKIT-HELP-RUNS: reads one mailbox's hold and retention settings. Apply changes them and asks first.
# TOOLKIT-HELP-BEGIN
# Mailbox Health SHOWS holds. This one SETS them.
#
# THREE SETTINGS, AND THEY DO DIFFERENT THINGS:
#
#   Litigation hold    Nothing in the mailbox can be permanently deleted while
#                      it is on -- not by the user, not by retention, not by
#                      emptying anything. This is the legal-hold switch.
#                      Duration is in days from when each item was received;
#                      blank means indefinite.
#
#   Retention policy   Which set of retention tags applies: how long things
#                      stay before they are archived or deleted. The list comes
#                      from the tenant.
#
#   Retention hold     PAUSES the retention policy without removing it. For
#                      someone on long leave whose mail should not be aged out
#                      while they are away.
#
# LITIGATION HOLD NEEDS A LICENCE. Exchange Online Plan 2, E3, E5, or the
# Exchange Online Archiving add-on. On a mailbox without one, Set-Mailbox
# refuses and this tool shows you the refusal rather than a green tick.
#
# TAKING A HOLD OFF is the consequential direction: anything the hold was
# keeping becomes deletable again, and retention catches up on the next pass.
# The confirmation says so.
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

function Get-HoldState {
    param([Parameter(Mandatory)][string]$Smtp)
    $m = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
    return [PSCustomObject]@{
        Smtp        = "$($m.PrimarySmtpAddress)"
        Display     = "$($m.DisplayName)"
        Type        = "$($m.RecipientTypeDetails)"
        LitHold     = [bool]$m.LitigationHoldEnabled
        LitDuration = "$($m.LitigationHoldDuration)"
        LitDate     = "$($m.LitigationHoldDate)"
        LitOwner    = "$($m.LitigationHoldOwner)"
        Policy      = "$($m.RetentionPolicy)"
        RetHold     = [bool]$m.RetentionHoldEnabled
        InPlace     = @($m.InPlaceHolds).Count
    }
}

function Get-RetentionPolicyName {
    # Learn 2026-08-27 (Get-RetentionPolicy): returns Name, Identity.
    return @(Invoke-EXOCommand -Name 'Get-RetentionPolicy' -Parameters @{ ErrorAction = 'Stop' } |
             ForEach-Object { "$($_.Name)" } | Sort-Object)
}
#endregion

#region UI --------------------------------------------------------------------

function Show-HoldForm {
    param([string]$Seed)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Hold and Retention'
    $form.ClientSize = New-Object System.Drawing.Size(760, 470)
    $form.StartPosition = 'Manual'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $y = 12
    $lblMbx = New-Object System.Windows.Forms.Label
    $lblMbx.Text = 'Mailbox:'
    $lblMbx.Location = New-Object System.Drawing.Point(12, ($y + 3)); $lblMbx.Size = New-Object System.Drawing.Size(70, 18)
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
    $txtNow.Location = New-Object System.Drawing.Point(12, $y); $txtNow.Size = New-Object System.Drawing.Size(736, 120)
    $txtNow.Multiline = $true; $txtNow.ReadOnly = $true; $txtNow.WordWrap = $true; $txtNow.ScrollBars = 'Vertical'
    $txtNow.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtNow)
    $y += 130

    $grp = New-Object System.Windows.Forms.GroupBox
    $grp.Text = 'Set'
    $grp.Location = New-Object System.Drawing.Point(12, $y); $grp.Size = New-Object System.Drawing.Size(736, 150)
    $form.Controls.Add($grp)

    $chkLit = New-Object System.Windows.Forms.CheckBox
    $chkLit.Text = 'Litigation hold'; $chkLit.Location = New-Object System.Drawing.Point(12, 24); $chkLit.Size = New-Object System.Drawing.Size(140, 22)
    $grp.Controls.Add($chkLit)
    $lblDur = New-Object System.Windows.Forms.Label
    $lblDur.Text = 'for'; $lblDur.Location = New-Object System.Drawing.Point(160, 27); $lblDur.Size = New-Object System.Drawing.Size(24, 18)
    $grp.Controls.Add($lblDur)
    $txtDur = New-Object System.Windows.Forms.TextBox
    $txtDur.Location = New-Object System.Drawing.Point(186, 23); $txtDur.Size = New-Object System.Drawing.Size(70, 24); $txtDur.MaxLength = 5
    $grp.Controls.Add($txtDur)
    $lblDur2 = New-Object System.Windows.Forms.Label
    $lblDur2.Text = 'days  (blank = indefinite)'; $lblDur2.Location = New-Object System.Drawing.Point(262, 27); $lblDur2.Size = New-Object System.Drawing.Size(200, 18)
    $lblDur2.ForeColor = [System.Drawing.Color]::DimGray
    $grp.Controls.Add($lblDur2)

    $lblPol = New-Object System.Windows.Forms.Label
    $lblPol.Text = 'Retention policy:'; $lblPol.Location = New-Object System.Drawing.Point(12, 62); $lblPol.Size = New-Object System.Drawing.Size(120, 18)
    $grp.Controls.Add($lblPol)
    $cboPol = New-Object System.Windows.Forms.ComboBox
    $cboPol.DropDownStyle = 'DropDownList'; $cboPol.Location = New-Object System.Drawing.Point(134, 58); $cboPol.Size = New-Object System.Drawing.Size(400, 24)
    $grp.Controls.Add($cboPol)

    $chkRet = New-Object System.Windows.Forms.CheckBox
    $chkRet.Text = 'Retention hold  (pause the policy -- long leave)'; $chkRet.Location = New-Object System.Drawing.Point(12, 96); $chkRet.Size = New-Object System.Drawing.Size(400, 22)
    $grp.Controls.Add($chkRet)

    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text = 'Apply'; $btnApply.Location = New-Object System.Drawing.Point(594, 110); $btnApply.Size = New-Object System.Drawing.Size(130, 28); $btnApply.Enabled = $false
    $grp.Controls.Add($btnApply)
    $y += 160

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Location = New-Object System.Drawing.Point(618, $y); $btnClose.Size = New-Object System.Drawing.Size(130, 28)
    $form.Controls.Add($btnClose); $btnClose.Add_Click({ $form.Close() })

    $form | Add-Member -NotePropertyName 'Busy'  -NotePropertyValue $false -Force
    $form | Add-Member -NotePropertyName 'State' -NotePropertyValue $null  -Force

    $show = {
        $h = $form.State
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("$($h.Display)  <$($h.Smtp)>  $($h.Type)")
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("  Litigation hold   : $(if ($h.LitHold) { 'ON' } else { 'off' })" +
            $(if ($h.LitHold) { "   duration $($h.LitDuration)   since $($h.LitDate)   by $($h.LitOwner)" } else { '' }))
        [void]$sb.AppendLine("  Retention policy  : $(if ($h.Policy) { $h.Policy } else { '(none)' })")
        [void]$sb.AppendLine("  Retention hold    : $(if ($h.RetHold) { 'ON -- policy paused' } else { 'off' })")
        [void]$sb.AppendLine("  In-place holds    : $($h.InPlace)  (eDiscovery / Purview -- not editable here)")
        $txtNow.Text = $sb.ToString()
        $chkLit.Checked = $h.LitHold
        $txtDur.Text = if ($h.LitDuration -and $h.LitDuration -ne 'Unlimited') { "$($h.LitDuration)" -replace '\..*$', '' } else { '' }
        $chkRet.Checked = $h.RetHold
        if ($cboPol.Items.Count) {
            $idx = $cboPol.Items.IndexOf($h.Policy)
            $cboPol.SelectedIndex = $(if ($idx -ge 0) { $idx } else { 0 })
        }
    }

    $btnLoad.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $btnApply.Enabled = $false
            $s = Get-TargetSmtp -Text $txtMbx.Text
            if (-not $s) { $lblStatus.Text = 'Enter a mailbox address.'; return }
            if ($cboPol.Items.Count -eq 0) {
                [void]$cboPol.Items.Add('(none)')
                foreach ($n in (Get-RetentionPolicyName)) { [void]$cboPol.Items.Add($n) }
            }
            $form.State = Get-HoldState -Smtp $s
            & $show
            $lblStatus.Text = 'Loaded. Change what you need and press Apply.'
            $btnApply.Enabled = $true
        } catch {
            $txtNow.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox'
            $lblStatus.Text = 'Could not read the mailbox.'
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $form.Busy = $false
        }
    })

    $btnApply.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $h = $form.State
            if (-not $h) { return }

            # Only what CHANGED is sent. Set-Mailbox with the same value is
            # harmless but shows up in the audit log as a change that was not
            # one, and a duration typo would otherwise re-set a hold nobody
            # touched.
            $p = @{ Identity = $h.Smtp; ErrorAction = 'Stop' }
            $changes = @()

            $wantLit = $chkLit.Checked
            $durTxt  = "$($txtDur.Text)".Trim()
            $wantDur = if ($durTxt) { $durTxt } else { 'Unlimited' }
            if ($durTxt -and ($durTxt -notmatch '^\d+$' -or [int]$durTxt -lt 1)) { $lblStatus.Text = 'Duration must be a whole number of days, or blank.'; return }
            $curDur  = if ($h.LitDuration -and $h.LitDuration -ne 'Unlimited') { "$($h.LitDuration)" -replace '\..*$', '' } else { 'Unlimited' }
            if ($wantLit -ne $h.LitHold) {
                $p.LitigationHoldEnabled = $wantLit
                $changes += "Litigation hold: $(if ($h.LitHold) {'ON'} else {'off'}) -> $(if ($wantLit) {'ON'} else {'off'})"
            }
            if ($wantLit -and $wantDur -ne $curDur) {
                $p.LitigationHoldDuration = $wantDur
                $changes += "Litigation hold duration: $curDur -> $wantDur"
            }

            $wantPol = "$($cboPol.SelectedItem)"; if ($wantPol -eq '(none)') { $wantPol = '' }
            if ($wantPol -ne $h.Policy) {
                $p.RetentionPolicy = $(if ($wantPol) { $wantPol } else { $null })
                $changes += "Retention policy: $(if ($h.Policy) {$h.Policy} else {'(none)'}) -> $(if ($wantPol) {$wantPol} else {'(none)'})"
            }

            if ($chkRet.Checked -ne $h.RetHold) {
                $p.RetentionHoldEnabled = $chkRet.Checked
                $changes += "Retention hold: $(if ($h.RetHold) {'ON'} else {'off'}) -> $(if ($chkRet.Checked) {'ON'} else {'off'})"
            }

            if ($changes.Count -eq 0) { $lblStatus.Text = 'Nothing changed.'; return }

            $warn = if ($h.LitHold -and -not $wantLit) {
                "`r`n`r`nTAKING THE LITIGATION HOLD OFF. Everything it was preserving becomes deletable, and retention will catch up on its next pass. Make sure Legal has released it."
            } else { '' }
            $a = Show-ToolDialog(
                ("Apply to $($h.Smtp)?" + "`r`n`r`n  " + ($changes -join "`r`n  ") + $warn),
                'Confirm', 'OKCancel', $(if ($warn) { 'Warning' } else { 'Question' }))
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters $p | Out-Null
            Write-CompletedLog -Action 'SetHoldRetention' -Target (Get-LogTarget $h.Smtp) -Result 'Success' -Detail ($changes -join '; ')
            $form.State = Get-HoldState -Smtp $h.Smtp
            & $show
            $lblStatus.Text = "Applied: $($changes -join '; ')"
        } catch {
            $msg = "$($_.Exception.Message)"
            $hint = if ($msg -match 'licen') { "  This usually means the mailbox has no Exchange Plan 2 / E3 / E5 / Archiving licence -- litigation hold needs one." } else { '' }
            $lblStatus.Text = "Failed: $msg$hint"
            $txtNow.Text = (Format-ErrorDetail -ErrorRecord $_ -Context 'Set-Mailbox') + $hint
            Write-FailedLog -Action 'SetHoldRetention' -Target "$($form.State.Smtp)" -ErrorRecord $_
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $form.Busy = $false
        }
    })

    Set-ConsoleTheme -Form $form
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Hold and Retention' | Out-Null
    [void]$form.ShowDialog()
}
#endregion

Show-HoldForm -Seed (Get-TargetSmtp -Text $SmtpAddress)
