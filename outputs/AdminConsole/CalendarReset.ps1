#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
<#
    CalendarReset.ps1
    Replace a calendar's permissions wholesale: clear it, then grant from two pasted lists. Author : Glenn Armbruster

    Author   : Glenn Armbruster
    Revised  : 2026-08-27
#>
# TOOLKIT-HELP-RUNS: reads the calendar's permissions. Apply REMOVES all of them and grants the two lists. Preview first.
# TOOLKIT-HELP-BEGIN
# "Clear the calendar and give Editor to these people and Reviewer to those."
#
# This started as a one-off script and came up a fourth time. Paste names or
# addresses, one per line, into the two boxes. Preview shows what will be
# removed and what will be granted. Apply does it.
#
# WHAT "CLEAR" MEANS HERE. Every explicit calendar permission is removed. Full
# Access, Send As and the rest of the mailbox are NOT touched -- this is the
# calendar folder only. For a full mailbox reset use Mailbox Permissions.
#
# DEFAULT AND ANONYMOUS ARE SET, NOT REMOVED -- every folder has them and
# Exchange refuses to delete them. The Default checkbox decides whether Default
# is set to None (nobody outside the lists sees anything, not even busy blocks)
# or left as it is (usually AvailabilityOnly -- free/busy for the company).
# That is a decision about what everybody else can see, so it is a checkbox
# and not a default.
#
# NAMES ARE RESOLVED BEFORE ANYTHING IS REMOVED. A list with a typo in it is
# found at Preview, not after the calendar has already been cleared.
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

function Get-CalendarFolderId {
    # FolderType, never Name -- names are localised and renameable.
    param([Parameter(Mandatory)][string]$Smtp)
    $f = @(Invoke-EXOCommand -Name 'Get-MailboxFolderStatistics' -Parameters @{ Identity = $Smtp; FolderScope = 'Calendar'; ErrorAction = 'Stop' })
    $p = $f | Where-Object { "$($_.FolderType)" -eq 'Calendar' } | Select-Object -First 1
    if (-not $p) { $p = $f | Select-Object -First 1 }
    if (-not $p) { throw "No calendar folder found on $Smtp." }
    return ($Smtp + ':\' + $p.Name)
}

function Resolve-NameList {
    <#
        Each line -> a recipient, or a reason it is not one.

        Done up front and in full, because the alternative is discovering the
        typo on line 14 after lines 1-13 have been granted and the old
        permissions are already gone.
    #>
    param([string]$Text)
    $out = @()
    foreach ($raw in ("$Text" -split "`r?`n")) {
        $n = "$raw".Trim()
        if (-not $n) { continue }
        try {
            $r = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{ Identity = $n; ErrorAction = 'Stop' }
            $out += [PSCustomObject]@{ Input = $n; Ok = $true;  Smtp = "$($r.PrimarySmtpAddress)"; Display = "$($r.DisplayName)"; Why = '' }
        } catch {
            $out += [PSCustomObject]@{ Input = $n; Ok = $false; Smtp = ''; Display = ''; Why = "$($_.Exception.Message)" }
        }
    }
    return $out
}
#endregion

#region UI --------------------------------------------------------------------

function Show-ResetForm {
    param([string]$Seed)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Calendar Reset'
    $form.ClientSize = New-Object System.Drawing.Size(900, 640)
    $form.StartPosition = 'Manual'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $y = 12
    $lblMbx = New-Object System.Windows.Forms.Label
    $lblMbx.Text = 'Calendar mailbox:'; $lblMbx.Location = New-Object System.Drawing.Point(12, ($y + 3)); $lblMbx.Size = New-Object System.Drawing.Size(120, 18)
    $form.Controls.Add($lblMbx)
    $txtMbx = New-Object System.Windows.Forms.TextBox
    $txtMbx.Location = New-Object System.Drawing.Point(134, $y); $txtMbx.Size = New-Object System.Drawing.Size(380, 24)
    $txtMbx.Font = New-Object System.Drawing.Font('Consolas', 9); $txtMbx.Text = "$Seed"
    $form.Controls.Add($txtMbx)
    $chkDef = New-Object System.Windows.Forms.CheckBox
    $chkDef.Text = 'Set Default to None (hide free/busy from everyone not listed)'
    $chkDef.Location = New-Object System.Drawing.Point(530, ($y + 2)); $chkDef.Size = New-Object System.Drawing.Size(360, 22)
    $form.Controls.Add($chkDef)
    $y += 34

    $lblEd = New-Object System.Windows.Forms.Label
    $lblEd.Text = 'Editor  (add and delete)  -- one per line:'; $lblEd.Location = New-Object System.Drawing.Point(12, $y); $lblEd.Size = New-Object System.Drawing.Size(430, 18)
    $form.Controls.Add($lblEd)
    $lblRv = New-Object System.Windows.Forms.Label
    $lblRv.Text = 'Reviewer  (view only)  -- one per line:'; $lblRv.Location = New-Object System.Drawing.Point(458, $y); $lblRv.Size = New-Object System.Drawing.Size(430, 18)
    $form.Controls.Add($lblRv)
    $y += 20

    $txtEd = New-Object System.Windows.Forms.TextBox
    $txtEd.Location = New-Object System.Drawing.Point(12, $y); $txtEd.Size = New-Object System.Drawing.Size(430, 180)
    $txtEd.Multiline = $true; $txtEd.ScrollBars = 'Vertical'; $txtEd.AcceptsReturn = $true
    $txtEd.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtEd)
    $txtRv = New-Object System.Windows.Forms.TextBox
    $txtRv.Location = New-Object System.Drawing.Point(458, $y); $txtRv.Size = New-Object System.Drawing.Size(430, 180)
    $txtRv.Multiline = $true; $txtRv.ScrollBars = 'Vertical'; $txtRv.AcceptsReturn = $true
    $txtRv.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtRv)
    $y += 190

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(12, $y); $lblStatus.Size = New-Object System.Drawing.Size(876, 18)
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray; $lblStatus.Text = 'Fill in the lists, then Preview. Nothing is changed until Apply.'
    $form.Controls.Add($lblStatus)
    $y += 22
    $pb = Add-ReportProgressBar -Form $form -X 12 -Y $y -Width 876 -Height 14
    $y += 24

    $btnPreview = New-Object System.Windows.Forms.Button
    $btnPreview.Text = 'Preview'; $btnPreview.Location = New-Object System.Drawing.Point(12, $y); $btnPreview.Size = New-Object System.Drawing.Size(130, 30)
    $form.Controls.Add($btnPreview)
    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text = 'Apply'; $btnApply.Location = New-Object System.Drawing.Point(150, $y); $btnApply.Size = New-Object System.Drawing.Size(130, 30); $btnApply.Enabled = $false
    $form.Controls.Add($btnApply)
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Location = New-Object System.Drawing.Point(758, $y); $btnClose.Size = New-Object System.Drawing.Size(130, 30)
    $form.Controls.Add($btnClose); $btnClose.Add_Click({ $form.Close() })
    $y += 40

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Location = New-Object System.Drawing.Point(12, $y); $txtOut.Size = New-Object System.Drawing.Size(876, 250)
    $txtOut.Multiline = $true; $txtOut.ReadOnly = $true; $txtOut.WordWrap = $true; $txtOut.ScrollBars = 'Vertical'
    $txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtOut)

    $form | Add-Member -NotePropertyName 'Busy' -NotePropertyValue $false -Force
    $form | Add-Member -NotePropertyName 'Plan' -NotePropertyValue $null  -Force

    $build = {
        $s = Get-TargetSmtp -Text $txtMbx.Text
        if (-not $s) { $lblStatus.Text = 'Enter the calendar mailbox.'; return $null }
        $p = New-ProgressReporter -Label $lblStatus -Bar $pb -Activity 'Preview'
        & $p 'Resolving the calendar...' 10
        $calId = Get-CalendarFolderId -Smtp $s
        & $p 'Reading current permissions...' 30
        $cur = @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' -Parameters @{ Identity = $calId; ErrorAction = 'Stop' } |
                 Where-Object { "$($_.User.DisplayName)" -notin @('Default', 'Anonymous') })
        & $p 'Resolving the Editor list...' 55
        $ed = @(Resolve-NameList -Text $txtEd.Text)
        & $p 'Resolving the Reviewer list...' 80
        $rv = @(Resolve-NameList -Text $txtRv.Text)
        & $p 'Ready.' 100
        return [PSCustomObject]@{ Smtp = $s; CalId = $calId; Current = $cur; Editor = $ed; Reviewer = $rv }
    }

    $btnPreview.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $btnApply.Enabled = $false
            $plan = & $build
            if (-not $plan) { return }
            $form.Plan = $plan
            $bad = @($plan.Editor + $plan.Reviewer | Where-Object { -not $_.Ok })
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine('PREVIEW -- nothing has been changed')
            [void]$sb.AppendLine("  Calendar : $($plan.CalId)")
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine("WOULD REMOVE  ($($plan.Current.Count))")
            foreach ($c in $plan.Current) { [void]$sb.AppendLine('  ' + "$($c.User.DisplayName)".PadRight(40) + (@($c.AccessRights) -join ', ')) }
            [void]$sb.AppendLine("  Default   -> " + $(if ($chkDef.Checked) { 'None' } else { 'unchanged' }))
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine("WOULD GRANT Editor  ($(@($plan.Editor | ? Ok).Count))")
            foreach ($e in ($plan.Editor | Where-Object Ok)) { [void]$sb.AppendLine('  ' + $e.Display.PadRight(40) + $e.Smtp) }
            [void]$sb.AppendLine("WOULD GRANT Reviewer  ($(@($plan.Reviewer | ? Ok).Count))")
            foreach ($e in ($plan.Reviewer | Where-Object Ok)) { [void]$sb.AppendLine('  ' + $e.Display.PadRight(40) + $e.Smtp) }
            if ($bad.Count) {
                [void]$sb.AppendLine('')
                [void]$sb.AppendLine("NOT RESOLVED  ($($bad.Count)) -- fix these before Apply is allowed")
                foreach ($b in $bad) { [void]$sb.AppendLine('  ' + $b.Input.PadRight(40) + $b.Why) }
            }
            $txtOut.Text = $sb.ToString()
            $btnApply.Enabled = ($bad.Count -eq 0)
            $lblStatus.Text = if ($bad.Count) { "$($bad.Count) name(s) did not resolve. Apply is blocked until they do." } else { 'Preview complete. Apply will do exactly this.' }
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Preview'
            $lblStatus.Text = 'Preview failed.'
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $form.Busy = $false
        }
    })

    $btnApply.Add_Click({
        if ($form.Busy) { return }
        $form.Busy = $true
        try {
            $plan = $form.Plan
            if (-not $plan) { $lblStatus.Text = 'Preview first.'; return }
            $ed = @($plan.Editor | Where-Object Ok); $rv = @($plan.Reviewer | Where-Object Ok)
            $a = Show-ToolDialog(
                ("Reset the calendar on $($plan.Smtp)?" + "`r`n`r`n  remove $($plan.Current.Count) existing permission(s)" +
                 "`r`n  grant Editor to $($ed.Count)" + "`r`n  grant Reviewer to $($rv.Count)" +
                 $(if ($chkDef.Checked) { "`r`n  set Default to None" } else { '' }) +
                 "`r`n`r`nPeople not on the lists lose their access."),
                'Confirm', 'OKCancel', 'Warning')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $btnApply.Enabled = $false; $btnPreview.Enabled = $false
            $p = New-ProgressReporter -Label $lblStatus -Bar $pb -Activity 'Applying'
            $total = $plan.Current.Count + $ed.Count + $rv.Count + 1; $i = 0; $ok = 0; $failed = 0; $log = @()

            foreach ($c in $plan.Current) {
                $i++; & $p "Removing $($c.User.DisplayName)..." ([int](100 * $i / $total))
                try { Invoke-EXOCommand -Name 'Remove-MailboxFolderPermission' -Parameters @{ Identity = $plan.CalId; User = "$($c.User.DisplayName)"; Confirm = $false; ErrorAction = 'Stop' } | Out-Null; $ok++; $log += "removed $($c.User.DisplayName)" }
                catch { $failed++; $log += "FAILED remove $($c.User.DisplayName): $($_.Exception.Message)" }
            }
            if ($chkDef.Checked) {
                $i++; & $p 'Setting Default to None...' ([int](100 * $i / $total))
                try {
                    Invoke-EXOCommand -Name 'Set-MailboxFolderPermission' -Parameters @{ Identity = $plan.CalId; User = 'Default';   AccessRights = 'None'; ErrorAction = 'Stop' } | Out-Null
                    Invoke-EXOCommand -Name 'Set-MailboxFolderPermission' -Parameters @{ Identity = $plan.CalId; User = 'Anonymous'; AccessRights = 'None'; ErrorAction = 'Stop' } | Out-Null
                    $ok++; $log += 'Default/Anonymous -> None'
                } catch { $failed++; $log += "FAILED Default: $($_.Exception.Message)" }
            }
            foreach ($pair in @(@($ed, 'Editor'), @($rv, 'Reviewer'))) {
                foreach ($e in $pair[0]) {
                    $i++; & $p "Granting $($pair[1]) to $($e.Display)..." ([int](100 * $i / $total))
                    try {
                        try { Invoke-EXOCommand -Name 'Add-MailboxFolderPermission' -Parameters @{ Identity = $plan.CalId; User = $e.Smtp; AccessRights = $pair[1]; ErrorAction = 'Stop' } | Out-Null }
                        catch { Invoke-EXOCommand -Name 'Set-MailboxFolderPermission' -Parameters @{ Identity = $plan.CalId; User = $e.Smtp; AccessRights = $pair[1]; ErrorAction = 'Stop' } | Out-Null }
                        $ok++; $log += "$($pair[1]) $($e.Display)"
                    } catch { $failed++; $log += "FAILED $($pair[1]) $($e.Display): $($_.Exception.Message)" }
                }
            }

            $final = @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' -Parameters @{ Identity = $plan.CalId; ErrorAction = 'Stop' })
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine($(if ($failed) { 'APPLIED WITH FAILURES' } else { 'APPLIED' }) + "  -- $ok ok, $failed failed")
            [void]$sb.AppendLine('')
            foreach ($l in $log) { [void]$sb.AppendLine('  ' + $l) }
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('CALENDAR NOW')
            foreach ($f in ($final | Sort-Object { @($_.AccessRights) -join ',' }, { "$($_.User.DisplayName)" })) { [void]$sb.AppendLine('  ' + "$($f.User.DisplayName)".PadRight(40) + (@($f.AccessRights) -join ', ')) }
            $txtOut.Text = $sb.ToString()
            $lblStatus.Text = "$ok ok, $failed failed."
            $result = if ($failed -and $ok) { 'Partial' } elseif ($failed) { 'Failed' } else { 'Success' }
            Write-CompletedLog -Action 'ResetCalendarPermissions' -Target (Get-LogTarget $plan.Smtp) -Result $result `
                -Detail "removed $($plan.Current.Count); Editor $($ed.Count); Reviewer $($rv.Count); Default->$(if ($chkDef.Checked) {'None'} else {'unchanged'}); $failed failed"
            $form.Plan = $null
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Apply'
            Write-FailedLog -Action 'ResetCalendarPermissions' -Target "$($txtMbx.Text)" -ErrorRecord $_
        } finally {
            $btnPreview.Enabled = $true
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $form.Busy = $false
        }
    })

    Set-ConsoleTheme -Form $form
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Calendar Reset' | Out-Null
    [void]$form.ShowDialog()
}
#endregion

Show-ResetForm -Seed (Get-TargetSmtp -Text $SmtpAddress)
