#Requires -Version 5.1
<#
    NewUser.ps1
    Enables on-prem Remote Mailboxes (with archive) in Exchange Hybrid, with
    automatic HQ vs Franchise detection.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-BEGIN
# Creates the mailbox for a new starter in the hybrid setup.
#
# It enables a remote mailbox on-premises -- the on-premises object that tells
# the directory this person's mailbox lives in the cloud -- along with their
# archive, and works out whether the account is HQ or Franchise on its own.
#
# The mailbox itself appears in the cloud after the next directory sync, not
# immediately. That gap is normal and is the usual reason a brand-new account
# looks broken for a while.
# TOOLKIT-HELP-END

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Configuration --------------------------------------------------------
$RemoteRoutingDomain = 'expresspros.mail.onmicrosoft.com'
$DefaultFranchiseDC  = 'fran-vm-dc03.franchises.expresspersonnel.com'
#endregion

#region Parsing --------------------------------------------------------------
function Get-AccountList {
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return @() }
    return $Raw -split "[`r`n,;]+" |
           ForEach-Object { $_.Trim() } |
           Where-Object   { $_ -ne '' } |
           Select-Object  -Unique
}
#endregion

#region Actions --------------------------------------------------------------
function Get-OnPremState {
    param([string]$Id, [string]$DC)
    $params = @{ Identity = $Id; ErrorAction = 'Stop' }
    if ($DC) { $params['DomainController'] = $DC }
    Get-RemoteMailbox @params |
        Select-Object Name, PrimarySmtpAddress, RemoteRoutingAddress, ArchiveStatus, ArchiveGuid
}

function Format-ArchiveState {
    param($State)
    $hasGuid = ($State.ArchiveGuid -and $State.ArchiveGuid -ne [Guid]::Empty)
    $status  = if ($State.ArchiveStatus) { $State.ArchiveStatus } else { 'None' }

    if (-not $hasGuid) {
        return ("      Archive              : Not enabled  (ArchiveStatus: {0})" -f $status)
    }
    if ($status -eq 'None') {
        return "      Archive              : Enabled -- provisioning; may not display immediately"
    }
    return ("      Archive              : Enabled  (ArchiveStatus: {0})" -f $status)
}

function Format-OnPremState {
    param([string]$Heading, $State)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("    --- $Heading ---")
    [void]$sb.AppendLine(("      Name                 : {0}" -f $State.Name))
    [void]$sb.AppendLine(("      PrimarySmtpAddress   : {0}" -f $State.PrimarySmtpAddress))
    [void]$sb.AppendLine(("      RemoteRoutingAddress : {0}" -f $State.RemoteRoutingAddress))
    [void]$sb.AppendLine((Format-ArchiveState -State $State))
    return $sb.ToString()
}

function Invoke-MailboxEnable {
    param(
        [string]$Account,
        [string]$DC,
        [string]$Mode
    )
    $out = New-Object System.Text.StringBuilder
    $mb  = $Account.Trim()
    $rra = "$mb@$RemoteRoutingDomain"

    [void]$out.AppendLine("----- $mb  [$Mode] -----")

    try {
        if ($DC) {
            Set-ADServerSettings -PreferredServer $DC -ErrorAction Stop
        } else {
            # GUARD: forest-wide scope is deliberate. Do not narrow.
            Set-ADServerSettings -PreferredServer $null -ViewEntireForest $true -ErrorAction Stop
        }
    } catch {
        [void]$out.AppendLine("  Could not set AD server scope: $($_.Exception.Message)")
        return $out.ToString()
    }

    $before = $null
    try { $before = Get-OnPremState -Id $mb -DC $DC } catch { $before = $null }
    if ($before) {
        [void]$out.AppendLine("  SKIPPED -- already has a remote mailbox.")
        [void]$out.Append((Format-OnPremState -Heading 'EXISTING (on-prem)' -State $before))
        return $out.ToString()
    }

    $enableParams = @{
        Identity             = $mb
        Alias                = $mb
        RemoteRoutingAddress = $rra
        Confirm              = $false
        ErrorAction          = 'Stop'
    }
    if ($DC) { $enableParams['DomainController'] = $DC }
    try {
        Enable-RemoteMailbox @enableParams | Out-Null
        [void]$out.AppendLine("  Mailbox enabled.  ($rra)")
    } catch {
        $msg = $_.Exception.Message
        if ($msg -match 'does not support recipients of this type' -or $msg -match 'already') {
            [void]$out.AppendLine("  SKIPPED -- already enabled, or not eligible for a remote mailbox.")
        } else {
            [void]$out.AppendLine("  FAILED to enable: $msg")
        }
        return $out.ToString()
    }

    $archiveParams = @{ Identity = $mb; Archive = $true; ErrorAction = 'Stop' }
    if ($DC) { $archiveParams['DomainController'] = $DC }
    $archiveOk = $false
    try {
        Enable-RemoteMailbox @archiveParams | Out-Null
        $archiveOk = $true
        [void]$out.AppendLine("  Archive enabled.")
    } catch {
        [void]$out.AppendLine("  WARNING -- mailbox enabled but archive failed: $($_.Exception.Message)")
    }

    try {
        $after = Get-OnPremState -Id $mb -DC $DC
        [void]$out.AppendLine("  Verified.")
        [void]$out.Append((Format-OnPremState -Heading 'AFTER (on-prem)' -State $after))
        $logTask = if ($archiveOk) { "Enabled RemoteMailbox $mb ($Mode)" }
                   else            { "Enabled RemoteMailbox $mb ($Mode) -- ARCHIVE FAILED" }
        Write-CompletedLog -Task $logTask
        [void]$out.AppendLine("  Logged to CompletedTasks.csv")
    } catch {
        [void]$out.AppendLine("  Could not verify after enable: $($_.Exception.Message)")
        [void]$out.AppendLine("  NOT logged (verification failed).")
    }
    return $out.ToString()
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Enable Remote Mailbox'
$form.Size            = New-Object System.Drawing.Size(760, 700)
$form.StartPosition   = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblWarn           = New-Object System.Windows.Forms.Label
$lblWarn.Text      = 'Connect to on-prem Exchange before running.'
$lblWarn.Location  = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size      = New-Object System.Drawing.Size(430, 20)
$lblWarn.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

$lblCtx           = New-Object System.Windows.Forms.Label
$lblCtx.Text      = "HQ forest: $($script:HQDomain)"
$lblCtx.Location  = New-Object System.Drawing.Point(450, 10)
$lblCtx.Size      = New-Object System.Drawing.Size(282, 20)
$lblCtx.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblCtx.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblCtx)

$lblDC          = New-Object System.Windows.Forms.Label
$lblDC.Text     = 'Franchise DC (used only for accounts found in Franchise):'
$lblDC.Location = New-Object System.Drawing.Point(12, 40)
$lblDC.AutoSize = $true
$form.Controls.Add($lblDC)

$cmbDC = New-Object System.Windows.Forms.ComboBox
$cmbDC.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbDC.Location = New-Object System.Drawing.Point(360, 37)
$cmbDC.Size     = New-Object System.Drawing.Size(372, 22)
$script:FranchiseDCs | ForEach-Object { [void]$cmbDC.Items.Add($_) }
$idx = $cmbDC.Items.IndexOf($DefaultFranchiseDC)
if ($idx -ge 0) { $cmbDC.SelectedIndex = $idx } else { $cmbDC.SelectedIndex = 0 }
$lblDC.Visible = $false
$cmbDC.Visible = $false
# DC selector hidden 2026-08-03: the toolkit binds to the default franchise
# controller from Common.ps1 and nobody was choosing another one. The control
# is kept rather than deleted so it still supplies that default to the code
# below -- .Visible is the whole change, and reversing it is one word.
$form.Controls.Add($cmbDC)

$lblAcc          = New-Object System.Windows.Forms.Label
$lblAcc.Text     = 'Account name(s) -- one per line (sAMAccountName). HQ and Franchise may be mixed; each is detected automatically.'
$lblAcc.Location = New-Object System.Drawing.Point(12, 70)
$lblAcc.AutoSize = $true
$form.Controls.Add($lblAcc)

$txtAcc               = New-Object System.Windows.Forms.TextBox
$txtAcc.Multiline     = $true
$txtAcc.ScrollBars    = 'Vertical'
$txtAcc.AcceptsReturn = $true
$txtAcc.Location      = New-Object System.Drawing.Point(12, 90)
$txtAcc.Size          = New-Object System.Drawing.Size(720, 160)
$txtAcc.Font          = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtAcc)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run'
$btnRun.Location = New-Object System.Drawing.Point(12, 260)
$btnRun.Size     = New-Object System.Drawing.Size(100, 28)
$btnRun.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRun)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text     = 'Clear'
$btnClear.Location = New-Object System.Drawing.Point(120, 260)
$btnClear.Size     = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClear)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(642, 260)
$btnClose.Size     = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClose)

$lblResult          = New-Object System.Windows.Forms.Label
$lblResult.Text     = 'Result:'
$lblResult.Location = New-Object System.Drawing.Point(12, 296)
$lblResult.Size     = New-Object System.Drawing.Size(720, 20)
$lblResult.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblResult)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Both'
$txtOut.WordWrap   = $false
$txtOut.Location   = New-Object System.Drawing.Point(12, 318)
$txtOut.Size       = New-Object System.Drawing.Size(720, 330)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.BackColor  = [System.Drawing.Color]::WhiteSmoke
$form.Controls.Add($txtOut)

$btnRun.Add_Click({
    $btnRun.Enabled   = $false
    $btnClose.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $accounts = Get-AccountList -Raw $txtAcc.Text
        if ($accounts.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show(
                'No account names entered.', 'Nothing to do',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }

        $franDC = $cmbDC.SelectedItem

        $lblResult.Text = 'Resolving accounts...'
        $form.Refresh()

        $resolved    = @()
        $missing     = @()
        $probeErrors = @()
        foreach ($a in $accounts) {
            $envInfo = Find-UserEnvironmentBySam -SamAccountName $a
            if ($envInfo.Errors.Count -gt 0) { $probeErrors += $envInfo.Errors }
            if ($envInfo.Environment) {
                $writeDC = ''
                if ($envInfo.Environment -eq 'Franchise') { $writeDC = $franDC }
                $resolved += [PSCustomObject]@{
                    Account     = $a
                    Environment = $envInfo.Environment
                    DC          = $writeDC
                }
            } else {
                $missing += $a
            }
        }

        if ($missing.Count -gt 0) {
            $extra = ''
            if ($resolved.Count -gt 0) {
                $extra = "`r`n`r`nThe other $($resolved.Count) account(s) can still be processed."
            }
            [void][System.Windows.Forms.MessageBox]::Show(
                ("Account does not yet exist:`r`n`r`n{0}`r`n`r`nNot found in HQ or Franchise AD. Create the account first, then run this again.{1}" -f `
                    ($missing -join "`r`n"), $extra),
                'Account does not yet exist',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
        }

        if ($resolved.Count -eq 0) {
            $lblResult.Text = 'Nothing to do -- no account resolved.'
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine('No account was found in HQ or Franchise AD:')
            foreach ($m in $missing) { [void]$sb.AppendLine("  $m") }
            if ($probeErrors.Count -gt 0) {
                [void]$sb.AppendLine('')
                [void]$sb.AppendLine('--- Lookup errors ---')
                foreach ($e in ($probeErrors | Select-Object -Unique)) { [void]$sb.AppendLine($e) }
            }
            $txtOut.Text = $sb.ToString()
            return
        }

        $hqN   = @($resolved | Where-Object { $_.Environment -eq 'HQ' }).Count
        $franN = @($resolved | Where-Object { $_.Environment -eq 'Franchise' }).Count
        $bits  = @()
        if ($hqN)   { $bits += "$hqN HQ" }
        if ($franN) { $bits += "$franN Franchise (on $franDC)" }

        $preview = ($resolved | ForEach-Object { "  $($_.Account)  ->  $($_.Environment)" }) -join "`r`n"
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            ("Enable remote mailbox + archive for {0} account(s)?`r`n`r`n{1}`r`n`r`n{2}" -f `
                $resolved.Count, $preview, ($bits -join '   |   ')),
            'Confirm',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
            $lblResult.Text = 'Cancelled.'
            $txtOut.Text = 'Cancelled. No changes made.'
            return
        }

        $out = New-Object System.Text.StringBuilder
        [void]$out.AppendLine("===== $(Get-Date -Format 'yyyy-MM-dd HH:mm') =====")
        if ($missing.Count -gt 0) {
            [void]$out.AppendLine('')
            [void]$out.AppendLine('DOES NOT EXIST YET (skipped, nothing attempted):')
            foreach ($m in $missing) { [void]$out.AppendLine("  $m") }
        }

        $n = 0
        foreach ($grp in @('HQ', 'Franchise')) {
            $set = @($resolved | Where-Object { $_.Environment -eq $grp })
            if ($set.Count -eq 0) { continue }
            $onDC = ''
            if ($grp -eq 'Franchise') { $onDC = " on $franDC" }
            [void]$out.AppendLine('')
            [void]$out.AppendLine("===== $grp batch ($($set.Count) account(s))$onDC =====")
            foreach ($r in $set) {
                $n++
                $lblResult.Text = "Processing $n of $($resolved.Count): $($r.Account) [$($r.Environment)] ..."
                $form.Refresh()
                [void]$out.Append((Invoke-MailboxEnable -Account $r.Account -DC $r.DC -Mode $r.Environment))
                [void]$out.AppendLine('')
            }
        }

        if ($probeErrors.Count -gt 0) {
            [void]$out.AppendLine('--- Lookup errors (the accounts above were still processed) ---')
            foreach ($e in ($probeErrors | Select-Object -Unique)) { [void]$out.AppendLine($e) }
        }

        $lblResult.Text = "Done -- $($resolved.Count) processed, $($missing.Count) missing."
        $txtOut.Text = $out.ToString()
    }
    catch {
        $lblResult.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Enable Remote Mailbox'
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $btnRun.Enabled   = $true
        $btnClose.Enabled = $true
    }
})

$btnClear.Add_Click({
    $txtAcc.Text = ''
    $txtOut.Text = ''
    $lblResult.Text = 'Result:'
})

$btnClose.Add_Click({ $form.Close() })

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'

# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled, and BringToFront'd inside the helper because WinForms
# draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'New User' | Out-Null

[void]$form.ShowDialog()
$form.Dispose()
#endregion
