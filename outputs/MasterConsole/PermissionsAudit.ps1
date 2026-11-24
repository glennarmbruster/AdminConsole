#Requires -Version 5.1
<#
    PermissionsAudit.ps1
    Reports Full Access, Send As and Send on Behalf delegations across every mailbox, and exports to CSV. Author : Glenn Armbruster

    Reasoning: ScriptNotes/PermissionsAudit.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-07-31
#>
# TOOLKIT-HELP-BEGIN
# Reports every mailbox delegation across the organisation -- Full Access, Send
# As and Send on Behalf. Writes a CSV.
#
# Answers "who can get into what", which is a question that only ever gets asked
# under time pressure. Access granted for a good reason years ago is still
# access.
#
# Read-only. It reports; it does not remove anything.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Configuration --------------------------------------------------------
$ReportDir = Get-ReportPath

$SystemPrincipals = @(
    'NT AUTHORITY\SELF'
    'NT AUTHORITY\SYSTEM'
    'S-1-5-'
    'Discovery Management'
    'Organization Management'
    'Exchange Servers'
    'Exchange Trusted Subsystem'
    'Public Folder Management'
    'Managed Availability Servers'
    'Delegated Setup'
    'Hygiene Management'
)
#endregion

#region Data -----------------------------------------------------------------
function Test-SystemPrincipal {
    param([string]$User)
    if (-not $User) { return $true }
    foreach ($p in $SystemPrincipals) {
        if ($User -like "*$p*") { return $true }
    }
    return $false
}

function Get-PermissionsAudit {
    param(
        [scriptblock]$Progress,
        [switch]$IncludeInherited
    )

    & $Progress 'Listing mailboxes...' 2
    $mbCmd = if (Get-Command 'Get-EXOMailbox' -ErrorAction SilentlyContinue) {
                 'Get-EXOMailbox'
             } else { 'Get-Mailbox' }

    $mbs = @(Invoke-EXOBulkQuery -Cmdlet $mbCmd `
                -Properties @('PrimarySmtpAddress', 'DisplayName',
                              'RecipientTypeDetails', 'GrantSendOnBehalfTo') `
                -Progress $Progress)
    if ($mbs.Count -eq 0) { throw 'No mailboxes returned.' }

    $permCmd = if (Get-Command 'Get-EXOMailboxPermission' -ErrorAction SilentlyContinue) {
                   'Get-EXOMailboxPermission'
               } else { 'Get-MailboxPermission' }
    $saCmd   = if (Get-Command 'Get-EXORecipientPermission' -ErrorAction SilentlyContinue) {
                   'Get-EXORecipientPermission'
               } else { 'Get-RecipientPermission' }

    $rows   = New-Object System.Collections.Generic.List[object]
    $failed = @()
    $total  = $mbs.Count
    $n = 0

    foreach ($mb in $mbs) {
        $n++
        if ($n % 10 -eq 0 -or $n -eq $total) {
            & $Progress "Reading permissions: $n of $total -- $($rows.Count) delegations found..." (5 + [int](92.0 * $n / $total))
        }
        $smtp = "$($mb.PrimarySmtpAddress)"
        if (-not $smtp) { continue }

        # --- Full Access -----------------------------------------------------
        try {
            $perms = @(Invoke-EXOCommand -Name $permCmd `
                        -Parameters @{ Identity = $smtp; ErrorAction = 'Stop' })
            foreach ($p in $perms) {
                $user = "$($p.User)"
                if (Test-SystemPrincipal -User $user) { continue }
                if ($p.IsInherited -and -not $IncludeInherited) { continue }
                if ("$($p.Deny)" -eq 'True') { continue }

                $rights = @($p.AccessRights) -join ','
                if ($rights -notmatch 'FullAccess') { continue }

                $rows.Add([PSCustomObject][ordered]@{
                    Mailbox           = $smtp
                    MailboxDisplay    = "$($mb.DisplayName)"
                    MailboxType       = "$($mb.RecipientTypeDetails)"
                    Permission        = 'Full Access'
                    Delegate          = $user
                    AccessRights      = $rights
                    Inherited         = [bool]$p.IsInherited
                })
            }
        } catch {
            $failed += "$smtp (Full Access)"
        }

        # --- Send As ---------------------------------------------------------
        try {
            $sa = @(Invoke-EXOCommand -Name $saCmd `
                        -Parameters @{ Identity = $smtp; ErrorAction = 'Stop' })
            foreach ($p in $sa) {
                $user = "$($p.Trustee)"
                if (Test-SystemPrincipal -User $user) { continue }
                if ("$($p.AccessControlType)" -eq 'Deny') { continue }
                $rights = @($p.AccessRights) -join ','
                if ($rights -notmatch 'SendAs') { continue }

                $rows.Add([PSCustomObject][ordered]@{
                    Mailbox           = $smtp
                    MailboxDisplay    = "$($mb.DisplayName)"
                    MailboxType       = "$($mb.RecipientTypeDetails)"
                    Permission        = 'Send As'
                    Delegate          = $user
                    AccessRights      = $rights
                    Inherited         = $false
                })
            }
        } catch {
            $failed += "$smtp (Send As)"
        }

        # --- Send on Behalf --------------------------------------------------
        foreach ($t in @($mb.GrantSendOnBehalfTo)) {
            if (-not $t) { continue }
            $rows.Add([PSCustomObject][ordered]@{
                Mailbox           = $smtp
                MailboxDisplay    = "$($mb.DisplayName)"
                MailboxType       = "$($mb.RecipientTypeDetails)"
                Permission        = 'Send on Behalf'
                Delegate          = "$t"
                AccessRights      = 'SendOnBehalf'
                Inherited         = $false
            })
        }
    }

    $sorted = @($rows | Sort-Object Delegate, Mailbox, Permission)

    return [PSCustomObject]@{
        Rows             = $sorted
        MailboxesScanned = $total
        Failed           = $failed
        IncludedInherited = [bool]$IncludeInherited
    }
}

function Format-ReportSummary {
    param($Result)
    $rows = @($Result.Rows)
    $sb = New-Object System.Text.StringBuilder

    [void]$sb.AppendLine("Mailboxes scanned    : $($Result.MailboxesScanned)")
    [void]$sb.AppendLine("Delegations found    : $($rows.Count)")
    [void]$sb.AppendLine("Mailboxes with any   : $(@($rows | Select-Object -ExpandProperty Mailbox -Unique).Count)")
    [void]$sb.AppendLine("Distinct delegates   : $(@($rows | Select-Object -ExpandProperty Delegate -Unique).Count)")
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine('By permission:')
    $rows | Group-Object Permission | Sort-Object Count -Descending | ForEach-Object {
        [void]$sb.AppendLine(("  {0,-18}{1,6}" -f $_.Name, $_.Count))
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('By mailbox type:')
    $rows | Group-Object MailboxType | Sort-Object Count -Descending | ForEach-Object {
        [void]$sb.AppendLine(("  {0,-22}{1,6}" -f $_.Name, $_.Count))
    }

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Widest access -- delegates holding the most mailboxes:')
    $rows | Group-Object Delegate | Sort-Object Count -Descending |
        Select-Object -First 15 | ForEach-Object {
            $mbCount = @($_.Group | Select-Object -ExpandProperty Mailbox -Unique).Count
            [void]$sb.AppendLine(("  {0,-42}{1,5} mailbox(es)" -f $_.Name, $mbCount))
        }

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Most delegated mailboxes:')
    $rows | Group-Object Mailbox | Sort-Object Count -Descending |
        Select-Object -First 10 | ForEach-Object {
            $dCount = @($_.Group | Select-Object -ExpandProperty Delegate -Unique).Count
            [void]$sb.AppendLine(("  {0,-42}{1,5} delegate(s)" -f $_.Name, $dCount))
        }

    $userMbx = @($rows | Where-Object { $_.MailboxType -eq 'UserMailbox' })
    if ($userMbx.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("Delegations on personal (user) mailboxes: $($userMbx.Count)")
        [void]$sb.AppendLine('  Shared mailboxes are meant to be delegated. Access to somebody')
        [void]$sb.AppendLine("  else's personal mailbox is the category worth reviewing.")
    }

    [void]$sb.AppendLine('')
    if (-not $Result.IncludedInherited) {
        [void]$sb.AppendLine('Inherited permissions were excluded (they come from admin roles, not')
        [void]$sb.AppendLine('  from individual grants). Tick the box to include them.')
    }
    [void]$sb.AppendLine('Not covered: folder-level permissions (calendar, inbox subfolders) --')
    [void]$sb.AppendLine('  those are per-folder and would multiply the runtime. Use the Calendar')
    [void]$sb.AppendLine('  Permissions tool for a specific mailbox.')

    if ($Result.Failed -and $Result.Failed.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("*** INCOMPLETE -- $($Result.Failed.Count) read(s) failed: ***")
        foreach ($f in ($Result.Failed | Select-Object -First 20)) { [void]$sb.AppendLine("  $f") }
        if ($Result.Failed.Count -gt 20) {
            [void]$sb.AppendLine("  ... and $($Result.Failed.Count - 20) more")
        }
    }
    return $sb.ToString()
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Report -- Mailbox Permissions Audit'
$form.Size            = New-Object System.Drawing.Size(760, 676)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Every Full Access, Send As and Send on Behalf delegation across the tenant.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblInfo)

$lblTenant           = New-Object System.Windows.Forms.Label
$lblTenant.Tag       = 'NoTheme'
$lblTenant.Location  = New-Object System.Drawing.Point(12, 32)
$lblTenant.Size      = New-Object System.Drawing.Size(720, 18)
$lblTenant.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblTenant)

$chkInherited          = New-Object System.Windows.Forms.CheckBox
$chkInherited.Text     = 'Include inherited permissions (from admin roles)'
$chkInherited.Location = New-Object System.Drawing.Point(12, 56)
$chkInherited.Size     = New-Object System.Drawing.Size(340, 20)
$form.Controls.Add($chkInherited)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run Report'
$btnRun.Location = New-Object System.Drawing.Point(12, 82)
$btnRun.Size     = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnRun)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Reports Folder'
$btnOpen.Location = New-Object System.Drawing.Point(140, 82)
$btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnOpen)


$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(622, 82)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 156)
$lblStatus.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 176 -Width 720 -Height 12

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Location   = New-Object System.Drawing.Point(12, 194)
$txtOut.Size       = New-Object System.Drawing.Size(720, 420)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'PermissionsAudit' `
                -X 12 -Y 118 -Width 720 -StatusLabel $lblStatus


$btnOpen.Add_Click({ Open-ReportFolder -Path $ReportDir })

$btnRun.Add_Click({
    $ans = [System.Windows.Forms.MessageBox]::Show(
        "This reads permissions one mailbox at a time -- there is no bulk form for these cmdlets -- so allow 30-60 minutes on a tenant this size.`n`nThe window stays responsive and shows progress. Continue?",
        'Long-running scan', 'YesNo', 'Question')
    if ($ans -ne 'Yes') { return }

    $btnRun.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = ''
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar `
                        -Activity 'Mailbox permissions audit'
        $report = Get-PermissionsAudit -Progress $progress -IncludeInherited:$chkInherited.Checked

        Complete-ProgressReporter -Activity 'Mailbox permissions audit' -Bar $pbar

        if (@($report.Rows).Count -eq 0) {
            $lblStatus.Text = 'No delegations found -- nothing written.'
            $txtOut.Text = (Format-ReportSummary -Result $report) +
                           [Environment]::NewLine +
                           'No CSV was written because there were no rows.'
            return
        }

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir "PermissionsAudit_$stamp.csv"
        $report.Rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $files.SelectPath($csv)
        $lblStatus.Text  = "Done. $(@($report.Rows).Count) row(s) written."

        $txtOut.Text = (Format-ReportSummary -Result $report) +
                       [Environment]::NewLine +
                       'Saved to:' + [Environment]::NewLine + "  $csv"

        # 'PARTIAL' is the keyword the log is scanned for.
        $note = if ($report.Failed.Count) { " -- PARTIAL ($($report.Failed.Count) unreadable)" } else { '' }
        # PARTIAL is the keyword the log is scanned for; Result is the same
        # fact in a column a filter can see.
        if ($note) { $logResult = 'Partial' } else { $logResult = 'Success' }
        Write-CompletedLog -Task "Permissions audit -- $(@($report.Rows).Count) delegations$note" -Result $logResult
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mailbox permissions audit'
    } finally {
        Complete-ProgressReporter -Activity 'Mailbox permissions audit'
        $form.Cursor    = $prev
        $btnRun.Enabled = $true
    }
})

$d = Get-ConnectedTenantDomain
if ($d) {
    $lblTenant.Text      = "Tenant: $d"
    $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(0, 128, 0)
} else {
    $lblTenant.Text      = 'Tenant: not connected -- connect to Exchange Online first.'
    $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(200, 0, 0)
}

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled, and BringToFront'd inside the helper because WinForms
# draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Permissions Audit' | Out-Null

[void]$form.ShowDialog()
#endregion
