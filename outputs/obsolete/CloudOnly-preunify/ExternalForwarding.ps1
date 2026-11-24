#Requires -Version 5.1
<#
    ExternalForwarding.ps1
    Finds mail being forwarded outside the tenant, from both mailbox settings and user inbox rules, and exports to CSV. Author : Glenn Armbruster

    Reasoning: ScriptNotes/ExternalForwarding.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-07-31
#>
# TOOLKIT-HELP-BEGIN
# Finds mail leaving the organisation, from BOTH places it can be set up:
#
#   * mailbox forwarding, configured by an admin or by the user
#   * inbox rules, which users create themselves and which are the ones nobody
#     thinks to check
#
# The classic sign of a compromised mailbox is a quiet rule forwarding
# everything to an outside address. This is the report that finds it.
#
# Writes a CSV. Read-only.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Configuration --------------------------------------------------------
$ReportDir = Get-ReportPath

$AllowlistedDomains = @()
#endregion

#region Data -----------------------------------------------------------------
function Test-ExternalDomain {
    param([string]$Domain, [string[]]$Accepted)
    if (-not $Domain) { return $false }
    return (-not ($Accepted -contains $Domain.ToLower()))
}

function Get-MailboxForwarding {
    param([string[]]$Accepted, [scriptblock]$Progress)

    $cmd = if (Get-Command 'Get-EXOMailbox' -ErrorAction SilentlyContinue) {
               'Get-EXOMailbox'
           } else { 'Get-Mailbox' }

    $props = @('PrimarySmtpAddress', 'DisplayName', 'RecipientTypeDetails',
               'ForwardingAddress', 'ForwardingSmtpAddress',
               'DeliverToMailboxAndForward')

    & $Progress 'Reading mailbox-level forwarding...' 5
    $all = @(Invoke-EXOBulkQuery -Cmdlet $cmd -Properties $props -Progress $Progress)

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($m in $all) {
        foreach ($pair in @(
            @{ Val = "$($m.ForwardingSmtpAddress)"; Kind = 'Mailbox ForwardingSmtpAddress' },
            @{ Val = "$($m.ForwardingAddress)";     Kind = 'Mailbox ForwardingAddress' }
        )) {
            if (-not $pair.Val) { continue }
            $dom = Get-SmtpDomain $pair.Val
            if (-not $dom) { continue }          # unresolvable / legacy DN
            if (-not (Test-ExternalDomain -Domain $dom -Accepted $Accepted)) { continue }

            $rows.Add([PSCustomObject][ordered]@{
                PrimarySmtpAddress = "$($m.PrimarySmtpAddress)"
                DisplayName        = "$($m.DisplayName)"
                RecipientType      = "$($m.RecipientTypeDetails)"
                Source             = $pair.Kind
                RuleName           = ''
                RuleEnabled        = ''
                Action             = 'Forward'
                Destination        = $pair.Val
                DestinationDomain  = $dom
                KeepsCopy          = [bool]$m.DeliverToMailboxAndForward
                Allowlisted        = ($AllowlistedDomains -contains $dom)
            })
        }
    }
    return $rows
}

function Get-InboxRuleForwarding {
    param(
        [object[]]$Mailboxes,
        [string[]]$Accepted,
        [scriptblock]$Progress
    )
    $rows   = New-Object System.Collections.Generic.List[object]
    $failed = @()
    $n = 0
    $total = @($Mailboxes).Count

    foreach ($mb in $Mailboxes) {
        $n++
        if ($n % 10 -eq 0 -or $n -eq $total) {
            $pct = 20 + [int](75.0 * $n / [Math]::Max(1, $total))
            & $Progress "Inbox rules: $n of $total -- $($rows.Count) external found..." $pct
        }
        $smtp = "$($mb.PrimarySmtpAddress)"
        if (-not $smtp) { continue }

        try {
            $rules = @(Invoke-EXOCommand -Name 'Get-InboxRule' `
                        -Parameters @{ Mailbox = $smtp; ErrorAction = 'Stop' })
        } catch {
            $failed += $smtp
            continue
        }

        foreach ($r in $rules) {
            foreach ($act in @(
                @{ Prop = 'ForwardTo';             Label = 'Forward' },
                @{ Prop = 'ForwardAsAttachmentTo'; Label = 'Forward as attachment' },
                @{ Prop = 'RedirectTo';            Label = 'Redirect (no copy kept)' }
            )) {
                $targets = @($r.($act.Prop))
                foreach ($t in $targets) {
                    if (-not $t) { continue }
                    $dom = Get-SmtpDomain "$t"
                    if (-not $dom) { continue }
                    if (-not (Test-ExternalDomain -Domain $dom -Accepted $Accepted)) { continue }

                    $rows.Add([PSCustomObject][ordered]@{
                        PrimarySmtpAddress = $smtp
                        DisplayName        = "$($mb.DisplayName)"
                        RecipientType      = "$($mb.RecipientTypeDetails)"
                        Source             = 'Inbox rule'
                        RuleName           = "$($r.Name)"
                        RuleEnabled        = [bool]$r.Enabled
                        Action             = $act.Label
                        Destination        = "$t"
                        DestinationDomain  = $dom
                        KeepsCopy          = ($act.Prop -ne 'RedirectTo')
                        Allowlisted        = ($AllowlistedDomains -contains $dom)
                    })
                }
            }
        }
    }
    return [PSCustomObject]@{ Rows = $rows; Failed = $failed }
}

function Get-ForwardingReport {
    param([scriptblock]$Progress, [switch]$ScanInboxRules)

    & $Progress 'Reading accepted domains...' 2
    $accepted = Get-AcceptedDomainList
    if ($accepted.Count -eq 0) {
        throw "Could not read accepted domains. Without them there is no way to tell external from internal, so the report would be meaningless."
    }

    $rows   = New-Object System.Collections.Generic.List[object]
    $failed = @()

    foreach ($r in (Get-MailboxForwarding -Accepted $accepted -Progress $Progress)) {
        $rows.Add($r)
    }

    $scanned = 0
    if ($ScanInboxRules) {
        & $Progress 'Listing mailboxes for inbox-rule scan...' 15
        $cmd = if (Get-Command 'Get-EXOMailbox' -ErrorAction SilentlyContinue) {
                   'Get-EXOMailbox'
               } else { 'Get-Mailbox' }
        $mbs = @(Invoke-EXOBulkQuery -Cmdlet $cmd `
                    -Properties @('PrimarySmtpAddress', 'DisplayName', 'RecipientTypeDetails') `
                    -Progress $Progress)
        $scanned = $mbs.Count
        $res = Get-InboxRuleForwarding -Mailboxes $mbs -Accepted $accepted -Progress $Progress
        foreach ($r in $res.Rows) { $rows.Add($r) }
        $failed = $res.Failed
    }

    $sorted = @($rows | Sort-Object `
        @{ Expression = { $_.Allowlisted } },
        @{ Expression = { $_.KeepsCopy } },
        @{ Expression = { $_.DestinationDomain } },
        @{ Expression = { $_.PrimarySmtpAddress } })

    return [PSCustomObject]@{
        Rows            = $sorted
        AcceptedDomains = $accepted
        ScannedRules    = $ScanInboxRules.IsPresent
        MailboxesScanned = $scanned
        Failed          = $failed
    }
}

function Format-ReportSummary {
    param($Result)
    $rows = @($Result.Rows)
    $sb = New-Object System.Text.StringBuilder

    $live = @($rows | Where-Object { -not $_.Allowlisted })
    $noCopy = @($live | Where-Object { -not $_.KeepsCopy })
    $ruleBased = @($live | Where-Object { $_.Source -eq 'Inbox rule' })

    [void]$sb.AppendLine("External forwarding found : $($rows.Count)")
    [void]$sb.AppendLine("  Not allowlisted         : $($live.Count)")
    [void]$sb.AppendLine("  No copy kept in mailbox : $($noCopy.Count)   <-- highest priority")
    [void]$sb.AppendLine("  From inbox rules        : $($ruleBased.Count)")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("Distinct mailboxes affected: $(@($live | Select-Object -ExpandProperty PrimarySmtpAddress -Unique).Count)")
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine('By source:')
    $live | Group-Object Source | Sort-Object Count -Descending | ForEach-Object {
        [void]$sb.AppendLine(("  {0,-32}{1,6}" -f $_.Name, $_.Count))
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Top destination domains:')
    $live | Group-Object DestinationDomain | Sort-Object Count -Descending |
        Select-Object -First 15 | ForEach-Object {
            [void]$sb.AppendLine(("  {0,-32}{1,6}" -f $_.Name, $_.Count))
        }

    [void]$sb.AppendLine('')
    if ($Result.ScannedRules) {
        [void]$sb.AppendLine("Inbox rules scanned across $($Result.MailboxesScanned) mailbox(es).")
    } else {
        [void]$sb.AppendLine('*** Inbox rules were NOT scanned. ***')
        [void]$sb.AppendLine('  Mailbox-level forwarding only. Rule-based forwarding is the more')
        [void]$sb.AppendLine('  common exfiltration route -- tick the box and re-run for full coverage.')
    }
    [void]$sb.AppendLine("Accepted domains treated as internal: $($Result.AcceptedDomains.Count)")

    if ($Result.Failed -and $Result.Failed.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("*** INCOMPLETE -- rules could not be read for $($Result.Failed.Count) mailbox(es): ***")
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
$form.Text            = 'Report -- External Forwarding'
$form.Size            = New-Object System.Drawing.Size(760, 676)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Finds mail forwarded outside the tenant, from mailbox settings and from user inbox rules.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblInfo)

$lblTenant           = New-Object System.Windows.Forms.Label
$lblTenant.Tag       = 'NoTheme'
$lblTenant.Location  = New-Object System.Drawing.Point(12, 32)
$lblTenant.Size      = New-Object System.Drawing.Size(720, 18)
$lblTenant.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblTenant)

$chkRules          = New-Object System.Windows.Forms.CheckBox
$chkRules.Text     = 'Also scan inbox rules (slow -- one call per mailbox, allow 30-60 min)'
$chkRules.Location = New-Object System.Drawing.Point(12, 56)
$chkRules.Size     = New-Object System.Drawing.Size(500, 20)
$form.Controls.Add($chkRules)

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

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'ExternalForwarding' `
                -X 12 -Y 118 -Width 720 -StatusLabel $lblStatus


$btnOpen.Add_Click({ Open-ReportFolder -Path $ReportDir })

$btnRun.Add_Click({
    if ($chkRules.Checked) {
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "Scanning inbox rules queries every mailbox individually and can take 30-60 minutes on a tenant this size.`n`nThe window will stay responsive and show progress. Continue?",
            'Long-running scan', 'YesNo', 'Question')
        if ($ans -ne 'Yes') { return }
    }

    $btnRun.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = ''
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar `
                        -Activity 'External forwarding report'
        $report = Get-ForwardingReport -Progress $progress -ScanInboxRules:$chkRules.Checked

        Complete-ProgressReporter -Activity 'External forwarding report' -Bar $pbar

        if (@($report.Rows).Count -eq 0) {
            $lblStatus.Text = 'No external forwarding found -- nothing written.'
            $txtOut.Text = (Format-ReportSummary -Result $report) +
                           [Environment]::NewLine +
                           'No CSV was written because there were no rows.'
            return
        }

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir "ExternalForwarding_$stamp.csv"
        $report.Rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $files.SelectPath($csv)
        $lblStatus.Text  = "Done. $(@($report.Rows).Count) row(s) written."

        $txtOut.Text = (Format-ReportSummary -Result $report) +
                       [Environment]::NewLine +
                       'Saved to:' + [Environment]::NewLine + "  $csv"

        $note = @()
        if (-not $report.ScannedRules) { $note += 'mailbox-level only' }
        if ($report.Failed.Count)      { $note += "$($report.Failed.Count) unreadable" }
        # 'PARTIAL' is the keyword the log is scanned for.
        $suffix = if ($note.Count) { " -- PARTIAL ($($note -join ', '))" } else { '' }
        # PARTIAL is the keyword the log is scanned for; Result is the same
        # fact in a column a filter can see.
        if ($suffix) { $logResult = 'Partial' } else { $logResult = 'Success' }
        Write-CompletedLog -Task "External forwarding report -- $(@($report.Rows).Count) findings$suffix" -Result $logResult
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'External forwarding report'
    } finally {
        Complete-ProgressReporter -Activity 'External forwarding report'
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
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'External Forwarding' | Out-Null
[void]$form.ShowDialog()
#endregion
