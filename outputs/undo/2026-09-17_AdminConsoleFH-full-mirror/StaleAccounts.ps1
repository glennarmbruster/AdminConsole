#Requires -Version 5.1
<#
    StaleAccounts.ps1
    Finds enabled, licensed accounts that have not signed in recently, and exports to CSV. Author : Glenn Armbruster

    Reasoning: ScriptNotes/StaleAccounts.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-07-31
#>
# TOOLKIT-HELP-BEGIN
# Finds accounts that are enabled and licensed but have not signed in for a
# while.
#
# Two reasons to care: a licence being paid for and not used, and an account
# nobody is watching that still works. Writes a CSV.
#
# Not every result is dead -- shared and service accounts show up here
# legitimately. Treat it as a list to review, not a list to disable.
#
# Read-only.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
#
# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost: the window flashes up, work may have started, and
# the operator has seen a tool they were not meant to be in.
if (-not (Test-ConsoleLaunch)) { return }

#region Configuration --------------------------------------------------------
$ReportDir = Get-ReportPath

$Buckets = @(
    [PSCustomObject]@{ Label = '90-180 days';   Min = 90;   Max = 180 }
    [PSCustomObject]@{ Label = '180-365 days';  Min = 180;  Max = 365 }
    [PSCustomObject]@{ Label = '1-2 years';     Min = 365;  Max = 730 }
    [PSCustomObject]@{ Label = 'over 2 years';  Min = 730;  Max = [int]::MaxValue }
)
$MinDays = 90

$GraceDays = 30
#endregion

#region Data -----------------------------------------------------------------
function Get-StaleBucket {
    param([int]$Days)
    foreach ($b in $Buckets) {
        if ($Days -ge $b.Min -and $Days -lt $b.Max) { return $b.Label }
    }
    return ''
}

function Get-StaleAccountReport {
    param([scriptblock]$Progress, [int]$Threshold = 90)

    & $Progress 'Building licensed-user map from licence groups...' 5
    $lic = Get-LicensedUserMap -Progress { param($m) & $Progress $m 10 }
    if ($lic.Map.Count -eq 0) {
        throw "No licensed users found. Check the licence group names in Common.ps1 -- a name mismatch returns nothing rather than erroring."
    }

    $rows   = New-Object System.Collections.Generic.List[object]
    $failed = @()
    $total  = $lic.Map.Count
    $n = 0

    foreach ($dn in $lic.Map.Keys) {
        $n++
        if ($n % 50 -eq 0 -or $n -eq $total) {
            & $Progress "Reading accounts: $n of $total -- $($rows.Count) stale so far..." (15 + [int](80.0 * $n / $total))
        }
        $entry = $lic.Map[$dn]
        try {
            $u = Get-ADUser -Identity $dn -Server $entry.Server `
                    -Properties Enabled, mail, DisplayName, LastLogonDate, whenCreated, `
                                PasswordLastSet, Title, Department, Manager, Description `
                    -ErrorAction Stop
        } catch { $failed += $dn; continue }

        if (-not $u.Enabled) { continue }

        $days = if ($u.LastLogonDate) {
                    [int]((Get-Date) - $u.LastLogonDate).TotalDays
                } else { $null }
        $age  = if ($u.whenCreated) { [int]((Get-Date) - $u.whenCreated).TotalDays } else { $null }
        $never = (-not $u.LastLogonDate)

        $isStale = if ($never) { ($null -eq $age -or $age -gt $GraceDays) }
                   else        { $days -ge $Threshold }
        if (-not $isStale) { continue }

        $pwdDays = if ($u.PasswordLastSet) { [int]((Get-Date) - $u.PasswordLastSet).TotalDays } else { $null }

        $rows.Add([PSCustomObject][ordered]@{
            sAMAccountName    = "$($u.SamAccountName)"
            DisplayName       = "$($u.DisplayName)"
            Mail              = "$($u.mail)"
            Environment       = $entry.Environment
            Sku               = $entry.Skus
            LicenseGroups     = $entry.Groups
            LastLogonDate     = if ($u.LastLogonDate) { $u.LastLogonDate.ToString('yyyy-MM-dd') } else { '' }
            DaysSinceLogon    = $days
            NeverLoggedOn     = $never
            StaleBucket       = if ($never) { 'never logged on' } else { Get-StaleBucket -Days $days }
            WhenCreated       = if ($u.whenCreated) { $u.whenCreated.ToString('yyyy-MM-dd') } else { '' }
            AccountAgeDays    = $age
            PasswordLastSet   = if ($u.PasswordLastSet) { $u.PasswordLastSet.ToString('yyyy-MM-dd') } else { '' }
            DaysSincePwdSet   = $pwdDays
            Title             = "$($u.Title)"
            Department        = "$($u.Department)"
            Description       = "$($u.Description)"
        })
    }

    $sorted = @($rows | Sort-Object `
        @{ Expression = { -not $_.NeverLoggedOn } },
        @{ Expression = { if ($null -eq $_.DaysSinceLogon) { [int]::MaxValue } else { $_.DaysSinceLogon } }; Descending = $true })

    return [PSCustomObject]@{
        Rows          = $sorted
        NotFound      = $lic.NotFound
        Renamed       = $lic.Renamed
        Failed        = $failed
        LicensedTotal = $total
        Threshold     = $Threshold
    }
}

function Format-ReportSummary {
    param($Result)
    $rows = @($Result.Rows)
    $sb = New-Object System.Text.StringBuilder

    $never = @($rows | Where-Object { $_.NeverLoggedOn })

    [void]$sb.AppendLine("Licensed accounts examined     : $($Result.LicensedTotal)")
    [void]$sb.AppendLine("Enabled, licensed, and stale   : $($rows.Count)")
    [void]$sb.AppendLine("  Never logged on at all       : $($never.Count)")
    if ($Result.LicensedTotal -gt 0) {
        [void]$sb.AppendLine(("  Share of licensed estate     : {0}%" -f [math]::Round(100 * $rows.Count / $Result.LicensedTotal)))
    }
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine('By dormancy:')
    foreach ($b in (@('never logged on') + ($Buckets | ForEach-Object { $_.Label }))) {
        $c = @($rows | Where-Object { $_.StaleBucket -eq $b }).Count
        if ($c -gt 0) { [void]$sb.AppendLine(("  {0,-18}{1,6}" -f $b, $c)) }
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('By SKU:')
    $rows | Group-Object Sku | Sort-Object Count -Descending | ForEach-Object {
        [void]$sb.AppendLine(("  {0,-18}{1,6}" -f $_.Name, $_.Count))
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('By environment:')
    $rows | Group-Object Environment | Sort-Object Name | ForEach-Object {
        [void]$sb.AppendLine(("  {0,-18}{1,6}" -f $_.Name, $_.Count))
    }

    $old = @($rows | Where-Object { $_.StaleBucket -eq 'over 2 years' -or $_.NeverLoggedOn })
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("Strongest candidates (never used, or dormant 2+ years): $($old.Count)")

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('--- How to read LastLogonDate ---')
    [void]$sb.AppendLine('This comes from AD lastLogonTimestamp, which by design replicates only')
    [void]$sb.AppendLine('  every ~14 days. Treat every figure as accurate to within a fortnight.')
    [void]$sb.AppendLine('  That is fine at 90+ days and useless for "did they log on this week".')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Mailbox statistics were deliberately NOT used. LastLogonTime from')
    [void]$sb.AppendLine('  Get-MailboxStatistics is updated by background and service processes in')
    [void]$sb.AppendLine('  Exchange Online, so mailboxes nobody has opened in years can appear')
    [void]$sb.AppendLine('  active. It would make this report look precise and be wrong.')
    [void]$sb.AppendLine('  Entra sign-in logs via Graph are the accurate source, when available.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("Accounts created within $GraceDays days that have never signed in are")
    [void]$sb.AppendLine('  excluded, on the assumption they are starters who have not begun.')

    if ($Result.Renamed -and $Result.Renamed.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('NOTE -- matched only after normalising lookalike characters:')
        foreach ($r in $Result.Renamed) { [void]$sb.AppendLine("  $r") }
        [void]$sb.AppendLine('  The AD name contains a dash or space that is not plain ASCII.')
    }
    if ($Result.NotFound.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('*** LICENCE GROUPS NOT FOUND -- their members were not examined: ***')
        foreach ($g in $Result.NotFound) { [void]$sb.AppendLine("  $g") }
    }
    if ($Result.Failed.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("*** INCOMPLETE -- $($Result.Failed.Count) account(s) could not be read. ***")
    }
    return $sb.ToString()
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Report -- Stale Licensed Accounts'
$form.Size            = New-Object System.Drawing.Size(760, 676)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Accounts that are ENABLED and licensed but have not signed in. Disabled accounts are covered by the Disabled-but-Licensed report.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblInfo)

$lblSrc           = New-Object System.Windows.Forms.Label
$lblSrc.Text      = 'Source: on-prem AD lastLogonTimestamp (accurate to ~14 days). Mailbox statistics are not used -- see the note in the results.'
$lblSrc.Location  = New-Object System.Drawing.Point(12, 32)
$lblSrc.Size      = New-Object System.Drawing.Size(720, 18)
$lblSrc.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblSrc)

$lblDays          = New-Object System.Windows.Forms.Label
$lblDays.Text     = 'Stale after (days):'
$lblDays.Location = New-Object System.Drawing.Point(12, 60)
$lblDays.Size     = New-Object System.Drawing.Size(110, 20)
$form.Controls.Add($lblDays)

$numDays = New-Object System.Windows.Forms.NumericUpDown
$numDays.Location = New-Object System.Drawing.Point(126, 58)
$numDays.Size     = New-Object System.Drawing.Size(70, 22)
$numDays.Minimum  = 30
$numDays.Maximum  = 3650
$numDays.Value    = $MinDays
$form.Controls.Add($numDays)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run Report'
$btnRun.Location = New-Object System.Drawing.Point(12, 88)
$btnRun.Size     = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnRun)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Reports Folder'
$btnOpen.Location = New-Object System.Drawing.Point(140, 88)
$btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnOpen)


$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(622, 88)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 162)
$lblStatus.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 182 -Width 720 -Height 12

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Location   = New-Object System.Drawing.Point(12, 200)
$txtOut.Size       = New-Object System.Drawing.Size(720, 414)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'StaleAccounts' `
                -X 12 -Y 124 -Width 720 -StatusLabel $lblStatus


$btnOpen.Add_Click({ Open-ReportFolder -Path $ReportDir })

$btnRun.Add_Click({
    $btnRun.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = ''
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar `
                        -Activity 'Stale licensed accounts'
        $report = Get-StaleAccountReport -Progress $progress -Threshold ([int]$numDays.Value)

        Complete-ProgressReporter -Activity 'Stale licensed accounts' -Bar $pbar

        if (@($report.Rows).Count -eq 0) {
            $lblStatus.Text = 'No stale licensed accounts found -- nothing written.'
            $txtOut.Text = (Format-ReportSummary -Result $report) +
                           [Environment]::NewLine +
                           'No CSV was written because there were no rows.'
            return
        }

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir "StaleAccounts_$stamp.csv"
        $report.Rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $files.SelectPath($csv)
        $lblStatus.Text  = "Done. $(@($report.Rows).Count) row(s) written."

        $txtOut.Text = (Format-ReportSummary -Result $report) +
                       [Environment]::NewLine +
                       'Saved to:' + [Environment]::NewLine + "  $csv"

        $note = if ($report.NotFound.Count -or $report.Failed.Count) { ' -- PARTIAL' } else { '' }
        # PARTIAL is the keyword the log is scanned for; Result is the same
        # fact in a column a filter can see.
        if ($note) { $logResult = 'Partial' } else { $logResult = 'Success' }
        Write-CompletedLog -Task "Stale licensed accounts -- $(@($report.Rows).Count) rows$note" -Result $logResult
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Stale licensed accounts'
    } finally {
        Complete-ProgressReporter -Activity 'Stale licensed accounts'
        $form.Cursor    = $prev
        $btnRun.Enabled = $true
    }
})

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled, and BringToFront'd inside the helper because WinForms
# draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'MB w No activity' | Out-Null

[void]$form.ShowDialog()
#endregion
