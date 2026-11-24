#Requires -Version 5.1
<#
    DisabledLicensed.ps1
    Reports disabled AD accounts that are still in a licence group, and exports to CSV. Author : Glenn Armbruster

    Reasoning: ScriptNotes/DisabledLicensed.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-07-31
#>
# TOOLKIT-HELP-BEGIN
# Finds accounts that are disabled in Active Directory but still sitting in a
# licence group -- so they are still costing money.
#
# Straightforward cleanup list, and usually a longer one than expected. Writes a
# CSV.
#
# Read-only. Removing the licence is a separate, deliberate step.
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

$LicenseGroups = $script:LicenseGroups

$FutureStarterDays = 30
#endregion

#region Data -----------------------------------------------------------------
function Get-DisabledLicensedReport {
    param(
        [scriptblock]$Progress,
        [switch]$IncludeEnabled
    )

    $rows      = New-Object System.Collections.Generic.List[object]
    $notFound  = @()
    $renamed   = @()
    $failed    = @()
    $totals    = @{}
    $seen      = @{}

    $i = 0
    foreach ($lg in $LicenseGroups) {
        $i++
        $pct = [int](100 * ($i - 1) / $LicenseGroups.Count)
        & $Progress "Resolving group '$($lg.Name)'..." $pct

        $found = Resolve-LicenseGroupObject -Name $lg.Name
        if (-not $found) {
            $notFound += $lg.Name
            & $Progress "'$($lg.Name)' not found in either forest -- skipping." $pct
            continue
        }

        if ($found.MatchedOn -eq 'normalised') {
            $renamed += "$($lg.Name)  ->  actual AD name: $($found.ActualName)"
        }

        & $Progress "Reading members of '$($lg.Name)' ($($found.Forest))..." $pct
        try {
            $members = @(Get-ADGroupMember -Identity $found.Group.DistinguishedName `
                            -Server $found.Server -Recursive -ErrorAction Stop |
                         Where-Object { $_.objectClass -eq 'user' })
        } catch {
            $failed += "$($lg.Name) ($($_.Exception.Message))"
            & $Progress "Could not read '$($lg.Name)' -- continuing." $pct
            continue
        }

        $totals[$lg.Name] = $members.Count

        $n = 0
        foreach ($m in $members) {
            $n++
            if ($n % 50 -eq 0) {
                $inner = $pct + [int]((100.0 / $LicenseGroups.Count) * ($n / [Math]::Max(1, $members.Count)))
                & $Progress "$($lg.Name): $n of $($members.Count) -- $($rows.Count) flagged so far..." $inner
            }
            try {
                $u = Get-ADUser -Identity $m.DistinguishedName -Server $found.Server `
                        -Properties Enabled, mail, DisplayName, whenCreated, whenChanged, `
                                    LastLogonDate, Description, Title, Department `
                        -ErrorAction Stop
            } catch { continue }

            if ($u.Enabled -and -not $IncludeEnabled) { continue }

            $key = "$($u.DistinguishedName)".ToLower()
            if ($seen.ContainsKey($key)) {
                $existing = $seen[$key]
                $existing.LicenseGroup = "$($existing.LicenseGroup); $($lg.Name)"
                if ($existing.Sku -notlike "*$($lg.Sku)*") {
                    $existing.Sku = "$($existing.Sku); $($lg.Sku)"
                }
                continue
            }

            $ageDays  = if ($u.whenCreated) { [int]((Get-Date) - $u.whenCreated).TotalDays } else { $null }
            $lastDays = if ($u.LastLogonDate) { [int]((Get-Date) - $u.LastLogonDate).TotalDays } else { $null }

            $row = [PSCustomObject][ordered]@{
                sAMAccountName   = "$($u.SamAccountName)"
                DisplayName      = "$($u.DisplayName)"
                Mail             = "$($u.mail)"
                Enabled          = [bool]$u.Enabled
                Sku              = $lg.Sku
                LicenseGroup     = $lg.Name
                Environment      = $lg.Environment
                Forest           = $found.Forest
                WhenCreated      = if ($u.whenCreated)   { $u.whenCreated.ToString('yyyy-MM-dd') }   else { '' }
                AccountAgeDays   = $ageDays
                LastLogonDate    = if ($u.LastLogonDate) { $u.LastLogonDate.ToString('yyyy-MM-dd') } else { '' }
                DaysSinceLogon   = $lastDays
                NeverLoggedOn    = (-not $u.LastLogonDate)
                LikelyFutureHire = ($null -ne $ageDays -and $ageDays -le $FutureStarterDays -and -not $u.LastLogonDate)
                Title            = "$($u.Title)"
                Department       = "$($u.Department)"
                LastADChange     = if ($u.whenChanged)   { $u.whenChanged.ToString('yyyy-MM-dd') }   else { '' }
            }
            $rows.Add($row)
            $seen[$key] = $row
        }
    }

    $sorted = @($rows | Sort-Object `
        @{ Expression = { $_.LikelyFutureHire } },
        @{ Expression = { if ($null -eq $_.DaysSinceLogon) { [int]::MaxValue } else { $_.DaysSinceLogon } }; Descending = $true })

    return [PSCustomObject]@{
        Rows          = $sorted
        NotFound      = $notFound
        Renamed       = $renamed
        Failed        = $failed
        GroupTotals   = $totals
        IncludedEnabled = [bool]$IncludeEnabled
    }
}

function Format-ReportSummary {
    param($Result)
    $rows = @($Result.Rows)
    $sb = New-Object System.Text.StringBuilder

    $flagged = @($rows | Where-Object { -not $_.Enabled })
    $future  = @($rows | Where-Object { $_.LikelyFutureHire })
    $real    = @($flagged | Where-Object { -not $_.LikelyFutureHire })

    [void]$sb.AppendLine("Disabled accounts still in a licence group : $($flagged.Count)")
    [void]$sb.AppendLine("  Likely future hires (created <= $FutureStarterDays days, never logged on) : $($future.Count)")
    [void]$sb.AppendLine("  Genuine reclaim candidates                                    : $($real.Count)")
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine('Reclaim candidates by SKU:')
    $real | Group-Object Sku | Sort-Object Name | ForEach-Object {
        [void]$sb.AppendLine(("  {0,-8}{1,6}" -f $_.Name, $_.Count))
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Reclaim candidates by environment:')
    $real | Group-Object Environment | Sort-Object Name | ForEach-Object {
        [void]$sb.AppendLine(("  {0,-12}{1,6}" -f $_.Name, $_.Count))
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Group membership totals (all members, enabled and disabled):')
    foreach ($k in ($Result.GroupTotals.Keys | Sort-Object)) {
        [void]$sb.AppendLine(("  {0,-26}{1,6}" -f $k, $Result.GroupTotals[$k]))
    }

    $never = @($real | Where-Object { $_.NeverLoggedOn })
    if ($never.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("Disabled, licensed, and never logged on: $($never.Count)")
    }

    if ($Result.Renamed -and $Result.Renamed.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('NOTE -- matched only after normalising lookalike characters:')
        foreach ($r in $Result.Renamed) { [void]$sb.AppendLine("  $r") }
        [void]$sb.AppendLine('  The AD name contains a dash or space that is not the plain ASCII')
        [void]$sb.AppendLine('  character. Worth fixing in AD, or copying the real name into')
        [void]$sb.AppendLine('  $script:LicenseGroups in Common.ps1.')
    }
    if ($Result.NotFound.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('*** GROUPS NOT FOUND -- these were not checked at all: ***')
        foreach ($g in $Result.NotFound) { [void]$sb.AppendLine("  $g") }
        [void]$sb.AppendLine('  Check the names in the Configuration region of this script.')
    }
    if ($Result.Failed.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('*** INCOMPLETE -- these groups could not be read: ***')
        foreach ($g in $Result.Failed) { [void]$sb.AppendLine("  $g") }
    }
    return $sb.ToString()
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Report -- Disabled but Licensed'
$form.Size            = New-Object System.Drawing.Size(760, 676)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Finds disabled AD accounts still in a licence group. Membership of those groups is what assigns the licence.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblInfo)

$lblGroups           = New-Object System.Windows.Forms.Label
$lblGroups.Text      = 'Groups: ' + (($LicenseGroups | ForEach-Object { $_.Name }) -join '  |  ')
$lblGroups.Location  = New-Object System.Drawing.Point(12, 32)
$lblGroups.Size      = New-Object System.Drawing.Size(720, 18)
$lblGroups.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblGroups)

$chkAll          = New-Object System.Windows.Forms.CheckBox
$chkAll.Text     = 'Include enabled members too (full licence inventory)'
$chkAll.Location = New-Object System.Drawing.Point(12, 56)
$chkAll.Size     = New-Object System.Drawing.Size(340, 20)
$form.Controls.Add($chkAll)

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

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'DisabledLicensed' `
                -X 12 -Y 118 -Width 720 -StatusLabel $lblStatus


$btnOpen.Add_Click({ Open-ReportFolder -Path $ReportDir })

$btnRun.Add_Click({
    $btnRun.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = ''
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar `
                        -Activity 'Disabled but licensed report'
        $report = Get-DisabledLicensedReport -Progress $progress -IncludeEnabled:$chkAll.Checked

        if (@($report.Rows).Count -eq 0) {
            $lblStatus.Text = 'Nothing to report -- nothing written.'
            $txtOut.Text = (Format-ReportSummary -Result $report) +
                           [Environment]::NewLine +
                           'No CSV was written because there were no rows.'
            return
        }

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir "DisabledLicensed_$stamp.csv"
        $report.Rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        Complete-ProgressReporter -Activity 'Disabled but licensed report' -Bar $pbar
        $files.SelectPath($csv)
        $lblStatus.Text  = "Done. $(@($report.Rows).Count) row(s) written."

        $txtOut.Text = (Format-ReportSummary -Result $report) +
                       [Environment]::NewLine +
                       'Saved to:' + [Environment]::NewLine + "  $csv"

        $note = if ($report.NotFound.Count -or $report.Failed.Count) { ' -- PARTIAL' } else { '' }
        # PARTIAL is the keyword the log is scanned for; Result is the same
        # fact in a column a filter can see.
        if ($note) { $logResult = 'Partial' } else { $logResult = 'Success' }
        Write-CompletedLog -Task "Disabled-but-licensed report -- $(@($report.Rows).Count) rows$note" -Result $logResult
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Disabled but licensed report'
    } finally {
        Complete-ProgressReporter -Activity 'Disabled but licensed report'
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
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Disabled/Licensed' | Out-Null

[void]$form.ShowDialog()
#endregion
