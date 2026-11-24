#Requires -Version 5.1
<#
    OrphanedOneDrive.ps1
    Finds OneDrive sites whose owner is disabled or gone from AD, and exports to CSV. Author : Glenn Armbruster

    Reasoning: ScriptNotes/OrphanedOneDrive.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-07-31
#>
# TOOLKIT-HELP-BEGIN
# Finds OneDrive sites whose owner is disabled or no longer in Active Directory.
#
# When someone leaves, their OneDrive does not leave with them. It sits there
# holding whatever they had, still consuming storage, with nobody responsible
# for it.
#
# Writes a CSV. Read-only -- deciding what to do with the content is a
# conversation, not a script.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Configuration --------------------------------------------------------
$ReportDir = Get-ReportPath

$ReasonDisabled = 'Owner disabled in AD'
$ReasonMissing  = 'Owner not found in AD'
$ReasonNoOwner  = 'No owner on the site'
#endregion

#region Data -----------------------------------------------------------------
function Get-OrphanedOneDriveReport {
    param([scriptblock]$Progress, [switch]$WithStorage)

    if (-not (Get-Command -Name 'Get-SPOSite' -ErrorAction SilentlyContinue)) {
        throw 'Get-SPOSite is not available. Import Microsoft.Online.SharePoint.PowerShell and run Connect-SPOService first.'
    }

    & $Progress 'Enumerating OneDrive sites...' 3

    $declared = Get-CmdletParameterNames -Name 'Get-SPOSite'
    $params = @{ ErrorAction = 'Stop' }
    if ($declared -contains 'Limit')               { $params['Limit'] = 'ALL' }
    if ($declared -contains 'IncludePersonalSite') { $params['IncludePersonalSite'] = $true }
    if ($declared -contains 'Filter')              { $params['Filter'] = "Url -like '/personal/'" }

    $cappedRisk = ($declared -notcontains 'Limit')

    $sites = @(Get-SPOSite @params)

    $sites = @($sites | Where-Object { "$($_.Url)" -match '/personal/' })

    $total = @($sites).Count
    if ($total -eq 0) {
        throw 'No OneDrive (personal) sites were returned. Confirm the account has SharePoint administrator rights.'
    }

    $rows    = New-Object System.Collections.Generic.List[object]
    $adCache = @{}
    $detailFailed = @()
    $n = 0

    foreach ($s in $sites) {
        $n++
        if ($n % 25 -eq 0 -or $n -eq $total) {
            & $Progress "Matching owners against AD: $n of $total -- $($rows.Count) orphan(s) so far..." (5 + [int](75.0 * $n / $total))
        }

        $owner = "$($s.Owner)"

        if (-not $owner) {
            $rows.Add((New-OrphanRow -Site $s -Reason $ReasonNoOwner -Ad $null))
            continue
        }

        $ad = Resolve-ADAccountBySmtp -Smtp $owner -Cache $adCache

        if ($ad.Error) { continue }

        if (-not $ad.Found) {
            $rows.Add((New-OrphanRow -Site $s -Reason $ReasonMissing -Ad $ad))
        } elseif (-not $ad.Enabled) {
            $rows.Add((New-OrphanRow -Site $s -Reason $ReasonDisabled -Ad $ad))
        }
    }

    if ($WithStorage -and $rows.Count -gt 0) {
        $m = 0
        foreach ($r in $rows) {
            $m++
            & $Progress "Reading storage for finding $m of $($rows.Count)..." (80 + [int](18.0 * $m / $rows.Count))
            try {
                $d = Get-SPOSite -Identity $r.Url -ErrorAction Stop
                $r.StorageUsedMB = [math]::Round(([double]$d.StorageUsageCurrent), 1)
                $r.LastModified  = if ($d.LastContentModifiedDate) { $d.LastContentModifiedDate.ToString('yyyy-MM-dd') } else { '' }
            } catch {
                $detailFailed += "$($r.Url) : $($_.Exception.Message)"
            }
        }
    }

    $sorted = @($rows | Sort-Object `
        @{ Expression = { if ($null -eq $_.StorageUsedMB) { -1 } else { $_.StorageUsedMB } }; Descending = $true },
        Owner)

    $unknown = @($adCache.Values | Where-Object { $_.Error })

    return [PSCustomObject]@{
        Rows          = $sorted
        SiteTotal     = $total
        OwnersChecked = $adCache.Count
        Unresolvable  = $unknown.Count
        DetailFailed  = $detailFailed
        CappedRisk    = $cappedRisk
        StorageRead   = [bool]$WithStorage
    }
}

function New-OrphanRow {
    param($Site, [string]$Reason, $Ad)
    return [PSCustomObject][ordered]@{
        Owner          = "$($Site.Owner)"
        Reason         = $Reason
        DisplayName    = if ($Ad) { $Ad.DisplayName } else { '' }
        sAMAccountName = if ($Ad) { $Ad.SamAccountName } else { '' }
        Environment    = if ($Ad) { $Ad.Environment } else { '' }
        StorageUsedMB  = $null
        LastModified   = $null
        LockState      = "$($Site.LockState)"
        Url            = "$($Site.Url)"
    }
}

function Format-ReportSummary {
    param($Result)
    $rows = @($Result.Rows)
    $sb = New-Object System.Text.StringBuilder

    [void]$sb.AppendLine("OneDrive sites examined         : $($Result.SiteTotal)")
    [void]$sb.AppendLine("Orphaned OneDrive sites        : $($rows.Count)")
    if ($Result.SiteTotal -gt 0) {
        [void]$sb.AppendLine(("  Share of all OneDrives       : {0}%" -f [math]::Round(100 * $rows.Count / $Result.SiteTotal)))
    }
    [void]$sb.AppendLine("Distinct owners looked up      : $($Result.OwnersChecked)")
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine('By reason:')
    $rows | Group-Object Reason | Sort-Object Count -Descending | ForEach-Object {
        [void]$sb.AppendLine(("  {0,-26}{1,6}" -f $_.Name, $_.Count))
    }

    if ($rows.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('By environment:')
        $rows | Group-Object Environment | Sort-Object Name | ForEach-Object {
            $name = if ($_.Name) { $_.Name } else { '(unknown)' }
            [void]$sb.AppendLine(("  {0,-26}{1,6}" -f $name, $_.Count))
        }
    }

    if ($Result.StorageRead) {
        $sized = @($rows | Where-Object { $null -ne $_.StorageUsedMB })
        $totalMB = ($sized | Measure-Object StorageUsedMB -Sum).Sum
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine(("Data held in orphaned OneDrives : {0:N1} GB across {1} site(s)" -f
            (([double]$totalMB) / 1024), $sized.Count))
        $big = @($sized | Where-Object { $_.StorageUsedMB -ge 1024 })
        [void]$sb.AppendLine("  Sites holding 1 GB or more    : $($big.Count)")
        if ($big.Count -gt 0) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('Largest:')
            foreach ($b in ($sized | Select-Object -First 10)) {
                [void]$sb.AppendLine(("  {0,9:N1} MB   {1,-22} {2}" -f $b.StorageUsedMB, $b.Reason, $b.Owner))
            }
        }
    } else {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Storage figures were not collected -- tick "Include storage used" to')
        [void]$sb.AppendLine('  fetch them. They need one extra call per finding, because')
        [void]$sb.AppendLine('  StorageUsageCurrent is not returned by a bulk site enumeration.')
    }

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('--- How owners are judged ---')
    [void]$sb.AppendLine('On-prem AD is the authority, matched on mail, userPrincipalName or')
    [void]$sb.AppendLine('  proxyAddresses. This is the same basis every other report here uses,')
    [void]$sb.AppendLine('  so the answers stay consistent between them.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('An owner the directory could not answer for is SKIPPED, not reported.')
    [void]$sb.AppendLine('  A timed-out DC would otherwise turn every user on it into an orphan.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Retention note: a disabled owner does not mean the OneDrive is due for')
    [void]$sb.AppendLine('  deletion. Check for a retention policy or legal hold before removing')
    [void]$sb.AppendLine('  anything -- this report finds data, it does not authorise disposal.')

    if ($Result.Unresolvable -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("NOTE -- $($Result.Unresolvable) owner(s) could not be looked up in any")
        [void]$sb.AppendLine('  directory and were skipped. Their sites are absent from this report.')
    }
    if ($Result.CappedRisk) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('*** Get-SPOSite on this module build does not declare -Limit, so the ***')
        [void]$sb.AppendLine('*** site list may be capped at 200. Treat as INCOMPLETE.             ***')
    }
    if ($Result.DetailFailed.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("*** PARTIAL -- storage could not be read for $($Result.DetailFailed.Count) site(s). ***")
        foreach ($f in ($Result.DetailFailed | Select-Object -First 5)) { [void]$sb.AppendLine("  $f") }
    }
    return $sb.ToString()
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Report -- Orphaned OneDrive Sites'
$form.Size            = New-Object System.Drawing.Size(760, 660)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'OneDrive sites whose owner is disabled in AD, missing from AD, or absent from the site itself.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblInfo)

$lblSrc           = New-Object System.Windows.Forms.Label
$lblSrc.Text      = 'Needs SharePoint Online connected (Connect-SPOService) and on-prem AD reachable.'
$lblSrc.Location  = New-Object System.Drawing.Point(12, 32)
$lblSrc.Size      = New-Object System.Drawing.Size(720, 18)
$lblSrc.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblSrc)

$chkStorage          = New-Object System.Windows.Forms.CheckBox
$chkStorage.Text     = 'Include storage used (one extra call per finding)'
$chkStorage.Location = New-Object System.Drawing.Point(12, 54)
$chkStorage.Size     = New-Object System.Drawing.Size(340, 20)
$chkStorage.Checked  = $true
$form.Controls.Add($chkStorage)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run Report'
$btnRun.Location = New-Object System.Drawing.Point(12, 80)
$btnRun.Size     = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnRun)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Reports Folder'
$btnOpen.Location = New-Object System.Drawing.Point(140, 80)
$btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnOpen)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(622, 80)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 154)
$lblStatus.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 174 -Width 720 -Height 12

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Location   = New-Object System.Drawing.Point(12, 192)
$txtOut.Size       = New-Object System.Drawing.Size(720, 406)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'OrphanedOneDrive' `
                -X 12 -Y 116 -Width 720 -StatusLabel $lblStatus

$btnOpen.Add_Click({ Open-ReportFolder -Path $ReportDir })

$btnRun.Add_Click({
    $btnRun.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = ''
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar `
                        -Activity 'Orphaned OneDrive sites'
        $report = Get-OrphanedOneDriveReport -Progress $progress -WithStorage:$chkStorage.Checked

        Complete-ProgressReporter -Activity 'Orphaned OneDrive sites' -Bar $pbar

        if (@($report.Rows).Count -eq 0) {
            $lblStatus.Text = 'Every OneDrive has a live owner -- nothing written.'
            $txtOut.Text = (Format-ReportSummary -Result $report) +
                           [Environment]::NewLine +
                           'No CSV was written because there were no rows.'
            return
        }

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir "OrphanedOneDrive_$stamp.csv"
        $report.Rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $files.SelectPath($csv)
        $lblStatus.Text = "Done. $(@($report.Rows).Count) row(s) written."

        $txtOut.Text = (Format-ReportSummary -Result $report) +
                       [Environment]::NewLine +
                       'Saved to:' + [Environment]::NewLine + "  $csv"

        $note = if ($report.DetailFailed.Count -or $report.CappedRisk -or $report.Unresolvable) { ' -- PARTIAL' } else { '' }
        # PARTIAL is the keyword the log is scanned for; Result is the same
        # fact in a column a filter can see.
        if ($note) { $logResult = 'Partial' } else { $logResult = 'Success' }
        Write-CompletedLog -Task "Orphaned OneDrive sites -- $(@($report.Rows).Count) rows$note" -Result $logResult
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Orphaned OneDrive sites'
    } finally {
        Complete-ProgressReporter -Activity 'Orphaned OneDrive sites'
        $form.Cursor    = $prev
        $btnRun.Enabled = $true
    }
})

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Orphaned OneDrive' | Out-Null
[void]$form.ShowDialog()
#endregion
