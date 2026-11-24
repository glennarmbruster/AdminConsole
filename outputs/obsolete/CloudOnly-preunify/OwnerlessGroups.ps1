#Requires -Version 5.1
<#
    OwnerlessGroups.ps1
    Finds Microsoft 365 groups with no effective owner, and exports to CSV. Author : Glenn Armbruster Revised : 2026-07-29

    Reasoning: ScriptNotes/OwnerlessGroups.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-07-31
#>
# TOOLKIT-HELP-BEGIN
# Finds Microsoft 365 groups with no effective owner.
#
# An ownerless group cannot be managed by anyone but an admin -- nobody can add
# members, change settings or delete it -- and it happens quietly whenever the
# last owner leaves. Nobody is notified.
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

$ReasonZero     = 'No owner listed'
$ReasonDisabled = 'All owners disabled in AD'
$ReasonMissing  = 'All owners missing from AD'
$ReasonMixed    = 'All owners disabled or missing'
#endregion

#region Data -----------------------------------------------------------------
function Get-OwnerlessGroupReport {
    param([scriptblock]$Progress)

    if (-not (Get-Command -Name 'Get-UnifiedGroup' -ErrorAction SilentlyContinue)) {
        throw 'Get-UnifiedGroup is not available. Connect to Exchange Online first.'
    }

    & $Progress 'Reading Microsoft 365 groups...' 3

    # ResultSize defaults to 1000; omitting it silently truncates the tenant.
    $params = @{ ErrorAction = 'Stop' }
    $declared = Get-CmdletParameterNames -Name 'Get-UnifiedGroup'
    if ($declared -contains 'ResultSize') { $params['ResultSize'] = 'Unlimited' }
    $truncationRisk = ($declared -notcontains 'ResultSize')

    $groups = @(Get-UnifiedGroup @params)

    $total = @($groups).Count
    if ($total -eq 0) { throw 'No Microsoft 365 groups were returned.' }

    $rows     = New-Object System.Collections.Generic.List[object]
    $adCache  = @{}
    $failed   = @()
    $linkless = @()
    $n = 0

    $canLinks = [bool](Get-Command -Name 'Get-UnifiedGroupLinks' -ErrorAction SilentlyContinue)

    foreach ($g in $groups) {
        $n++
        if ($n % 10 -eq 0 -or $n -eq $total) {
            & $Progress "Checking owners: $n of $total -- $($rows.Count) finding(s) so far..." (5 + [int](90.0 * $n / $total))
        }

        $managed = @($g.ManagedBy | Where-Object { $_ })

        $ownerAddrs = @()
        if ($managed.Count -gt 0 -and $canLinks) {
            try {
                $ownerAddrs = @(Get-UnifiedGroupLinks -Identity $g.Identity -LinkType Owners -ErrorAction Stop |
                                ForEach-Object { "$($_.PrimarySmtpAddress)" } |
                                Where-Object { $_ })
            } catch {
                $failed += "$($g.DisplayName) : $($_.Exception.Message)"
                continue
            }
            if ($ownerAddrs.Count -eq 0) { $linkless += "$($g.DisplayName)" }
        }

        $ownerStates = @()
        foreach ($addr in $ownerAddrs) {
            $ownerStates += Resolve-ADAccountBySmtp -Smtp $addr -Cache $adCache
        }

        $unknown = @($ownerStates | Where-Object { $_.Error })
        $live    = @($ownerStates | Where-Object { $_.Found -and $_.Enabled })
        $off     = @($ownerStates | Where-Object { $_.Found -and -not $_.Enabled })
        $gone    = @($ownerStates | Where-Object { -not $_.Found -and -not $_.Error })

        $reason = $null
        if ($managed.Count -eq 0) {
            $reason = $ReasonZero
        } elseif ($live.Count -eq 0 -and $unknown.Count -eq 0 -and $ownerStates.Count -gt 0) {
            $reason = if ($off.Count -gt 0 -and $gone.Count -gt 0) { $ReasonMixed }
                      elseif ($gone.Count -gt 0)                   { $ReasonMissing }
                      else                                          { $ReasonDisabled }
        }
        if (-not $reason) { continue }

        $members = if ($null -ne $g.GroupMemberCount) { [int]$g.GroupMemberCount } else { $null }

        $isTeam = ($g.ResourceProvisioningOptions -contains 'Team')

        $rows.Add([PSCustomObject][ordered]@{
            DisplayName        = "$($g.DisplayName)"
            PrimarySmtpAddress = "$($g.PrimarySmtpAddress)"
            Reason             = $reason
            TeamsEnabled       = $isTeam
            Privacy            = "$($g.AccessType)"
            Members            = $members
            OwnersListed       = $managed.Count
            OwnersDisabled     = $off.Count
            OwnersMissing      = $gone.Count
            OwnerAddresses     = ($ownerAddrs -join '; ')
            DisabledOwners     = (($off | ForEach-Object { $_.Smtp }) -join '; ')
            WhenCreated        = if ($g.WhenCreated) { $g.WhenCreated.ToString('yyyy-MM-dd') } else { '' }
            SharePointSiteUrl  = "$($g.SharePointSiteUrl)"
            Notes              = "$($g.Notes)"
        })
    }

    $sorted = @($rows | Sort-Object `
        @{ Expression = { -not $_.TeamsEnabled } },
        @{ Expression = { if ($null -eq $_.Members) { -1 } else { $_.Members } }; Descending = $true })

    return [PSCustomObject]@{
        Rows           = $sorted
        GroupTotal     = $total
        Failed         = $failed
        Linkless       = $linkless
        TruncationRisk = $truncationRisk
        LinksAvailable = $canLinks
        OwnersResolved = $adCache.Count
    }
}

function Format-ReportSummary {
    param($Result)
    $rows = @($Result.Rows)
    $sb = New-Object System.Text.StringBuilder

    [void]$sb.AppendLine("Microsoft 365 groups examined  : $($Result.GroupTotal)")
    [void]$sb.AppendLine("Groups with no effective owner : $($rows.Count)")
    if ($Result.GroupTotal -gt 0) {
        [void]$sb.AppendLine(("  Share of all groups          : {0}%" -f [math]::Round(100 * $rows.Count / $Result.GroupTotal)))
    }
    [void]$sb.AppendLine("Distinct owners resolved in AD : $($Result.OwnersResolved)")
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine('By reason:')
    $rows | Group-Object Reason | Sort-Object Count -Descending | ForEach-Object {
        [void]$sb.AppendLine(("  {0,-32}{1,6}" -f $_.Name, $_.Count))
    }
    [void]$sb.AppendLine('')

    $teams = @($rows | Where-Object { $_.TeamsEnabled })
    [void]$sb.AppendLine("Teams-enabled among the findings: $($teams.Count)")
    [void]$sb.AppendLine("Plain groups (no Team)          : $($rows.Count - $teams.Count)")

    $big = @($rows | Where-Object { $null -ne $_.Members -and $_.Members -ge 10 })
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("Unmanaged groups with 10+ members: $($big.Count)")
    if ($big.Count -gt 0) {
        foreach ($b in ($big | Select-Object -First 10)) {
            [void]$sb.AppendLine(("  {0,-40} {1,4} members   {2}" -f
                ($b.DisplayName -replace '\s+', ' '), $b.Members, $b.Reason))
        }
        if ($big.Count -gt 10) { [void]$sb.AppendLine("  ... and $($big.Count - 10) more in the CSV.") }
    }

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('--- What this report does and does not check ---')
    [void]$sb.AppendLine('"All owners disabled" is the finding that matters here. A group can list')
    [void]$sb.AppendLine('  an owner and still be unmanaged, because the owner left and the disable')
    [void]$sb.AppendLine('  script turned the account off without touching group ownership. Any')
    [void]$sb.AppendLine('  ownerless query that only looks for an empty owner list misses these.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Owner accounts are judged by on-premises AD Enabled state, matched on mail,')
    [void]$sb.AppendLine('  userPrincipalName or proxyAddresses. An owner the directory could not')
    [void]$sb.AppendLine('  answer for is treated as ALIVE, not dead -- so a directory problem')
    [void]$sb.AppendLine('  cannot invent findings.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Soft-deleted groups are excluded. They are already being removed.')

    if (-not $Result.LinksAvailable) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('*** Get-UnifiedGroupLinks was not available, so owner accounts could ***')
        [void]$sb.AppendLine('*** not be checked. Only "No owner listed" rows are reliable.       ***')
    }
    if ($Result.TruncationRisk) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('*** Get-UnifiedGroup on this module build does not declare ResultSize. ***')
        [void]$sb.AppendLine('*** Results may be capped at 1000 groups -- treat as INCOMPLETE.       ***')
    }
    if ($Result.Linkless.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("NOTE -- $($Result.Linkless.Count) group(s) list an owner but returned no owner")
        [void]$sb.AppendLine('  links. Reported as "No owner listed" only if ManagedBy was also empty.')
    }
    if ($Result.Failed.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("*** INCOMPLETE -- $($Result.Failed.Count) group(s) could not be read: ***")
        foreach ($f in ($Result.Failed | Select-Object -First 10)) { [void]$sb.AppendLine("  $f") }
    }
    return $sb.ToString()
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Report -- Ownerless Microsoft 365 Groups'
$form.Size            = New-Object System.Drawing.Size(760, 660)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Microsoft 365 groups with no owner, or whose owners are all disabled in AD. Includes Teams-connected groups.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblInfo)

$lblSrc           = New-Object System.Windows.Forms.Label
$lblSrc.Text      = 'Source: Exchange Online groups plus on-prem AD for owner state. No Graph permission or admin consent needed.'
$lblSrc.Location  = New-Object System.Drawing.Point(12, 32)
$lblSrc.Size      = New-Object System.Drawing.Size(720, 18)
$lblSrc.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblSrc)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run Report'
$btnRun.Location = New-Object System.Drawing.Point(12, 58)
$btnRun.Size     = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnRun)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Reports Folder'
$btnOpen.Location = New-Object System.Drawing.Point(140, 58)
$btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnOpen)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(622, 58)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 132)
$lblStatus.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 152 -Width 720 -Height 12

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Location   = New-Object System.Drawing.Point(12, 170)
$txtOut.Size       = New-Object System.Drawing.Size(720, 428)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'OwnerlessGroups' `
                -X 12 -Y 94 -Width 720 -StatusLabel $lblStatus

$btnOpen.Add_Click({ Open-ReportFolder -Path $ReportDir })

$btnRun.Add_Click({
    $btnRun.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = ''
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar `
                        -Activity 'Ownerless groups'
        $report = Get-OwnerlessGroupReport -Progress $progress

        Complete-ProgressReporter -Activity 'Ownerless groups' -Bar $pbar

        if (@($report.Rows).Count -eq 0) {
            $lblStatus.Text = 'Every group has at least one enabled owner -- nothing written.'
            $txtOut.Text = (Format-ReportSummary -Result $report) +
                           [Environment]::NewLine +
                           'No CSV was written because there were no rows.'
            return
        }

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir "OwnerlessGroups_$stamp.csv"
        $report.Rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $files.SelectPath($csv)
        $lblStatus.Text = "Done. $(@($report.Rows).Count) row(s) written."

        $txtOut.Text = (Format-ReportSummary -Result $report) +
                       [Environment]::NewLine +
                       'Saved to:' + [Environment]::NewLine + "  $csv"

        $note = if ($report.Failed.Count -or $report.TruncationRisk -or -not $report.LinksAvailable) { ' -- PARTIAL' } else { '' }
        # PARTIAL is the keyword the log is scanned for; Result is the same
        # fact in a column a filter can see.
        if ($note) { $logResult = 'Partial' } else { $logResult = 'Success' }
        Write-CompletedLog -Task "Ownerless M365 groups -- $(@($report.Rows).Count) rows$note" -Result $logResult
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Ownerless groups'
    } finally {
        Complete-ProgressReporter -Activity 'Ownerless groups'
        $form.Cursor    = $prev
        $btnRun.Enabled = $true
    }
})

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Ownerless Groups' | Out-Null
[void]$form.ShowDialog()
#endregion
