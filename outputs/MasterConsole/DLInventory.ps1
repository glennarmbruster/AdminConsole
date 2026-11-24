#Requires -Version 5.1
<#
    DLInventory.ps1
    Inventories mail-enabled groups from on-prem AD for migration planning, and exports to CSV. Author : Glenn Armbruster

    Reasoning: ScriptNotes/DLInventory.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-07-31
#>
# TOOLKIT-HELP-BEGIN
# Inventories the mail-enabled groups in on-premises Active Directory, for
# migration planning. Writes a CSV.
#
# Built for the distribution list move to the cloud: what exists, how big it is,
# and who manages it, in one file you can sort and plan against.
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

#region Configuration --------------------------------------------------------
$ReportDir = Get-ReportPath

$StaleDays = 365
#endregion

#region Data -----------------------------------------------------------------
function Get-ForestTargets {
    # Read from AD, not Exchange: Get-DistributionGroup exists in BOTH shells
    # and resolves by module load order. This report is about the on-prem side.
    $t = @()
    if ($script:HQDomain) {
        $t += [PSCustomObject]@{ Environment = 'HQ'; Server = $script:HQDomain }
    }
    if ($script:FranchiseAutoDetectDC) {
        $t += [PSCustomObject]@{ Environment = 'Franchise'; Server = $script:FranchiseAutoDetectDC }
    }
    return $t
}

function Get-DLInventory {
    param([scriptblock]$Progress)

    $rows       = New-Object System.Collections.Generic.List[object]
    $failed     = @()
    $ownerCache = @{}
    $targets = Get-ForestTargets
    if ($targets.Count -eq 0) {
        throw "No AD targets configured. Check `$script:HQDomain and `$script:FranchiseAutoDetectDC in Common.ps1."
    }

    $ti = 0
    foreach ($t in $targets) {
        $ti++
        $base = 100 * ($ti - 1) / $targets.Count
        & $Progress "Querying $($t.Environment) for mail-enabled groups..." ([int]$base)

        try {
            $groups = @(Get-ADGroup -Filter 'mail -like "*"' -Server $t.Server `
                            -Properties mail, ManagedBy, GroupCategory, GroupScope, `
                                        member, whenCreated, whenChanged, `
                                        Description, msExchHideFromAddressLists, `
                                        DisplayName, info `
                            -ErrorAction Stop)
        } catch {
            $failed += "$($t.Environment) ($($_.Exception.Message))"
            & $Progress "$($t.Environment) failed -- continuing." ([int]$base)
            continue
        }

        & $Progress "$($t.Environment): indexing group DNs for nesting check..." ([int]$base)
        $groupDNs = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        try {
            foreach ($ag in (Get-ADGroup -Filter * -Server $t.Server -ErrorAction Stop)) {
                [void]$groupDNs.Add("$($ag.DistinguishedName)")
            }
        } catch {
            $failed += "$($t.Environment) nesting index ($($_.Exception.Message))"
        }

        $n = 0
        foreach ($g in $groups) {
            $n++
            if ($n % 25 -eq 0 -or $n -eq $groups.Count) {
                $pct = [int]($base + (100.0 / $targets.Count) * ($n / [Math]::Max(1, $groups.Count)))
                & $Progress "$($t.Environment): $n of $($groups.Count) groups..." $pct
            }

            $members = @($g.member)
            $direct  = $members.Count

            $nested = 0
            if ($direct -gt 0 -and $groupDNs.Count -gt 0) {
                foreach ($dn in $members) {
                    if ($groupDNs.Contains("$dn")) { $nested++ }
                }
            }

            $owner = ''
            if ($g.ManagedBy) {
                $odn = "$($g.ManagedBy)"
                if ($ownerCache.ContainsKey($odn)) {
                    $owner = $ownerCache[$odn]
                } else {
                    try {
                        $o = Get-ADObject -Identity $odn -Server $t.Server `
                                -Properties DisplayName, mail -ErrorAction Stop
                        $owner = if ($o.mail) { "$($o.mail)" } else { "$($o.Name)" }
                    } catch { $owner = $odn }
                    $ownerCache[$odn] = $owner
                }
            }

            $changedDays = if ($g.whenChanged) { [int]((Get-Date) - $g.whenChanged).TotalDays } else { $null }

            $rows.Add([PSCustomObject][ordered]@{
                Name              = "$($g.Name)"
                DisplayName       = "$($g.DisplayName)"
                PrimarySmtpAddress = "$($g.mail)"
                Environment       = $t.Environment
                GroupCategory     = "$($g.GroupCategory)"
                GroupScope        = "$($g.GroupScope)"
                MailEnabledSecurity = ("$($g.GroupCategory)" -eq 'Security')
                DirectMembers     = $direct
                NestedGroups      = $nested
                IsEmpty           = ($direct -eq 0)
                IsSingleMember    = ($direct -eq 1)
                HasOwner          = [bool]$g.ManagedBy
                ManagedBy         = $owner
                HiddenFromGAL     = [bool]$g.msExchHideFromAddressLists
                WhenCreated       = if ($g.whenCreated) { $g.whenCreated.ToString('yyyy-MM-dd') } else { '' }
                WhenChanged       = if ($g.whenChanged) { $g.whenChanged.ToString('yyyy-MM-dd') } else { '' }
                DaysSinceChange   = $changedDays
                LooksStale        = ($null -ne $changedDays -and $changedDays -gt $StaleDays)
                Description       = "$($g.Description)"
                Notes             = "$($g.info)"
                DistinguishedName = "$($g.DistinguishedName)"
            })
        }
    }

    $sorted = @($rows | Sort-Object `
        @{ Expression = { $_.NestedGroups }; Descending = $true },
        @{ Expression = { $_.HasOwner } },
        @{ Expression = { $_.DirectMembers }; Descending = $true })

    return [PSCustomObject]@{ Rows = $sorted; Failed = $failed }
}

function Format-ReportSummary {
    param($Result)
    $rows = @($Result.Rows)
    $sb = New-Object System.Text.StringBuilder

    [void]$sb.AppendLine("Mail-enabled groups found : $($rows.Count)")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('By environment:')
    $rows | Group-Object Environment | Sort-Object Name | ForEach-Object {
        [void]$sb.AppendLine(("  {0,-14}{1,6}" -f $_.Name, $_.Count))
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('By type:')
    $rows | Group-Object GroupCategory | Sort-Object Name | ForEach-Object {
        [void]$sb.AppendLine(("  {0,-14}{1,6}" -f $_.Name, $_.Count))
    }

    $empty   = @($rows | Where-Object { $_.IsEmpty })
    $single  = @($rows | Where-Object { $_.IsSingleMember })
    $noOwner = @($rows | Where-Object { -not $_.HasOwner })
    $nested  = @($rows | Where-Object { $_.NestedGroups -gt 0 })
    $stale   = @($rows | Where-Object { $_.LooksStale })
    $hidden  = @($rows | Where-Object { $_.HiddenFromGAL })
    $mesg    = @($rows | Where-Object { $_.MailEnabledSecurity })

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('--- Migration planning ---')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("Retire rather than migrate:")
    [void]$sb.AppendLine(("  Empty (no members)          {0,6}" -f $empty.Count))
    [void]$sb.AppendLine(("  Single member only          {0,6}" -f $single.Count))
    [void]$sb.AppendLine(("  Unchanged > $StaleDays days        {0,6}" -f $stale.Count))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Need a decision before migrating:')
    [void]$sb.AppendLine(("  No owner (ManagedBy empty)  {0,6}   <-- hardest: nobody to ask" -f $noOwner.Count))
    [void]$sb.AppendLine(("  Contain nested groups       {0,6}" -f $nested.Count))
    [void]$sb.AppendLine(("  Mail-enabled security       {0,6}   membership AND permissions" -f $mesg.Count))
    [void]$sb.AppendLine(("  Hidden from the GAL         {0,6}" -f $hidden.Count))

    $clean = @($rows | Where-Object {
        -not $_.IsEmpty -and -not $_.IsSingleMember -and $_.HasOwner -and
        $_.NestedGroups -eq 0 -and -not $_.MailEnabledSecurity })
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine(("Straightforward migrations   {0,6}" -f $clean.Count))

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Note on staleness: DaysSinceChange is when the OBJECT last changed,')
    [void]$sb.AppendLine('  not when the group was last EMAILED. AD holds no last-used data and')
    [void]$sb.AppendLine('  message trace only reaches back ~10 days, so a group can look stale')
    [void]$sb.AppendLine('  and still be in daily use. Confirm before retiring anything.')

    if ($Result.Failed -and $Result.Failed.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('*** INCOMPLETE -- these forests could not be read: ***')
        foreach ($f in $Result.Failed) { [void]$sb.AppendLine("  $f") }
    }
    return $sb.ToString()
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Report -- Distribution List Inventory'
$form.Size            = New-Object System.Drawing.Size(760, 676)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Inventories mail-enabled groups from on-prem AD, with migration-readiness indicators.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblInfo)

$lblSrc           = New-Object System.Windows.Forms.Label
$lblSrc.Text      = "Source: on-prem AD (HQ: $($script:HQDomain))  --  Exchange is not queried, so this cannot hit the on-prem/cloud cmdlet collision."
$lblSrc.Location  = New-Object System.Drawing.Point(12, 32)
$lblSrc.Size      = New-Object System.Drawing.Size(720, 18)
$lblSrc.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblSrc)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run Report'
$btnRun.Location = New-Object System.Drawing.Point(12, 60)
$btnRun.Size     = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnRun)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Reports Folder'
$btnOpen.Location = New-Object System.Drawing.Point(140, 60)
$btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnOpen)


$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(622, 60)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 136)
$lblStatus.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 156 -Width 720 -Height 12

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Location   = New-Object System.Drawing.Point(12, 174)
$txtOut.Size       = New-Object System.Drawing.Size(720, 440)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'DLInventory' `
                -X 12 -Y 96 -Width 720 -StatusLabel $lblStatus


$btnOpen.Add_Click({ Open-ReportFolder -Path $ReportDir })

$btnRun.Add_Click({
    $btnRun.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = ''
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar `
                        -Activity 'DL inventory'
        $report = Get-DLInventory -Progress $progress

        Complete-ProgressReporter -Activity 'DL inventory' -Bar $pbar

        if (@($report.Rows).Count -eq 0) {
            $lblStatus.Text = 'No mail-enabled groups found -- nothing written.'
            $txtOut.Text = (Format-ReportSummary -Result $report) +
                           [Environment]::NewLine +
                           'No CSV was written because there were no rows.'
            return
        }

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir "DLInventory_$stamp.csv"
        $report.Rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $files.SelectPath($csv)
        $lblStatus.Text  = "Done. $(@($report.Rows).Count) group(s) written."

        $txtOut.Text = (Format-ReportSummary -Result $report) +
                       [Environment]::NewLine +
                       'Saved to:' + [Environment]::NewLine + "  $csv"

        $note = if ($report.Failed.Count) { ' -- PARTIAL' } else { '' }
        # PARTIAL is the keyword the log is scanned for; Result is the same
        # fact in a column a filter can see.
        if ($note) { $logResult = 'Partial' } else { $logResult = 'Success' }
        Write-CompletedLog -Task "DL inventory -- $(@($report.Rows).Count) groups$note" -Result $logResult
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'DL inventory'
    } finally {
        Complete-ProgressReporter -Activity 'DL inventory'
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
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'DL Inventory' | Out-Null

[void]$form.ShowDialog()
#endregion
