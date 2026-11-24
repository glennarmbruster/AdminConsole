#Requires -Version 5.1
<#
    FranUser_Office.ps1
    Every user in the franchise directory, with their primary SMTP address and
    office.

    Author   : Glenn Armbruster
    Revised  : 2026-08-11
#>
# TOOLKIT-HELP-BEGIN
# Lists every user in the franchise directory with their primary SMTP address
# and the Office field from Active Directory, and writes a CSV.
#
# Read-only. It reads Active Directory and writes a file; it changes nothing in
# the directory.
#
# ACCOUNTS BEGINNING "test" ARE LEFT OUT, matched on either the logon name or
# the display name, because they are not real offices and they distort a count
# that gets read as one. The number skipped is reported, so the rule is visible
# rather than silent.
#
# Users with no Office set are still listed, with the field blank. That is the
# more useful answer: a missing office is usually the thing worth finding.
#
# The domain controller and search base come from the Franchise directory in
# Tenant Manager -- not from anything typed into this script. The DC box is
# editable if you need a specific one for a single run.
#
# The CSV goes to the reports folder, alongside every other report, and the
# previous runs are one dropdown away.
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

#region Configuration -----------------------------------------------------------

# WHICH DIRECTORY. The only thing named here; everything else is read from it.
#
# The server and search base used to be literals -- a DC hostname copied from
# the script this replaced, and a distinguished name spelling out the domain.
# Both were tenant-specific values living in code, which the Tester flags as
# TEN001/TEN002, and it was right to: the hostname had drifted from the one
# Tenant.json actually configures.
$script:ReportDirectoryId = 'Franchise'

function Get-ReportScope {
    <#
        Server and search base for the directory, from Tenant.json.

        The search base is DERIVED from the directory's domain rather than
        stored separately. Two spellings of the same fact drift apart, and a
        distinguished name is a mechanical transform of a domain name --
        'franchises.contoso.com' is 'DC=franchises,DC=contoso,DC=com', always.
    #>
    $d = Get-TenantDirectory -Id $script:ReportDirectoryId
    if (-not $d) { return $null }

    $domain = "$($d.Domain)".Trim()
    if (-not $domain) { return $null }

    return [PSCustomObject]@{
        Domain     = $domain
        Server     = (Get-TenantDirectoryServer -Id $script:ReportDirectoryId)
        SearchBase = ('DC=' + (($domain -split '\.') -join ',DC='))
    }
}

$script:Scope = Get-ReportScope
if (-not $script:Scope) {
    $null = Show-ToolDialog(
        ("This report reads the '$($script:ReportDirectoryId)' directory, and Tenant.json " +
         "does not have one with a domain set." + "`r`n`r`n" +
         "Add it in Tenant Manager under Directories."),
        'Directory not configured', 'OK', 'Warning')
    return
}

# Matched against SamAccountName AND Name, case-insensitively. -like rather than
# StartsWith so the pattern stays readable and stays in one place.
$script:SkipPattern = 'test*'

#endregion

#region Query -------------------------------------------------------------------

function Get-FranchiseUserOffice {
    <#
        Returns the rows plus the counts behind them.

        The skipped count is returned rather than discarded: "1,482 users" and
        "1,482 users, 37 test accounts left out" are different statements, and
        only the second one can be checked.
    #>
    param([string]$Server, [scriptblock]$Progress)

    if ($Progress) { & $Progress 'Reading the franchise directory...' 10 }

    $all = @(Get-ADUser -Filter * -Server $Server -SearchBase $script:Scope.SearchBase `
                        -Properties mail, physicalDeliveryOfficeName -ErrorAction Stop)

    if ($Progress) { & $Progress "Filtering $($all.Count) account(s)..." 60 }

    # Filtered here rather than in -Filter. An LDAP filter cannot express "not
    # starting with" without getting hard to read, and this runs once against a
    # list already in memory. The cost is one pass; the benefit is a rule anyone
    # can check by eye.
    $skipped = @($all | Where-Object {
        "$($_.SamAccountName)" -like $script:SkipPattern -or "$($_.Name)" -like $script:SkipPattern
    })
    $keep = @($all | Where-Object {
        -not ("$($_.SamAccountName)" -like $script:SkipPattern -or "$($_.Name)" -like $script:SkipPattern)
    })

    if ($Progress) { & $Progress 'Building the report...' 85 }

    $rows = @($keep | Sort-Object Name | Select-Object `
        SamAccountName,
        Name,
        @{ Name = 'PrimarySMTPAddress'; Expression = { "$($_.mail)" } },
        @{ Name = 'Office';             Expression = { "$($_.physicalDeliveryOfficeName)" } })

    $noOffice = @($rows | Where-Object { -not "$($_.Office)".Trim() }).Count

    if ($Progress) { & $Progress 'Done.' 100 }

    return [PSCustomObject]@{
        Rows        = $rows
        TotalFound  = $all.Count
        Skipped     = $skipped.Count
        NoOffice    = $noOffice
        SkippedList = @($skipped | ForEach-Object { "$($_.SamAccountName)" } | Sort-Object)
    }
}

#endregion

#region Form --------------------------------------------------------------------

$ReportDir = Get-ReportPath

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Franchise Users -- office report'
$form.Size = New-Object System.Drawing.Size(780, 560)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblTop = New-Object System.Windows.Forms.Label
$lblTop.Text = 'Every user in the franchise directory with their primary SMTP address and Office. Accounts beginning "test" are left out.'
$lblTop.Location = New-Object System.Drawing.Point(12, 12)
$lblTop.Size = New-Object System.Drawing.Size(740, 32)
$lblTop.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblTop)

$lblServer = New-Object System.Windows.Forms.Label
$lblServer.Text = 'Domain controller:'
$lblServer.Location = New-Object System.Drawing.Point(12, 54)
$lblServer.Size = New-Object System.Drawing.Size(120, 18)
$form.Controls.Add($lblServer)

$txtServer = New-Object System.Windows.Forms.TextBox
$txtServer.Location = New-Object System.Drawing.Point(136, 51)
$txtServer.Size = New-Object System.Drawing.Size(400, 22)
$txtServer.Text = $(if ("$DomainController".Trim()) { "$DomainController".Trim() } else { $script:Scope.Server })
$form.Controls.Add($txtServer)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = 'Run'
$btnRun.Location = New-Object System.Drawing.Point(548, 49)
$btnRun.Size = New-Object System.Drawing.Size(100, 26)
$form.Controls.Add($btnRun)

$btnOpen = New-Object System.Windows.Forms.Button
$btnOpen.Text = 'Reports folder'
$btnOpen.Location = New-Object System.Drawing.Point(654, 49)
$btnOpen.Size = New-Object System.Drawing.Size(98, 26)
$form.Controls.Add($btnOpen)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Idle.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 84)
$lblStatus.Size = New-Object System.Drawing.Size(740, 18)
$form.Controls.Add($lblStatus)

$pbar = New-Object System.Windows.Forms.ProgressBar
$pbar.Location = New-Object System.Drawing.Point(12, 104)
$pbar.Size = New-Object System.Drawing.Size(740, 12)
$form.Controls.Add($pbar)

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'FranUser_Office' `
             -X 12 -Y 126 -Width 740 -StatusLabel $lblStatus

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(12, 166)
$txtOut.Size = New-Object System.Drawing.Size(740, 340)
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$btnOpen.Add_Click({ Open-ReportFolder -Path $ReportDir })

$btnRun.Add_Click({
    $server = "$($txtServer.Text)".Trim()
    if (-not $server) { $lblStatus.Text = 'Enter a domain controller first.'; return }

    $btnRun.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = ''
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar
        $report = Get-FranchiseUserOffice -Server $server -Progress $progress

        if (@($report.Rows).Count -eq 0) {
            $lblStatus.Text = 'No users returned.'
            $txtOut.Text = "The query returned nothing.$([Environment]::NewLine)" +
                           "Server     : $server$([Environment]::NewLine)" +
                           "Search base: $($script:Scope.SearchBase)"
            Write-CompletedLog -Action 'Report' -Target 'FranchiseUserOffice' -Result 'Partial' `
                -Detail "No users returned from $server"
            return
        }

        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir "FranUser_Office_$stamp.csv"
        $report.Rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $files.SelectPath($csv)
        $lblStatus.Text = "Done. $(@($report.Rows).Count) row(s) written."

        $nl = [Environment]::NewLine
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("Server       : $server")
        [void]$sb.AppendLine("Search base  : $($script:Scope.SearchBase)")
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("Accounts read: $($report.TotalFound)")
        [void]$sb.AppendLine("Exported     : $(@($report.Rows).Count)")
        [void]$sb.AppendLine("Skipped      : $($report.Skipped)  (name or logon begins 'test')")
        [void]$sb.AppendLine("No Office set: $($report.NoOffice)")
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Saved to:')
        [void]$sb.AppendLine("  $csv")

        # The skipped accounts are named, not just counted. A rule that removes
        # rows from a report someone will act on should be auditable by the
        # person reading it.
        if ($report.Skipped -gt 0) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('Skipped accounts:')
            foreach ($s in $report.SkippedList) { [void]$sb.AppendLine("  $s") }
        }
        $txtOut.Text = $sb.ToString()

        Write-CompletedLog -Action 'Report' -Target 'FranchiseUserOffice' -Result 'Success' `
            -Detail ("$(@($report.Rows).Count) row(s), $($report.Skipped) skipped, " +
                     "$($report.NoOffice) with no office; $csv")
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Franchise user office report'
        Write-FailedLog -Action 'Report' -Target 'FranchiseUserOffice' -ErrorRecord $_ `
            -Detail "server: $server"
    } finally {
        $btnRun.Enabled = $true
        $form.Cursor = $prev
    }
})

Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Franchise Users' | Out-Null
Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'

[void]$form.ShowDialog()
$form.Dispose()
#endregion
