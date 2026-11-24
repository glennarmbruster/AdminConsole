#Requires -Version 5.1
<#
    MissingRemoteMailbox.ps1
    Finds cloud user mailboxes with no on-premises remote mailbox stub, and stubs whose ExchangeGuid disagrees with the cloud. Author : Glenn Armbruster

    Author   : Glenn Armbruster
    Revised  : 2026-09-13

    2026-09-13  Split into two buttons -- Missing Stub and GUID Mismatch --
                each writing its own CSV with its own list of previous runs.
                The two faults are different sizes and different jobs, and one
                combined file buried the hundred-row work queue in four
                thousand rows of mismatch nobody is chasing. Adds an opt-in
                reuse of the last scan, and stamps the scan time into the
                filename, the summary and every row.
#>
# TOOLKIT-HELP-RUNS: reads every user mailbox in the cloud and every remote mailbox on-premises. Changes nothing.
# TOOLKIT-HELP-BEGIN
# Finds people whose mailbox exists in the cloud but who have NO on-premises
# remote mailbox object.
#
# It happens when a starter is licensed without anyone telling us. The licence
# provisions a mailbox in Exchange Online and everything looks fine to the
# user -- mail arrives, Outlook works -- while the on-premises stub that the
# rest of our process depends on was never created.
#
# WHAT BREAKS LATER, NOT NOW:
#
#   * scan-to-email, which routes through on-premises Exchange
#   * the disable process, which works from the on-premises object
#
# Neither fails at the time. They fail on the day somebody scans a document or
# the day that person leaves, which is the worst day to find out.
#
# IT ALSO FINDS A SECOND FAULT, because the same comparison finds it: a stub
# that EXISTS but whose ExchangeGuid does not match the cloud mailbox. Mail
# still flows, so nothing looks wrong, but the mailbox cannot be moved and a
# migration will fail on it later.
#
# TWO BUTTONS, TWO REPORTS, ON PURPOSE. In a long-lived hybrid the two counts
# are nothing like each other -- thousands of GUID mismatches that are causing
# no problem today, against a hundred missing stubs that are a work queue. One
# combined CSV buries the queue in the noise, so each button reads both
# directories, writes its OWN file, and has its own list of previous runs. Open
# either one without running the other.
#
# USER MAILBOXES ONLY. Shared, room, equipment, scheduling and Microsoft 365
# group mailboxes are out of scope -- they are provisioned differently and a
# missing stub does not mean the same thing for them.
#
# Writes a CSV. Read-only -- it tells you who to fix, it does not create
# anything. Needs both Exchange Online and on-premises Exchange.
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
if (-not (Test-ConsoleLaunch)) { return }

#region Configuration --------------------------------------------------------
$ReportDir = Get-ReportPath

# THE ONLY RECIPIENT TYPE IN SCOPE.
#
# Asked for explicitly rather than filtered afterwards, which also sidesteps
# C12: a Microsoft 365 group reports as 'GroupMailbox', a word that contains
# 'Mailbox' and would survive any substring test.
$InScopeType = 'UserMailbox'

# THE TWO ISSUE LABELS, WRITTEN ONCE.
#
# Each is the value in the CSV's Issue column AND the key that splits one scan
# into two reports. A typo in either place would not error -- it would silently
# write an empty CSV, which reads as "nothing to fix" and is the worst kind of
# wrong answer this tool can give. Named once, used everywhere.
$IssueMissing  = 'MISSING STUB'
$IssueMismatch = 'GUID MISMATCH'

# Report file prefixes. Separate prefixes are what make the two "previous runs"
# lists independent -- Get-ReportFileList matches on the prefix, so a shared one
# would show every file in both dropdowns.
$PrefixMissing  = 'MissingStub'
$PrefixMismatch = 'GuidMismatch'
#endregion

#region Data -----------------------------------------------------------------
function Get-CloudUserMailboxes {
    <#
        Every user mailbox in Exchange Online.

        ROUTED THROUGH Invoke-EXOCommand BECAUSE Get-Mailbox EXISTS IN BOTH
        EDITIONS (C1). This tool holds an on-premises session at the same time,
        so an unqualified call is a coin toss that reports on-premises data
        under a cloud heading -- and the whole point here is telling the two
        apart.

        ResultSize is explicit (C3). The default is 1000 and the short answer
        arrives without a warning, which on a report like this reads as "only a
        thousand people have mailboxes".
    #>
    param([scriptblock]$Progress)

    if ($Progress) { & $Progress 'Reading cloud user mailboxes...' }
    $rows = @(Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
        RecipientTypeDetails = $InScopeType
        ResultSize           = 'Unlimited'
        ErrorAction          = 'Stop'
    })
    if ($Progress) { & $Progress "Cloud user mailboxes: $($rows.Count)" }
    return $rows
}

function Get-OnPremRemoteMailboxIndex {
    <#
        Every on-premises remote mailbox, indexed for O(1) lookup.

        TWO INDEXES, ON PURPOSE, AND THE PAIR IS WHAT MAKES THE REPORT USEFUL:

            ByGuid  ExchangeGuid -> stub.  The authoritative link between an
                    on-premises stub and its cloud mailbox.
            ByAddr  every smtp address -> stub.  A weaker link, and the one
                    that catches an object which IS this person but whose GUID
                    has drifted.

        A cloud mailbox found in neither is a MISSING STUB. Found in ByAddr but
        not ByGuid is a GUID MISMATCH. Doing it with one index would collapse
        those two into a single wrong answer.

        NO NESTED LOOPS. A tenant with several thousand mailboxes compared
        against several thousand stubs is millions of comparisons done the
        naive way, and this report is meant to be run casually.
    #>
    param([scriptblock]$Progress)

    if ($Progress) { & $Progress 'Reading on-premises remote mailboxes...' }
    $stubs = @(Invoke-OnPremCommand -Name 'Get-RemoteMailbox' -Parameters @{
        ResultSize  = 'Unlimited'
        ErrorAction = 'Stop'
    })

    $byGuid = @{}
    $byAddr = @{}
    foreach ($s in $stubs) {
        $g = "$($s.ExchangeGuid)"
        # An empty GUID is not an identity. Indexing it would make every
        # unprovisioned stub match every unprovisioned mailbox.
        if ($g -and $g -ne [Guid]::Empty.ToString()) { $byGuid[$g.ToLower()] = $s }

        $p = "$($s.PrimarySmtpAddress)".Trim()
        if ($p) { $byAddr[$p.ToLower()] = $s }

        # EmailAddresses carries typed proxy addresses -- 'SMTP:' for the
        # primary, 'smtp:' for secondaries, 'X500:' and others that are not
        # addresses anyone mails. Take only the smtp ones, and never write the
        # collection back (C11) -- this function only reads it.
        foreach ($a in @($s.EmailAddresses)) {
            $t = "$a"
            if ($t -match '^(?i)smtp:(.+)$') {
                $addr = $Matches[1].Trim().ToLower()
                if ($addr -and -not $byAddr.ContainsKey($addr)) { $byAddr[$addr] = $s }
            }
        }
    }

    if ($Progress) { & $Progress "On-premises remote mailboxes: $($stubs.Count)" }
    return [PSCustomObject]@{
        Count  = $stubs.Count
        ByGuid = $byGuid
        ByAddr = $byAddr
    }
}

function Get-MissingStubReport {
    <#
        Compare the two sides and classify every cloud user mailbox.

        Returns the findings only. A mailbox with a healthy stub produces no
        row -- a report whose output is mostly "this one is fine" is a report
        nobody reads to the bottom.
    #>
    param([scriptblock]$Progress)

    # STAMPED WHEN THE READ STARTS, NOT WHEN THE FILE IS WRITTEN. One scan can
    # produce two CSVs minutes apart, and the useful question about either of
    # them is "how old is this data", not "when did somebody press Save".
    $scanAt = Get-Date

    $cloud = Get-CloudUserMailboxes -Progress $Progress
    $index = Get-OnPremRemoteMailboxIndex -Progress $Progress

    if ($Progress) { & $Progress 'Comparing...' }

    $rows    = New-Object System.Collections.Generic.List[object]
    $okCount = 0
    $i       = 0
    $total   = @($cloud).Count

    foreach ($mbx in $cloud) {
        $i++
        if ($Progress -and ($i % 250 -eq 0)) {
            & $Progress "Comparing $i of $total..." ([int](100 * $i / [Math]::Max($total, 1)))
        }

        $guid = "$($mbx.ExchangeGuid)".ToLower()
        $smtp = "$($mbx.PrimarySmtpAddress)".Trim().ToLower()

        $stub = $null
        if ($guid -and $index.ByGuid.ContainsKey($guid)) { $stub = $index.ByGuid[$guid] }

        if ($stub) { $okCount++; continue }

        # No GUID match. Is there an object that IS this person?
        $byAddr = $null
        if ($smtp -and $index.ByAddr.ContainsKey($smtp)) { $byAddr = $index.ByAddr[$smtp] }

        if ($byAddr) {
            $issue     = $IssueMismatch
            $onPremGud = "$($byAddr.ExchangeGuid)"
            $onPremDn  = "$($byAddr.DistinguishedName)"
        } else {
            $issue     = $IssueMissing
            $onPremGud = ''
            $onPremDn  = ''
        }

        # WhenCreated IS THE FIELD THAT DECIDES URGENCY. A mailbox created this
        # morning may simply be waiting for the next directory sync. One created
        # in March is a person who has been scanning to a mailbox that cannot
        # receive it for five months.
        $created = $null
        try { $created = [datetime]$mbx.WhenCreated } catch { $created = $null }
        $ageDays = if ($created) { [int]((Get-Date) - $created).TotalDays } else { $null }

        $rows.Add([PSCustomObject]@{
            Issue              = $issue
            DisplayName        = "$($mbx.DisplayName)"
            PrimarySmtpAddress = "$($mbx.PrimarySmtpAddress)"
            UserPrincipalName  = "$($mbx.UserPrincipalName)"
            CloudExchangeGuid  = "$($mbx.ExchangeGuid)"
            OnPremExchangeGuid = $onPremGud
            OnPremObject       = $onPremDn
            WhenCreated        = $created
            AgeDays            = $ageDays
        })
    }

    # OLDEST FIRST. The report is a work queue, and the person who has been
    # broken longest is the one to fix first.
    $sorted = @($rows.ToArray() | Sort-Object @{ Expression = 'Issue' }, @{ Expression = 'WhenCreated' })

    return [PSCustomObject]@{
        Rows         = $sorted
        ScanAt       = $scanAt
        CloudCount   = $total
        OnPremCount  = $index.Count
        HealthyCount = $okCount
        MissingCount = @($sorted | Where-Object { $_.Issue -eq $IssueMissing }).Count
        MismatchCount= @($sorted | Where-Object { $_.Issue -eq $IssueMismatch }).Count
    }
}

function Format-IssueSummary {
    <#
        The on-screen summary for ONE issue class.

        Deliberately NOT a summary of both. The whole reason this tool has two
        buttons is that the two faults are different sizes and different jobs:
        thousands of mismatches nobody is chasing, against a hundred missing
        stubs that are a queue. Printing both counts under either button puts
        the noise back on the screen the split was meant to clear.

        The OTHER count still appears, on one line, because "3,914 mismatches
        not shown" is the difference between a short report and a report you
        suspect of having missed something.
    #>
    param($Result, [string]$Issue, [string]$Title, [string]$Meaning, [string]$OtherLabel, [int]$OtherCount)

    $mine = @($Result.Rows | Where-Object { $_.Issue -eq $Issue })

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine($Title)
    [void]$sb.AppendLine(('-' * $Title.Length))
    [void]$sb.AppendLine("Scan taken                    : $($Result.ScanAt.ToString('yyyy-MM-dd HH:mm:ss'))")
    [void]$sb.AppendLine("Cloud user mailboxes examined : $($Result.CloudCount)")
    [void]$sb.AppendLine("On-premises remote mailboxes  : $($Result.OnPremCount)")
    [void]$sb.AppendLine("Correctly linked              : $($Result.HealthyCount)")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("FOUND: $($mine.Count)   $Meaning")
    [void]$sb.AppendLine("(the same scan found $OtherCount $OtherLabel -- not in this report)")
    [void]$sb.AppendLine('')
    if ($mine.Count -eq 0) {
        [void]$sb.AppendLine('Nothing to fix.')
        return $sb.ToString()
    }
    [void]$sb.AppendLine('Oldest first -- age is how long the mailbox has existed unlinked:')
    [void]$sb.AppendLine('')
    foreach ($r in @($mine | Select-Object -First 40)) {
        $age = if ($null -ne $r.AgeDays) { "{0,5} days" -f $r.AgeDays } else { '    ? days' }
        [void]$sb.AppendLine(("  {0}  {1}" -f $age, $r.PrimarySmtpAddress))
    }
    if ($mine.Count -gt 40) {
        [void]$sb.AppendLine("  ... and $($mine.Count - 40) more -- see the CSV.")
    }
    return $sb.ToString()
}
#endregion

#region UI -------------------------------------------------------------------
$form             = New-Object System.Windows.Forms.Form
$form.Text        = 'Missing Remote Mailbox'
$form.ClientSize  = New-Object System.Drawing.Size(760, 660)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblTenant          = New-Object System.Windows.Forms.Label
$lblTenant.Location = New-Object System.Drawing.Point(12, 12)
$lblTenant.Size     = New-Object System.Drawing.Size(736, 18)
$form.Controls.Add($lblTenant)

$lblIntro          = New-Object System.Windows.Forms.Label
$lblIntro.Text     = 'Two faults, two reports. MISSING STUB: a cloud user mailbox with no on-premises remote mailbox -- the object scan-to-email and the disable process both need. GUID MISMATCH: a stub exists but its ExchangeGuid disagrees with the cloud, so the mailbox cannot be moved.'
$lblIntro.Location = New-Object System.Drawing.Point(12, 36)
$lblIntro.Size     = New-Object System.Drawing.Size(736, 34)
$form.Controls.Add($lblIntro)

# TWO RUN BUTTONS, NOT ONE RUN BUTTON AND A FILTER.
#
# Each reads both directories and writes its own file, so either report can be
# produced, opened and mailed on without the other existing. The missing-stub
# button is first and leftmost because that is the one with a queue behind it.
$btnMissing          = New-Object System.Windows.Forms.Button
$btnMissing.Text     = 'Missing Stub'
$btnMissing.Location = New-Object System.Drawing.Point(12, 78)
$btnMissing.Size     = New-Object System.Drawing.Size(130, 30)
$form.Controls.Add($btnMissing)

$btnMismatch          = New-Object System.Windows.Forms.Button
$btnMismatch.Text     = 'GUID Mismatch'
$btnMismatch.Location = New-Object System.Drawing.Point(150, 78)
$btnMismatch.Size     = New-Object System.Drawing.Size(130, 30)
$form.Controls.Add($btnMismatch)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Reports folder'
$btnOpen.Location = New-Object System.Drawing.Point(288, 78)
$btnOpen.Size     = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnOpen)

# REUSE, OFFERED RATHER THAN ASSUMED.
#
# Both buttons do identical work -- read every cloud mailbox, read every stub,
# compare -- and on a tenant of this size that is minutes, twice, for two views
# of the same answer. So the scan is kept in memory and this box lets the second
# button use it.
#
# It starts UNCHECKED and only enables once a scan exists, because a cached
# answer that silently outlives a fix is worse than waiting. Whichever way it
# is set, the scan's own timestamp goes in the filename, in the summary and in
# every CSV row -- so a reused scan is never mistaken for a fresh one.
$chkReuse          = New-Object System.Windows.Forms.CheckBox
$chkReuse.Text     = 'Reuse last scan'
$chkReuse.Location = New-Object System.Drawing.Point(420, 84)
$chkReuse.Size     = New-Object System.Drawing.Size(210, 20)
$chkReuse.Enabled  = $false
$form.Controls.Add($chkReuse)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(638, 78)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 190)
$lblStatus.Size     = New-Object System.Drawing.Size(736, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 210 -Width 736 -Height 12

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Location   = New-Object System.Drawing.Point(12, 228)
$txtOut.Size       = New-Object System.Drawing.Size(736, 376)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
# Word wrap on (B6). A diagnosis off the right edge is a diagnosis nobody reads.
$txtOut.WordWrap   = $true
$form.Controls.Add($txtOut)

# TWO "previous runs" ROWS, ONE PER REPORT, EACH WITH ITS OWN PREFIX.
#
# Get-ReportFileList matches on the prefix, so these two lists never show each
# other's files -- which is the point. Open last week's missing-stub list
# without scrolling past forty mismatch files to find it.
#
# The same LabelWidth on both so the two dropdowns share a left edge.
$filesMissing = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix $PrefixMissing `
                -X 12 -Y 118 -Width 736 -StatusLabel $lblStatus `
                -LabelText 'Missing stub:' -LabelWidth 104

$filesMismatch = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix $PrefixMismatch `
                -X 12 -Y 152 -Width 736 -StatusLabel $lblStatus `
                -LabelText 'GUID mismatch:' -LabelWidth 104

$btnClose.Add_Click({ $form.Close() })
$btnOpen.Add_Click({ Open-ReportFolder -Path $ReportDir })
#endregion

#region Handlers -------------------------------------------------------------
# The last scan, kept so the second button can reuse it. $script: and not a
# plain local because the handlers below run from the message loop, long after
# the assignment statement finished -- see the scope note in Add-ReportFileControls.
$script:LastScan = $null

function Update-ReuseState {
    <#
        Enable the reuse box only when there is something to reuse, and say how
        old it is. The AGE is the whole point: "Reuse last scan" is a promise
        about data, and a promise about data has to carry a timestamp or the
        operator is guessing.
    #>
    if (-not $script:LastScan) {
        $chkReuse.Enabled = $false
        $chkReuse.Checked = $false
        $chkReuse.Text    = 'Reuse last scan'
        return
    }
    $mins = [int]((Get-Date) - $script:LastScan.ScanAt).TotalMinutes
    $age  = if ($mins -lt 1) { 'just now' } elseif ($mins -eq 1) { '1 min old' } else { "$mins mins old" }
    $chkReuse.Enabled = $true
    $chkReuse.Text    = "Reuse last scan ($($script:LastScan.ScanAt.ToString('HH:mm')), $age)"
}

function Test-BothConnections {
    <#
        BOTH SIDES ARE REQUIRED, AND THE MESSAGE SAYS WHICH ONE IS MISSING.
        "Not connected" on a tool that needs two connections is a guess.
    #>
    $exo = Test-ExoConnectionState
    if (-not $exo.Connected) {
        Show-ToolDialog(
            "Not connected to Exchange Online. Use the Connect button on the console first.`r`n`r`n$($exo.Detail)",
            'Exchange Online required', 'OK', 'Warning') | Out-Null
        return $false
    }
    if (-not (Test-EmsAvailable)) {
        Show-ToolDialog(
            "No on-premises Exchange session in this process. Connect OnPrem from the console first.`r`n`r`nWithout it every cloud mailbox would look like it was missing a stub.",
            'On-premises Exchange required', 'OK', 'Warning') | Out-Null
        return $false
    }
    return $true
}

function Invoke-IssueReport {
    <#
        Run (or reuse) one scan, then write the ONE issue class asked for.

        Both buttons come through here. They differ only in which Issue value
        they keep, which prefix they write under, and which "previous runs" row
        they refresh -- so the scan, the guard rails and the logging exist once.
        Two near-identical handlers would drift, and the way they would drift is
        one of them quietly losing the empty-file rule below.
    #>
    param(
        [Parameter(Mandatory)][string]$Issue,
        [Parameter(Mandatory)][string]$Prefix,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Meaning,
        [Parameter(Mandatory)][string]$OtherLabel,
        [Parameter(Mandatory)]$Files
    )

    if (-not (Test-BothConnections)) { return }

    $btnMissing.Enabled  = $false
    $btnMismatch.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = ''
    try {
        $reused = $false
        if ($chkReuse.Checked -and $script:LastScan) {
            $report = $script:LastScan
            $reused = $true
            $lblStatus.Text = "Reusing the scan taken at $($report.ScanAt.ToString('HH:mm:ss'))."
        } else {
            $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity $Title
            $report   = Get-MissingStubReport -Progress $progress
            Complete-ProgressReporter -Activity $Title -Bar $pbar
            $script:LastScan = $report
        }
        Update-ReuseState

        $other = if ($Issue -eq $IssueMissing) { $report.MismatchCount } else { $report.MissingCount }
        $txtOut.Text = Format-IssueSummary -Result $report -Issue $Issue -Title $Title `
                            -Meaning $Meaning -OtherLabel $OtherLabel -OtherCount $other
        if ($reused) {
            $txtOut.Text += [Environment]::NewLine +
                "NOTE: this used the scan already in memory, not a fresh read." + [Environment]::NewLine
        }

        $rows = @($report.Rows | Where-Object { $_.Issue -eq $Issue })

        if ($rows.Count -eq 0) {
            # NEVER A ZERO-ROW FILE. An empty CSV and a failed run look
            # identical in a folder listing a week later.
            $lblStatus.Text = "$Title -- none found. Nothing written."
            $txtOut.Text += [Environment]::NewLine + 'No CSV was written because there were no rows.'
            Write-CompletedLog -Action 'Report' -Target $Title -Result 'Success' `
                -Detail "Examined $($report.CloudCount) cloud user mailbox(es); no $Issue findings."
            return
        }

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }

        # NAMED FROM THE SCAN, NOT FROM NOW. Two reports off one scan then carry
        # the SAME stamp, which is true and useful -- they are two views of one
        # moment. Stamping them with the save time would make a reused scan look
        # like fresh data, which is the one thing the reuse box must not do.
        $stamp = $report.ScanAt.ToString('yyyy-MM-dd_HHmm')
        $csv   = Join-Path $ReportDir "${Prefix}_$stamp.csv"

        # ScanTakenAt on every row, so the answer survives the file being
        # renamed, copied into a ticket, or pasted into a spreadsheet.
        $rows |
            Select-Object *, @{ Name = 'ScanTakenAt'; Expression = { $report.ScanAt } } |
            Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $Files.SelectPath($csv)
        $lblStatus.Text = "$Title -- $($rows.Count) found. Saved."
        $txtOut.Text += [Environment]::NewLine + 'Saved to:' + [Environment]::NewLine + "  $csv"

        Write-CompletedLog -Action 'Report' -Target $csv -Result 'Success' `
            -Detail ("Examined $($report.CloudCount) cloud user mailbox(es): " +
                     "$($rows.Count) with $Issue." +
                     $(if ($reused) { ' Reused an in-memory scan.' } else { '' }))
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context $Title
        Write-FailedLog -Action 'Report' -Target $Title -ErrorRecord $_
    } finally {
        Complete-ProgressReporter -Activity $Title
        $form.Cursor         = $prev
        $btnMissing.Enabled  = $true
        $btnMismatch.Enabled = $true
    }
}

# Every handler wrapped (B1). An exception inside a WinForms handler is caught
# by .NET rather than PowerShell, so without this the operator gets a CLR crash
# dialog with no script name and no line number. Here the wrapping lives in
# Invoke-IssueReport, which both handlers are.
$btnMissing.Add_Click({
    Invoke-IssueReport -Issue $IssueMissing -Prefix $PrefixMissing `
        -Title      'Missing on-premises stub' `
        -Meaning    'cloud user mailboxes with NO on-premises remote mailbox at all' `
        -OtherLabel 'GUID mismatches' `
        -Files      $filesMissing
})

$btnMismatch.Add_Click({
    Invoke-IssueReport -Issue $IssueMismatch -Prefix $PrefixMismatch `
        -Title      'GUID mismatch' `
        -Meaning    'stubs that exist but whose ExchangeGuid disagrees with the cloud' `
        -OtherLabel 'missing stubs' `
        -Files      $filesMismatch
})

# The age in the checkbox caption goes stale while the window sits open. It is
# refreshed on Activated rather than on a timer: the number only matters at the
# moment somebody comes back to the window to press the other button.
$form.Add_Activated({ try { Update-ReuseState } catch { } })
#endregion

#region Show -----------------------------------------------------------------
$d = Get-ConnectedTenantDomain
if ($d) {
    $lblTenant.Text      = "Tenant: $d"
    $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(0, 128, 0)
} else {
    $lblTenant.Text      = 'Tenant: not connected -- connect to Exchange Online and OnPrem first.'
    $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(200, 0, 0)
}

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
# Consistent '?' in the top-right of every tool. Added after the theme so it is
# not restyled, and BringToFront'd inside the helper because WinForms draws
# EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Missing Remote Mailbox' | Out-Null

[void]$form.ShowDialog()
#endregion
