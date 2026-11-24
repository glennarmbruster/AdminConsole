#Requires -Version 5.1
<#
    ProtocolUsage.ps1
    Cross-references which legacy protocols are ENABLED against who has actually USED them, and writes the evidence for a change request. Author : Glenn Armbruster

    Author   : Glenn Armbruster
    Revised  : 2026-08-24
#>
# TOOLKIT-HELP-RUNS: reads every mailbox's protocol settings and a usage export you provide. Changes nothing.
# TOOLKIT-HELP-BEGIN
# Protocols.ps1 says what is ENABLED. This says what is USED.
#
# Built to get a change request approved. "These are enabled on 4,000
# mailboxes" is a configuration statement and invites the question that stops
# the change: how do you know nobody needs them.
#
# HOW TO USE IT
#   1. Press "Open sign-in logs" and export the sign-in report from Entra,
#      filtered to the legacy client apps. Or export the Email App Usage
#      report from the M365 admin centre.
#   2. Point this at the CSV.
#   3. It reads the enabled state live, joins the two, and writes both a
#      per-mailbox CSV and a summary paragraph you can paste into the ticket.
#
# HOW FAR BACK YOU CAN ACTUALLY PROVE -- READ THIS BEFORE PROMISING A YEAR
#
#   Entra sign-in logs      30 days. Not negotiable without a Log Analytics
#                           workspace that was ALREADY collecting them.
#   Email App Usage report  180 days maximum.
#
#   Neither gives you a year. A year of evidence exists only if somebody set
#   up log shipping before you needed it.
#
# THE ARGUMENT THAT DOES NOT NEED A YEAR
#
#   Microsoft REMOVED Basic authentication for POP, IMAP, ActiveSync, EWS and
#   the rest in late 2022. Nobody has signed in to those with a password since.
#   What the enabled flag still permits is POP/IMAP over OAuth, which is rare
#   and which a 30-day sign-in export will show.
#
#   So the case is: the dangerous half has been dead for three years, the
#   surviving half has zero users in the window you can evidence, and the
#   setting is on for N mailboxes with no consumer.
#
# Read-only. Writes a CSV and nothing else.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

#region Configuration --------------------------------------------------------
$ReportDir = Get-ReportPath

# Entra: sign-in logs, which is the definitive protocol-level source.
$SignInUrl = 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/SignInEventsV3.ReactView'
# M365 admin centre: Email App Usage, which reaches 180 days.
$AppUsageUrl = 'https://admin.microsoft.com/#/reportsUsage/EmailAppsUsage'

# THE PROTOCOLS IN QUESTION, and the Get-CASMailbox property that reports each.
$script:Protocols = [ordered]@{
    'POP'          = 'PopEnabled'
    'IMAP'         = 'ImapEnabled'
    'ActiveSync'   = 'ActiveSyncEnabled'
    'EWS'          = 'EwsEnabled'
}
#endregion

#region Usage import ---------------------------------------------------------
function Get-UsageColumnMap {
    <#
        Work out which columns in an exported CSV mean what.

        THE HEADERS ARE NOT STABLE AND I CANNOT SEE YOUR EXPORT. Entra and the
        M365 admin centre both rename columns between releases, and the admin
        centre localises them. Matching on a SUBSTRING of the header, rather
        than on an exact name, survives that -- and when it does not, the
        window prints every header it saw so the mismatch is visible instead of
        producing a confident zero.

        A REPORT THAT CANNOT FIND ITS COLUMNS MUST SAY SO. Silently reading no
        usage from a file it did not understand is the worst possible outcome
        here: it would produce exactly the answer you were hoping for.
    #>
    param($Rows)

    if (@($Rows).Count -eq 0) { return $null }
    $headers = @($Rows[0].PSObject.Properties.Name)

    $map = @{ Headers = $headers; User = $null; Protocols = @{}; Kind = 'Unknown' }

    foreach ($h in $headers) {
        $l = "$h".ToLower()
        if (-not $map.User -and ($l -match 'user principal name|userprincipalname|^upn$|user id|username')) {
            $map.User = $h
        }
    }

    # SIGN-IN EXPORT: one row per sign-in, with the client app in a column.
    $appCol = @($headers | Where-Object { "$_".ToLower() -match 'client app|clientappused|application' })[0]
    if ($appCol) { $map.Kind = 'SignIn'; $map.AppColumn = $appCol }

    # APP USAGE EXPORT: one row per user, one column per protocol.
    foreach ($h in $headers) {
        $l = "$h".ToLower()
        if ($l -match '\bpop3?\b')                     { $map.Protocols['POP']        = $h }
        elseif ($l -match '\bimap4?\b')                { $map.Protocols['IMAP']       = $h }
        elseif ($l -match 'activesync|\beas\b')        { $map.Protocols['ActiveSync'] = $h }
        elseif ($l -match 'ews|web services')          { $map.Protocols['EWS']        = $h }
    }
    if ($map.Protocols.Count -gt 0 -and $map.Kind -eq 'Unknown') { $map.Kind = 'AppUsage' }

    return $map
}

function Get-UsageByUser {
    <#
        user -> the set of protocols they were seen using.

        Both export shapes collapse to the same answer, which is what lets the
        rest of this tool not care which one you exported.
    #>
    param($Rows, $Map)

    $used = @{}
    if (-not $Map -or -not $Map.User) { return $used }

    foreach ($r in @($Rows)) {
        $u = "$($r.$($Map.User))".Trim().ToLower()
        if (-not $u) { continue }
        if (-not $used.ContainsKey($u)) { $used[$u] = New-Object System.Collections.Generic.List[string] }

        if ($Map.Kind -eq 'SignIn') {
            $app = "$($r.$($Map.AppColumn))".ToLower()
            foreach ($p in $script:Protocols.Keys) {
                $hit = switch ($p) {
                    'POP'        { $app -match '\bpop' }
                    'IMAP'       { $app -match '\bimap' }
                    'ActiveSync' { $app -match 'activesync|\beas\b' }
                    'EWS'        { $app -match 'exchange web services|\bews\b' }
                }
                if ($hit -and -not $used[$u].Contains($p)) { $used[$u].Add($p) }
            }
        }
        else {
            foreach ($p in @($Map.Protocols.Keys)) {
                $v = "$($r.$($Map.Protocols[$p]))".Trim()
                # A DATE OR A NUMBER MEANS USED; blank, 0 or 'No' means not.
                if ($v -and $v -notmatch '^(0|no|false|n/a|-)$') {
                    if (-not $used[$u].Contains($p)) { $used[$u].Add($p) }
                }
            }
        }
    }
    return $used
}
#endregion

#region Enabled state --------------------------------------------------------
function Get-EnabledState {
    <#
        Every mailbox, and which of these protocols is switched on.

        Get-CASMailbox exists in BOTH Exchange editions (C1), so it goes
        through Invoke-EXOCommand -- an on-premises answer here would report a
        different population entirely.

        ResultSize explicit (C3): this enumerates, and the 1000-item default
        would quietly cap a tenant-wide report at a thousand mailboxes.
    #>
    param([scriptblock]$Progress)

    if ($Progress) { & $Progress 'Reading protocol settings for every mailbox...' }
    $rows = @(Invoke-EXOCommand -Name 'Get-CASMailbox' -Parameters @{
        ResultSize  = 'Unlimited'
        ErrorAction = 'Stop'
    })
    if ($Progress) { & $Progress "Mailboxes read: $($rows.Count)" }
    return $rows
}
#endregion

#region Report ---------------------------------------------------------------
function Get-ProtocolUsageReport {
    param([string]$CsvPath, [scriptblock]$Progress)

    $usageRows = @()
    $map       = $null
    $used      = @{}

    if ($CsvPath) {
        if ($Progress) { & $Progress 'Reading the usage export...' }
        $usageRows = @(Import-Csv -LiteralPath $CsvPath)
        $map       = Get-UsageColumnMap -Rows $usageRows
        $used      = Get-UsageByUser -Rows $usageRows -Map $map
    }

    $mailboxes = Get-EnabledState -Progress $Progress
    if ($Progress) { & $Progress 'Cross-referencing...' }

    $rows    = New-Object System.Collections.Generic.List[object]
    $enabled = @{}; $usedCount = @{}
    foreach ($p in $script:Protocols.Keys) { $enabled[$p] = 0; $usedCount[$p] = 0 }

    foreach ($m in $mailboxes) {
        $addr = "$($m.PrimarySmtpAddress)".Trim().ToLower()
        $seen = if ($addr -and $used.ContainsKey($addr)) { @($used[$addr].ToArray()) } else { @() }

        $row = [ordered]@{
            PrimarySmtpAddress = "$($m.PrimarySmtpAddress)"
            DisplayName        = "$($m.DisplayName)"
        }
        $anyEnabledUnused = @()
        foreach ($p in $script:Protocols.Keys) {
            $on = [bool]$m.$($script:Protocols[$p])
            $wasUsed = ($seen -contains $p)
            if ($on)      { $enabled[$p]++ }
            if ($wasUsed) { $usedCount[$p]++ }
            $row["${p}_Enabled"] = $on
            $row["${p}_Used"]    = $wasUsed
            if ($on -and -not $wasUsed) { $anyEnabledUnused += $p }
        }
        $row['EnabledButUnused'] = ($anyEnabledUnused -join ', ')
        $rows.Add([PSCustomObject]$row)
    }

    return [PSCustomObject]@{
        Rows         = $rows.ToArray()
        Map          = $map
        UsageRows    = @($usageRows).Count
        MailboxCount = @($mailboxes).Count
        Enabled      = $enabled
        Used         = $usedCount
        CsvPath      = $CsvPath
    }
}

function Format-ChangeRequestSummary {
    <#
        The paragraph that goes in the ticket.

        Written to be pasted, not read off a screen -- which is why it states
        the evidence window in the same breath as the numbers. A reviewer's
        first question is always "over what period", and answering it before
        it is asked is most of getting a change approved.
    #>
    param($Result)

    $sb = New-Object System.Text.StringBuilder

    [void]$sb.AppendLine('=== EVIDENCE SUMMARY -- paste into the change request ===')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("Mailboxes examined: $($Result.MailboxCount)")
    if ($Result.CsvPath) {
        [void]$sb.AppendLine("Usage export      : $(Split-Path -Leaf $Result.CsvPath)  ($($Result.UsageRows) row(s))")
    } else {
        [void]$sb.AppendLine('Usage export      : NONE PROVIDED -- the "used" columns below are all false')
        [void]$sb.AppendLine('                    because nothing was read, NOT because nothing was used.')
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Protocol      Enabled on      Used by      Enabled with no user')
    foreach ($p in $script:Protocols.Keys) {
        $en = [int]$Result.Enabled[$p]
        $us = [int]$Result.Used[$p]
        [void]$sb.AppendLine(
            '  ' + "$p".PadRight(12) + "$en".PadLeft(10) + "$us".PadLeft(13) + "$($en - $us)".PadLeft(21))
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('SUPPORTING FACT, INDEPENDENT OF THE LOGS ABOVE:')
    [void]$sb.AppendLine('  Microsoft removed Basic authentication for POP, IMAP, ActiveSync and')
    [void]$sb.AppendLine('  EWS in late 2022. Password-based access over these protocols has been')
    [void]$sb.AppendLine('  impossible since then regardless of this setting. What the setting')
    [void]$sb.AppendLine('  still permits is OAuth-based access, which is what the figures above')
    [void]$sb.AppendLine('  measure.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('EVIDENCE WINDOW -- state this, do not let it be asked:')
    [void]$sb.AppendLine('  Entra sign-in logs retain 30 days. The Email App Usage report reaches')
    [void]$sb.AppendLine('  180 days. Neither covers a year unless log shipping was already in')
    [void]$sb.AppendLine('  place. Quote the window of the export you used.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('ALSO IN THE CHANGE, OR IT UNDOES ITSELF:')
    [void]$sb.AppendLine('  Get-CASMailboxPlan | Set-CASMailboxPlan -PopEnabled $false -ImapEnabled $false')
    [void]$sb.AppendLine('  Mailbox plans set the defaults NEW mailboxes inherit. Without this,')
    [void]$sb.AppendLine('  every new starter arrives with the protocols back on. Set every plan,')
    [void]$sb.AppendLine('  not just the one flagged IsDefault.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('BEFORE DISABLING ACTIVESYNC, SEPARATELY FROM POP AND IMAP:')
    [void]$sb.AppendLine('  ActiveSync is what the built-in iOS and Android mail apps use. Outlook')
    [void]$sb.AppendLine('  mobile does not need it. If the numbers above show ActiveSync in use,')
    [void]$sb.AppendLine('  those are real people on native mail apps and they lose mail on cutover.')

    if ($Result.Map) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('--- How the export was read ---')
        [void]$sb.AppendLine("  Detected as     : $($Result.Map.Kind)")
        [void]$sb.AppendLine("  User column     : $(if ($Result.Map.User) { $Result.Map.User } else { 'NOT FOUND -- no usage could be matched' })")
        if ($Result.Map.Kind -eq 'SignIn') {
            [void]$sb.AppendLine("  Client app col  : $($Result.Map.AppColumn)")
        } elseif ($Result.Map.Protocols.Count -gt 0) {
            foreach ($k in @($Result.Map.Protocols.Keys)) {
                [void]$sb.AppendLine("  $($k.PadRight(14)): $($Result.Map.Protocols[$k])")
            }
        }
        [void]$sb.AppendLine('  Headers seen    :')
        foreach ($h in @($Result.Map.Headers)) { [void]$sb.AppendLine("      $h") }
        [void]$sb.AppendLine('  If the columns above look wrong, the numbers are wrong. Say so')
        [void]$sb.AppendLine('  rather than filing them.')
    }

    return $sb.ToString()
}
#endregion

#region UI -------------------------------------------------------------------
$form             = New-Object System.Windows.Forms.Form
$form.Text        = 'Protocol Usage'
$form.ClientSize  = New-Object System.Drawing.Size(820, 700)
$form.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblIntro          = New-Object System.Windows.Forms.Label
$lblIntro.Text     = 'Protocols.ps1 says what is ENABLED. This says what is USED, and writes the evidence for the change request. Export a sign-in report from Entra (30 days) or the Email App Usage report (180 days), then point this at the CSV.'
$lblIntro.Location = New-Object System.Drawing.Point(12, 12)
$lblIntro.Size     = New-Object System.Drawing.Size(796, 46)
$form.Controls.Add($lblIntro)

$btnSignIn          = New-Object System.Windows.Forms.Button
$btnSignIn.Text     = 'Open sign-in logs'
$btnSignIn.Location = New-Object System.Drawing.Point(12, 64)
$btnSignIn.Size     = New-Object System.Drawing.Size(140, 30)
$form.Controls.Add($btnSignIn)

$btnAppUsage          = New-Object System.Windows.Forms.Button
$btnAppUsage.Text     = 'Open usage report'
$btnAppUsage.Location = New-Object System.Drawing.Point(160, 64)
$btnAppUsage.Size     = New-Object System.Drawing.Size(140, 30)
$form.Controls.Add($btnAppUsage)

$txtCsv          = New-Object System.Windows.Forms.TextBox
$txtCsv.Location = New-Object System.Drawing.Point(12, 104)
$txtCsv.Size     = New-Object System.Drawing.Size(660, 24)
$txtCsv.ReadOnly = $true
$form.Controls.Add($txtCsv)

$btnBrowse          = New-Object System.Windows.Forms.Button
$btnBrowse.Text     = 'Browse...'
$btnBrowse.Location = New-Object System.Drawing.Point(680, 102)
$btnBrowse.Size     = New-Object System.Drawing.Size(128, 28)
$form.Controls.Add($btnBrowse)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run report'
$btnRun.Location = New-Object System.Drawing.Point(12, 140)
$btnRun.Size     = New-Object System.Drawing.Size(140, 30)
$form.Controls.Add($btnRun)

$btnCopy          = New-Object System.Windows.Forms.Button
$btnCopy.Text     = 'Copy summary'
$btnCopy.Location = New-Object System.Drawing.Point(160, 140)
$btnCopy.Size     = New-Object System.Drawing.Size(140, 30)
$form.Controls.Add($btnCopy)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(698, 140)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 180)
$lblStatus.Size     = New-Object System.Drawing.Size(796, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 200 -Width 796 -Height 12

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Location   = New-Object System.Drawing.Point(12, 220)
$txtOut.Size       = New-Object System.Drawing.Size(796, 430)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtOut.WordWrap   = $true
$form.Controls.Add($txtOut)

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'ProtocolUsage' `
                -X 12 -Y 660 -Width 796 -StatusLabel $lblStatus
#endregion

#region Handlers -------------------------------------------------------------
$btnClose.Add_Click({ $form.Close() })
# -Purpose 'Admin', NOT a switch. These portals need the elevated account and
# Start-Process alone would land in whichever Edge profile was used last --
# which for an admin portal is the wrong identity. The profile for each purpose
# is configured in Quick Launch Tool; 'Admin' is the name the map expects, and
# it is what M365 Reports, Message Trace and activate.ps1 all pass.
$btnSignIn.Add_Click({
    try {
        $st = Start-UrlInBrowserProfile -Url $SignInUrl -Purpose 'Admin'
        $lblStatus.Text = "Sign-in logs -- $st"
    } catch {
        $lblStatus.Text = "Could not open: $($_.Exception.Message)"
    }
})
$btnAppUsage.Add_Click({
    try {
        $st = Start-UrlInBrowserProfile -Url $AppUsageUrl -Purpose 'Admin'
        $lblStatus.Text = "Usage report -- $st"
    } catch {
        $lblStatus.Text = "Could not open: $($_.Exception.Message)"
    }
})
$btnCopy.Add_Click({
    try {
        if ("$($txtOut.Text)".Trim()) { [System.Windows.Forms.Clipboard]::SetText($txtOut.Text) }
        $btnCopy.Text = 'Copied!'
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Copy' }
})
$btnBrowse.Add_Click({
    try {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = 'CSV export (*.csv)|*.csv|All files (*.*)|*.*'
        if ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtCsv.Text = $dlg.FileName
        }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Browse' }
})

$btnRun.Add_Click({
    $exo = Test-ExoConnectionState
    if (-not $exo.Connected) {
        Show-ToolDialog(
            "Not connected to Exchange Online. Use the Connect button on the console first.`r`n`r`n$($exo.Detail)",
            'Exchange Online required', 'OK', 'Warning') | Out-Null
        return
    }
    # NO CSV IS ALLOWED, AND IT IS CALLED OUT LOUDLY. The enabled counts are
    # still worth having on their own -- but a summary that showed zeros in the
    # "used" column without saying why would be evidence of nothing, filed as
    # evidence of something.
    if (-not "$($txtCsv.Text)".Trim()) {
        $ans = Show-ToolDialog(
            "No usage export selected.`r`n`r`nThe report will show what is ENABLED but every 'used' figure will read zero -- because nothing was measured, not because nothing was used. That is not evidence.`r`n`r`nRun it anyway?",
            'No usage data', 'YesNo', 'Warning', 'Button2')
        if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    $btnRun.Enabled = $false
    $prev = $form.Cursor; $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = ''
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity 'Protocol usage'
        $report = Get-ProtocolUsageReport -CsvPath "$($txtCsv.Text)".Trim() -Progress $progress
        Complete-ProgressReporter -Activity 'Protocol usage' -Bar $pbar

        $txtOut.Text = Format-ChangeRequestSummary -Result $report

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir "ProtocolUsage_$stamp.csv"
        $report.Rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
        $files.SelectPath($csv)

        $lblStatus.Text = "Done. $($report.MailboxCount) mailbox(es)."
        $txtOut.Text += [Environment]::NewLine + 'Per-mailbox detail saved to:' +
                        [Environment]::NewLine + "  $csv"

        $result = if ($report.CsvPath) { 'Success' } else { 'Partial' }
        Write-CompletedLog -Action 'Report' -Target $csv -Result $result `
            -Detail ("Protocol usage across $($report.MailboxCount) mailbox(es)" +
                     $(if (-not $report.CsvPath) { ' -- NO usage export provided' } else { '' }))
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Protocol usage report'
        Write-FailedLog -Action 'Report' -Target 'Protocol Usage' -ErrorRecord $_
    } finally {
        Complete-ProgressReporter -Activity 'Protocol usage'
        $form.Cursor = $prev
        $btnRun.Enabled = $true
    }
})
#endregion

#region Show -----------------------------------------------------------------
$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
Set-FormResizable -Form $form
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Protocol Usage' | Out-Null

[void]$form.ShowDialog()
#endregion
