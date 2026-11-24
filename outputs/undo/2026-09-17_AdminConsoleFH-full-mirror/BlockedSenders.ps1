#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# ---------------------------------------------------------------------------
<#
    BlockedSenders.ps1
    Users the service has restricted from sending, and a guarded way to
    release them.

    The same list as Defender's "Restricted entities" page, and the same
    action -- but with the reason, the history, and a confirmation attached.

    Author   : Glenn Armbruster
    Revised  : 2026-08-12
#>
# TOOLKIT-HELP-BEGIN
# Shows every account the service has blocked from sending mail, and lets you
# release one.
#
# An account lands here after sending mail the service classified as spam. In
# practice that is usually a deliberate bulk send that went out the wrong way
# -- but it is ALSO what a compromised mailbox looks like, and the two are
# indistinguishable from this list alone. Find out which before releasing.
#
# WHAT YOU SEE BEFORE YOU ACT:
#
#   Reason         why the service blocked them, in its own words
#   Blocked        when it happened
#   Released here  how many times this toolkit has released this address
#                  before, and when it last did
#
# That last column is worth reading. A second release is a coincidence; a
# fourth is a pattern, and the pattern is usually a team that keeps forgetting
# the bulk-send tool. Releasing them again fixes today and not next week.
#
# MICROSOFT LIMITS HOW MANY TIMES AN ACCOUNT CAN BE RELEASED. Exceed it and the
# release fails and only a support case will clear it -- so the count is not
# only a habit tracker, it is a budget.
#
# The ticket reference you type is recorded with the release, in the audit log
# AND in the service, so "who released this and why" is answerable later.
#
# Releasing does not fix the cause. If the account was compromised, releasing
# it hands the mailbox back to whoever took it.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment
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

#region Data --------------------------------------------------------------------

function Get-BlockedSenderList {
    <#
        Every blocked sender, in ONE call.

        Get-BlockedSenderAddress with no arguments returns the whole list --
        the service keeps it, so this is not an enumeration over mailboxes and
        does not get slower as the tenant grows. Worth stating because the
        obvious assumption is the opposite, and that assumption is what stops
        people opening a tool while they are troubleshooting.
    #>
    # NOT -ErrorAction Stop, and that is the whole point.
    #
    # With nothing blocked, this cmdlet does not return an empty set -- it
    # writes "A server side error has occurred ... please try again after some
    # time", which -ErrorAction Stop then promotes into a terminating error and
    # a red stack trace. The commonest state of a healthy tenant looked like a
    # broken tool.
    #
    # The error is captured rather than thrown, because it means one of two
    # things and the difference matters:
    #
    #   nothing is blocked        -- the normal, good case
    #   the service is unhappy    -- retry in a bit
    #
    # Nothing in the response distinguishes them, so the tool says so instead
    # of picking one and being confidently wrong.
    $err = $null
    $rows = @(Invoke-EXOCommand -Name 'Get-BlockedSenderAddress' -Parameters @{
        ErrorAction   = 'SilentlyContinue'
        ErrorVariable = 'exoErr'
    } 2>$null)

    if ($rows.Count -gt 0) {
        return [PSCustomObject]@{ Ok = $true; Rows = $rows; Message = '' }
    }

    $detail = ''
    if ($Error.Count -gt 0) { $detail = "$($Error[0].Exception.Message)".Trim() }

    if ($detail -match '(?i)server side error|try again') {
        return [PSCustomObject]@{
            Ok      = $true
            Rows    = @()
            Message = ("No blocked senders returned. Two things look identical here: nobody is " +
                       "blocked (the usual case, and what the Defender page shows as 0 items), or " +
                       "the service is briefly unavailable. Press Refresh in a minute to tell them " +
                       "apart." + [Environment]::NewLine + [Environment]::NewLine +
                       "Service said: $detail")
        }
    }

    if ($detail) {
        return [PSCustomObject]@{ Ok = $false; Rows = @(); Message = $detail }
    }
    return [PSCustomObject]@{ Ok = $true; Rows = @(); Message = 'Nobody is blocked.' }
}

function Get-ReleaseHistory {
    <#
        How many times THIS TOOLKIT has released an address, and when it last
        did.

        Microsoft does not expose its own unblock counter, so this is the only
        count available -- and it is honest about its scope: releases made in
        the portal, or by someone else's console, are not in here. The column
        header says "Released here" for exactly that reason.

        Read from the audit log rather than tracked separately. The log already
        records it, and a second copy of a fact is a copy that drifts.
    #>
    $out = @{}
    try {
        $dir = Get-LogPath
        if (-not $dir -or -not (Test-Path -LiteralPath $dir)) { return $out }
        foreach ($lf in @(Get-ChildItem -LiteralPath $dir -Filter 'CompletedTasks*.csv' -File -ErrorAction SilentlyContinue)) {
            foreach ($row in (Import-Csv -LiteralPath $lf.FullName -ErrorAction SilentlyContinue)) {
                if ("$($row.Action)" -ne 'ReleaseBlockedSender') { continue }
                if ("$($row.Result)" -ne 'Success') { continue }
                $key = "$($row.Target)".Trim().ToLower()
                if (-not $key) { continue }
                $when = $null
                try { $when = [datetimeoffset]::Parse("$($row.Timestamp)") } catch { }
                if (-not $out.ContainsKey($key)) {
                    $out[$key] = [PSCustomObject]@{ Count = 0; Last = $null }
                }
                $out[$key].Count++
                if ($when -and (-not $out[$key].Last -or $when -gt $out[$key].Last)) { $out[$key].Last = $when }
            }
        }
    } catch { }
    return $out
}

function Get-BlockedRow {
    <#
        The service's record joined to our release history.

        Properties are read DEFENSIVELY. This object's shape is not documented
        on the cmdlet page, and a tool that assumes a property name shows an
        empty column when the service renames one -- which reads as "no reason
        given" rather than "this tool is out of date".
    #>
    param($Entry, $History)

    $prop = { param($o, [string[]]$names)
        foreach ($n in $names) {
            if ($o.PSObject.Properties[$n] -and "$($o.$n)".Trim()) { return "$($o.$n)".Trim() }
        }
        return ''
    }

    $addr = & $prop $Entry @('SenderAddress', 'Sender', 'Identity')
    $key  = $addr.ToLower()
    $h    = $(if ($History.ContainsKey($key)) { $History[$key] } else { $null })

    return [PSCustomObject]@{
        Address  = $addr
        Reason   = (& $prop $Entry @('Reason', 'BlockReason', 'Description'))
        Blocked  = (& $prop $Entry @('CreatedDatetime', 'CreatedDateTime', 'WhenCreated', 'Created'))
        Released = $(if ($h) { $h.Count } else { 0 })
        LastRel  = $(if ($h -and $h.Last) { $h.Last.ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { '' })
        Raw      = $Entry
    }
}

#endregion

#region Form --------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Blocked Senders -- restricted from sending'
$form.Size = New-Object System.Drawing.Size(960, 700)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblTop = New-Object System.Windows.Forms.Label
$lblTop.Text = 'Accounts the service has stopped from sending mail. Usually a bulk send that went out the wrong way -- but it looks identical to a compromised mailbox, so check before releasing.'
$lblTop.Location = New-Object System.Drawing.Point(12, 12)
$lblTop.Size = New-Object System.Drawing.Size(920, 32)
$lblTop.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblTop)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = 'Refresh'
$btnRefresh.Location = New-Object System.Drawing.Point(12, 52)
$btnRefresh.Size = New-Object System.Drawing.Size(110, 28)
$form.Controls.Add($btnRefresh)

$lblFind = New-Object System.Windows.Forms.Label
$lblFind.Text = 'Find:'
$lblFind.Location = New-Object System.Drawing.Point(136, 58)
$lblFind.Size = New-Object System.Drawing.Size(36, 18)
$form.Controls.Add($lblFind)

$txtFind = New-Object System.Windows.Forms.TextBox
$txtFind.Location = New-Object System.Drawing.Point(176, 55)
$txtFind.Size = New-Object System.Drawing.Size(260, 22)
$form.Controls.Add($txtFind)

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Text = 'Not loaded yet.'
$lblCount.Location = New-Object System.Drawing.Point(450, 58)
$lblCount.Size = New-Object System.Drawing.Size(482, 18)
$form.Controls.Add($lblCount)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(12, 88)
$grid.Size = New-Object System.Drawing.Size(920, 250)
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.SelectionMode = 'FullRowSelect'
$grid.MultiSelect = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.RowHeadersVisible = $false
[void]$grid.Columns.Add('Address', 'Address')
[void]$grid.Columns.Add('Reason', 'Reason the service gave')
[void]$grid.Columns.Add('Blocked', 'Blocked')
[void]$grid.Columns.Add('Released', 'Released here')
[void]$grid.Columns.Add('LastRel', 'Last released')
$grid.Columns['Address'].FillWeight  = 26
$grid.Columns['Reason'].FillWeight   = 34
$grid.Columns['Blocked'].FillWeight  = 16
$grid.Columns['Released'].FillWeight = 10
$grid.Columns['LastRel'].FillWeight  = 14
$form.Controls.Add($grid)

$lblDetail = New-Object System.Windows.Forms.Label
$lblDetail.Text = 'Everything the service holds for the selected row:'
$lblDetail.Location = New-Object System.Drawing.Point(12, 346)
$lblDetail.Size = New-Object System.Drawing.Size(400, 18)
$form.Controls.Add($lblDetail)

$txtDetail = New-Object System.Windows.Forms.TextBox
$txtDetail.Location = New-Object System.Drawing.Point(12, 368)
$txtDetail.Size = New-Object System.Drawing.Size(920, 120)
$txtDetail.Multiline = $true
$txtDetail.ReadOnly = $true
$txtDetail.ScrollBars = 'Vertical'
$txtDetail.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtDetail)

$lblTicket = New-Object System.Windows.Forms.Label
$lblTicket.Text = 'Ticket / reason:'
$lblTicket.Location = New-Object System.Drawing.Point(12, 504)
$lblTicket.Size = New-Object System.Drawing.Size(100, 18)
$form.Controls.Add($lblTicket)

$txtTicket = New-Object System.Windows.Forms.TextBox
$txtTicket.Location = New-Object System.Drawing.Point(116, 501)
$txtTicket.Size = New-Object System.Drawing.Size(320, 22)
$form.Controls.Add($txtTicket)

$btnRelease = New-Object System.Windows.Forms.Button
$btnRelease.Text = 'Release selected'
$btnRelease.Location = New-Object System.Drawing.Point(450, 499)
$btnRelease.Size = New-Object System.Drawing.Size(140, 28)
$btnRelease.Enabled = $false
$form.Controls.Add($btnRelease)

$lblRelNote = New-Object System.Windows.Forms.Label
$lblRelNote.Text = 'Required. It is written to the service and to the audit log.'
$lblRelNote.Location = New-Object System.Drawing.Point(600, 505)
$lblRelNote.Size = New-Object System.Drawing.Size(332, 18)
$lblRelNote.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblRelNote)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(12, 534)
$txtOut.Size = New-Object System.Drawing.Size(920, 62)
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(832, 604)
$btnClose.Size = New-Object System.Drawing.Size(100, 28)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

#endregion

#region Behaviour ---------------------------------------------------------------

$script:Rows = @()

$renderGrid = {
    $needle = "$($txtFind.Text)".Trim()
    $show = @($script:Rows | Where-Object {
        -not $needle -or
        "$($_.Address)" -like "*$needle*" -or "$($_.Reason)" -like "*$needle*"
    })

    $grid.Rows.Clear()
    foreach ($r in $show) {
        $i = $grid.Rows.Add()
        $grid.Rows[$i].Cells['Address'].Value  = $r.Address
        $grid.Rows[$i].Cells['Reason'].Value   = $r.Reason
        $grid.Rows[$i].Cells['Blocked'].Value  = $r.Blocked
        $grid.Rows[$i].Cells['Released'].Value = $r.Released
        $grid.Rows[$i].Cells['LastRel'].Value  = $r.LastRel
        $grid.Rows[$i].Tag = $r

        # Repeat offenders are coloured, because the number that matters is the
        # one you were not looking for. Third release is where a habit stops
        # being an accident, and where the Microsoft limit starts to matter.
        if ([int]$r.Released -ge 3) {
            $grid.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Firebrick
        } elseif ([int]$r.Released -gt 0) {
            $grid.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 110, 0)
        }
    }

    $lblCount.Text = $(if ($script:Rows.Count -eq 0) {
        'Nobody is blocked. Nothing to do.'
    } elseif ($needle) {
        "$($show.Count) of $($script:Rows.Count) blocked account(s) match '$needle'."
    } else {
        "$($script:Rows.Count) blocked account(s)."
    })
}

$loadList = {
    $form.Cursor = 'WaitCursor'
    $btnRefresh.Enabled = $false
    try {
        $history = Get-ReleaseHistory
        $result  = Get-BlockedSenderList

        $script:Rows = @($result.Rows | ForEach-Object { Get-BlockedRow -Entry $_ -History $history } |
                         Sort-Object -Property @{ Expression = 'Released'; Descending = $true }, Address)
        & $renderGrid
        $txtDetail.Text = ''
        $btnRelease.Enabled = $false

        # An empty list is a RESULT, not a failure, so it goes in the output box
        # as plain text rather than as an error. A red panel on the healthy case
        # trains people to ignore the red.
        if (-not $result.Ok) {
            $lblCount.Text = 'The service returned an error.'
            $txtOut.Text = $result.Message
        } elseif ($result.Message) {
            $txtOut.Text = $result.Message
        } else {
            $txtOut.Text = ''
        }

        if ("$SmtpAddress".Trim()) { $txtFind.Text = "$SmtpAddress".Trim(); & $renderGrid }
    } catch {
        $lblCount.Text = 'Could not read the blocked list.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Reading blocked senders'
    } finally {
        $btnRefresh.Enabled = $true
        $form.Cursor = 'Default'
    }
}

$btnRefresh.Add_Click($loadList)
$txtFind.Add_TextChanged($renderGrid)

$grid.Add_SelectionChanged({
    $btnRelease.Enabled = ($grid.SelectedRows.Count -gt 0)
    if ($grid.SelectedRows.Count -eq 0) { $txtDetail.Text = ''; return }
    $r = $grid.SelectedRows[0].Tag
    if (-not $r) { return }

    # Every property the service returned, not a chosen few. The reason a
    # mailbox was blocked is sometimes in a field this tool does not know the
    # name of, and a troubleshooting window should not hide it.
    $sb = New-Object System.Text.StringBuilder
    foreach ($p in @($r.Raw.PSObject.Properties | Sort-Object Name)) {
        [void]$sb.AppendLine(("{0,-28} {1}" -f $p.Name, "$($p.Value)"))
    }
    if ([int]$r.Released -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine(("{0,-28} {1} time(s), last {2}" -f 'Released by this toolkit', $r.Released, $r.LastRel))
    }
    $txtDetail.Text = $sb.ToString()
})

$btnRelease.Add_Click({
    if ($grid.SelectedRows.Count -eq 0) { return }
    $r = $grid.SelectedRows[0].Tag
    if (-not $r -or -not $r.Address) { return }

    $ticket = "$($txtTicket.Text)".Trim()
    if (-not $ticket) {
        $txtOut.Text = 'Enter a ticket reference or reason first. It is recorded with the release.'
        $txtTicket.Focus()
        return
    }

    # The confirmation states the case AGAINST releasing, not just the action.
    # A dialog that only says "are you sure" is answered without reading; one
    # that names the reason and the count is answered by thinking.
    $warn = ''
    if ([int]$r.Released -ge 3) {
        $warn = [Environment]::NewLine + [Environment]::NewLine +
                "This address has been released $($r.Released) times before. Microsoft limits how " +
                "many times an account can be released -- past that, only a support case clears it. " +
                "Releasing again treats the symptom."
    } elseif ([int]$r.Released -gt 0) {
        $warn = [Environment]::NewLine + [Environment]::NewLine +
                "Released $($r.Released) time(s) before, most recently $($r.LastRel)."
    }

    $answer = Show-ToolDialog(
        ("Release this account so it can send mail again?" + [Environment]::NewLine + [Environment]::NewLine +
         "Address : $($r.Address)" + [Environment]::NewLine +
         "Reason  : $($(if ($r.Reason) { $r.Reason } else { '(none recorded)' }))" + [Environment]::NewLine +
         "Blocked : $($r.Blocked)" + [Environment]::NewLine +
         "Ticket  : $ticket" + $warn + [Environment]::NewLine + [Environment]::NewLine +
         "If the account was compromised rather than sending in bulk, this hands " +
         "the mailbox back to whoever took it. Confirm the cause first."),
        'Release blocked sender', 'YesNo', 'Warning', 'Button2')
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $form.Cursor = 'WaitCursor'
    $btnRelease.Enabled = $false
    try {
        Invoke-EXOCommand -Name 'Remove-BlockedSenderAddress' -Parameters @{
            SenderAddress = $r.Address
            Reason        = $ticket
            ErrorAction   = 'Stop'
        }
        $txtOut.Text = ("RELEASED" + [Environment]::NewLine + [Environment]::NewLine +
                        "Address : $($r.Address)" + [Environment]::NewLine +
                        "Ticket  : $ticket" + [Environment]::NewLine + [Environment]::NewLine +
                        "It can take a short while before sending works again. The cause is " +
                        "unchanged -- if this was a bulk send, point them at the tool for it.")
        Write-CompletedLog -Action 'ReleaseBlockedSender' -Target $r.Address -Result 'Success' `
            -Detail "$ticket; service reason: $($r.Reason); prior releases: $($r.Released)"
        & $loadList
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Releasing the blocked sender'
        Write-FailedLog -Action 'ReleaseBlockedSender' -Target $r.Address -ErrorRecord $_ `
            -Detail "$ticket; prior releases: $($r.Released)"
    } finally {
        $form.Cursor = 'Default'
        $btnRelease.Enabled = ($grid.SelectedRows.Count -gt 0)
    }
})

$form.Add_Shown($loadList)

Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Blocked Senders' | Out-Null
Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'

[void]$form.ShowDialog()
$form.Dispose()
#endregion
