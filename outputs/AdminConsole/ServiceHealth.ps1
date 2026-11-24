#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-GRAPH-SCOPES: ServiceHealth.Read.All, ServiceMessage.Read.All
<#
    ServiceHealth.ps1
    Is it us, or is it Microsoft? Service health and the message centre,
    without opening the admin portal.

    Author   : Glenn Armbruster
    Created  : 2026-09-17
#>
# TOOLKIT-HELP-RUNS: Reads Microsoft 365 service health and message centre posts. Changes nothing.
# TOOLKIT-HELP-BEGIN
# Three things from the Microsoft 365 admin centre, on one screen:
#
#   Service status    Every service (Exchange Online, Teams, SharePoint, Entra,
#                     Intune...) and whether Microsoft says it is healthy right
#                     now. Anything not operational is coloured.
#
#   Incidents and advisories   Open issues, newest first, with the service,
#                     the user impact and Microsoft's latest update. Tick
#                     "include resolved" to see what closed in the last week.
#
#   Message centre    Planned changes and announcements. Filter to the ones
#                     marked major, or that need action by a date.
#
# When a user reports mail or Teams behaving oddly, this is the two-second
# check before anyone starts a trace: if Exchange Online shows an incident,
# the answer is already here.
#
# NEEDS: Graph scopes ServiceHealth.Read.All and ServiceMessage.Read.All, and
# an admin role on YOUR account (Microsoft applies role-based access to this
# data; Global Reader or Service Support Administrator is enough).
#
# Nothing here changes anything.
# TOOLKIT-HELP-END

param(
    [string]$Environment
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

$script:Scopes = @('ServiceHealth.Read.All', 'ServiceMessage.Read.All')
$ReportDir     = Get-ReportPath
$ReportPrefix  = 'ServiceHealth'

#region Data ------------------------------------------------------------------

function Get-ServiceOverview {
    # Learn 2026-09-17 (List healthOverviews): GET /admin/serviceAnnouncement/healthOverviews -> service, status, id.
    $r = Invoke-ToolkitGraph -Uri 'admin/serviceAnnouncement/healthOverviews'
    if ($r -and $r.PSObject.Properties['value']) { return @($r.value) }
    return @()
}

function Get-ServiceIssue {
    <#
        Learn 2026-09-17 (List issues): GET /admin/serviceAnnouncement/issues,
        OData $filter; properties id, title, service, feature, status,
        classification (advisory | incident), startDateTime, endDateTime,
        lastModifiedDateTime, isResolved, impactDescription, posts[].
    #>
    param([switch]$IncludeResolved, [int]$ResolvedDays = 7)
    $f = 'isResolved eq false'
    if ($IncludeResolved) {
        $since = (Get-Date).ToUniversalTime().AddDays(-1 * $ResolvedDays).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $f = "isResolved eq false or lastModifiedDateTime ge $since"
    }
    $uri = "admin/serviceAnnouncement/issues?`$filter=" + [uri]::EscapeDataString($f)
    try {
        return @(Invoke-ToolkitGraph -Uri $uri -All -MaxRows 1000)
    } catch {
        # A rejected filter is not a reason to show nothing: pull the lot and
        # filter here. It is a few hundred rows at most.
        $all = @(Invoke-ToolkitGraph -Uri 'admin/serviceAnnouncement/issues' -All -MaxRows 2000)
        if ($IncludeResolved) {
            $cut = (Get-Date).ToUniversalTime().AddDays(-1 * $ResolvedDays)
            # Parse inside the filter, not in the comparison: an issue with no
            # lastModifiedDateTime would throw from inside Where-Object and
            # take the whole grid with it.
            return @($all | Where-Object {
                if (-not $_.isResolved) { return $true }
                $d = $null
                try { $d = [DateTimeOffset]::Parse("$($_.lastModifiedDateTime)").UtcDateTime } catch { }
                return ($null -ne $d -and $d -ge $cut)
            })
        }
        return @($all | Where-Object { -not $_.isResolved })
    }
}

function Get-ServiceMessage {
    <#
        Learn 2026-09-17 (List messages): GET /admin/serviceAnnouncement/messages;
        properties id, title, category (preventOrFixIssue | planForChange |
        stayInformed), severity (normal | high | critical), services[],
        startDateTime, lastModifiedDateTime, actionRequiredByDateTime,
        isMajorChange, body{content}.
    #>
    param([int]$Days = 30)
    $since = (Get-Date).ToUniversalTime().AddDays(-1 * $Days).ToString('yyyy-MM-ddTHH:mm:ssZ')
    $uri = "admin/serviceAnnouncement/messages?`$filter=" + [uri]::EscapeDataString("lastModifiedDateTime ge $since")
    try {
        return @(Invoke-ToolkitGraph -Uri $uri -All -MaxRows 1000)
    } catch {
        $all = @(Invoke-ToolkitGraph -Uri 'admin/serviceAnnouncement/messages' -All -MaxRows 2000)
        $cut = (Get-Date).ToUniversalTime().AddDays(-1 * $Days)
        return @($all | Where-Object {
            $d = $null
            try { $d = [DateTimeOffset]::Parse("$($_.lastModifiedDateTime)").UtcDateTime } catch { }
            return ($null -ne $d -and $d -ge $cut)
        })
    }
}

function ConvertFrom-HtmlText {
    # Message bodies are HTML. Enough stripping to read in a text box; not a renderer.
    param([string]$Html)
    $t = "$Html"
    $t = $t -replace '(?i)<br\s*/?>', "`r`n" -replace '(?i)</p>', "`r`n`r`n" -replace '(?i)</li>', "`r`n" -replace '(?i)<li[^>]*>', '  - '
    $t = $t -replace '<[^>]+>', ''
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    return ($t -replace '(\r?\n){3,}', "`r`n`r`n").Trim()
}

function Get-IssueDetailText {
    param($i)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("$($i.title)")
    [void]$sb.AppendLine("$($i.id)   $($i.classification)   service: $($i.service)   feature: $($i.feature)")
    [void]$sb.AppendLine("status: $($i.status)   started: $(ConvertTo-LocalGraphTime $i.startDateTime)   last update: $(ConvertTo-LocalGraphTime $i.lastModifiedDateTime)$(if ($i.isResolved) { "   RESOLVED $(ConvertTo-LocalGraphTime $i.endDateTime)" })")
    [void]$sb.AppendLine('')
    if ("$($i.impactDescription)".Trim()) {
        [void]$sb.AppendLine("User impact: $($i.impactDescription)")
        [void]$sb.AppendLine('')
    }
    $posts = @()
    if ($i.PSObject.Properties['posts']) { $posts = @($i.posts | Sort-Object @{ Expression = { "$($_.createdDateTime)" }; Descending = $true }) }
    if ($posts.Count -gt 0) {
        [void]$sb.AppendLine("Latest update ($(ConvertTo-LocalGraphTime $posts[0].createdDateTime), $($posts[0].postType)):")
        [void]$sb.AppendLine((ConvertFrom-HtmlText "$($posts[0].description.content)"))
        if ($posts.Count -gt 1) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine("($($posts.Count - 1) earlier update(s) not shown)")
        }
    }
    return $sb.ToString()
}

function Get-MessageDetailText {
    param($m)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("$($m.title)")
    [void]$sb.AppendLine("$($m.id)   $($m.category)   severity: $($m.severity)   major change: $($m.isMajorChange)")
    [void]$sb.AppendLine("services: $(@($m.services) -join ', ')")
    [void]$sb.AppendLine("published: $(ConvertTo-LocalGraphTime $m.startDateTime)   last modified: $(ConvertTo-LocalGraphTime $m.lastModifiedDateTime)$(if ("$($m.actionRequiredByDateTime)") { "   ACTION REQUIRED BY $(ConvertTo-LocalGraphTime $m.actionRequiredByDateTime 'yyyy-MM-dd')" })")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine((ConvertFrom-HtmlText "$($m.body.content)"))
    return $sb.ToString()
}
#endregion

#region UI --------------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Service Health'
$form.ClientSize    = New-Object System.Drawing.Size(1180, 740)
$form.StartPosition = 'Manual'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$y = 12
$btnLoad          = New-Object System.Windows.Forms.Button
$btnLoad.Text     = 'Refresh'
$btnLoad.Location = New-Object System.Drawing.Point(12, ($y - 2))
$btnLoad.Size     = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnLoad)

$chkResolved          = New-Object System.Windows.Forms.CheckBox
$chkResolved.Text     = 'Include issues resolved in the last 7 days'
$chkResolved.Location = New-Object System.Drawing.Point(124, ($y + 2))
$chkResolved.Size     = New-Object System.Drawing.Size(280, 20)
$form.Controls.Add($chkResolved)

$chkMajor          = New-Object System.Windows.Forms.CheckBox
$chkMajor.Text     = 'Message centre: major changes and action-required only'
$chkMajor.Location = New-Object System.Drawing.Point(410, ($y + 2))
$chkMajor.Size     = New-Object System.Drawing.Size(360, 20)
$form.Controls.Add($chkMajor)

$btnExport          = New-Object System.Windows.Forms.Button
$btnExport.Text     = 'Export issues CSV'
$btnExport.Location = New-Object System.Drawing.Point(800, ($y - 2))
$btnExport.Size     = New-Object System.Drawing.Size(140, 28)
$btnExport.Enabled  = $false
$form.Controls.Add($btnExport)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(1068, ($y - 2))
$btnClose.Size     = New-Object System.Drawing.Size(100, 28)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)
$y += 34

$lblStatus           = New-Object System.Windows.Forms.Label
$lblStatus.Location  = New-Object System.Drawing.Point(12, $y)
$lblStatus.Size      = New-Object System.Drawing.Size(1156, 18)
$lblStatus.ForeColor = [System.Drawing.Color]::DimGray
$lblStatus.Text      = 'Press Refresh.'
$form.Controls.Add($lblStatus)
$y += 20

$pb = Add-ReportProgressBar -Form $form -X 12 -Y $y -Width 1156 -Height 12
$y += 18

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix $ReportPrefix `
            -X 12 -Y $y -Width 1156 -StatusLabel $lblStatus
$y += 34

function New-HealthGrid {
    param([object[]]$Columns, [int]$Height, [int]$Top, [int]$Left = 12, [int]$Width = 1156)
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Location            = New-Object System.Drawing.Point($Left, $Top)
    $g.Size                = New-Object System.Drawing.Size($Width, $Height)
    $g.ReadOnly            = $true
    $g.AllowUserToAddRows  = $false
    $g.SelectionMode       = 'FullRowSelect'
    $g.MultiSelect         = $false
    $g.RowHeadersVisible   = $false
    $g.AutoSizeColumnsMode = 'None'
    $g.ScrollBars          = 'Both'
    $g.Font                = New-Object System.Drawing.Font('Consolas', 9)
    foreach ($c in $Columns) {
        $idx = $g.Columns.Add($c[0], $c[0])
        $g.Columns[$idx].Width    = $c[1]
        $g.Columns[$idx].SortMode = 'Automatic'
    }
    return $g
}

$lblA          = New-Object System.Windows.Forms.Label
$lblA.Text     = 'Service status:'
$lblA.Location = New-Object System.Drawing.Point(12, $y)
$lblA.Size     = New-Object System.Drawing.Size(300, 18)
$form.Controls.Add($lblA)

$lblB          = New-Object System.Windows.Forms.Label
$lblB.Text     = 'Incidents and advisories (pick one for the detail):'
$lblB.Location = New-Object System.Drawing.Point(420, $y)
$lblB.Size     = New-Object System.Drawing.Size(500, 18)
$form.Controls.Add($lblB)
$y += 20

$gridSvc = New-HealthGrid -Top $y -Height 200 -Left 12 -Width 396 -Columns @(@('Service', 250), @('Status', 120))
$form.Controls.Add($gridSvc)

$gridIss = New-HealthGrid -Top $y -Height 200 -Left 420 -Width 748 -Columns @(
    @('Updated', 118), @('Kind', 80), @('Service', 150), @('Title', 340), @('Status', 130), @('Id', 90))
$form.Controls.Add($gridIss)
$y += 208

$txtDetail            = New-Object System.Windows.Forms.TextBox
$txtDetail.Multiline  = $true
$txtDetail.ReadOnly   = $true
$txtDetail.ScrollBars = 'Vertical'
$txtDetail.WordWrap   = $true
$txtDetail.Location   = New-Object System.Drawing.Point(12, $y)
$txtDetail.Size       = New-Object System.Drawing.Size(1156, 150)
$txtDetail.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtDetail.BackColor  = [System.Drawing.Color]::WhiteSmoke
$txtDetail.ForeColor  = [System.Drawing.Color]::Black
$txtDetail.Tag        = 'NoTheme'
$form.Controls.Add($txtDetail)
$y += 158

$lblC          = New-Object System.Windows.Forms.Label
$lblC.Text     = 'Message centre, last 30 days (pick one for the detail):'
$lblC.Location = New-Object System.Drawing.Point(12, $y)
$lblC.Size     = New-Object System.Drawing.Size(600, 18)
$form.Controls.Add($lblC)
$y += 20

$gridMsg = New-HealthGrid -Top $y -Height (740 - $y - 12) -Columns @(
    @('Modified', 118), @('Severity', 80), @('Category', 130), @('Major', 60), @('Action by', 100), @('Services', 220), @('Title', 420))
$form.Controls.Add($gridMsg)

$form | Add-Member -NotePropertyName 'Busy' -NotePropertyValue $false -Force
$script:Issues = @()

$load = {
    $gridSvc.Rows.Clear(); $gridIss.Rows.Clear(); $gridMsg.Rows.Clear(); $txtDetail.Text = ''
    $script:Issues = @(); $btnExport.Enabled = $false
    $p = New-ProgressReporter -Label $lblStatus -Bar $pb -Activity 'Service health'

    & $p 'Reading service status...' -1
    $bad = 0
    foreach ($o in @(Get-ServiceOverview | Sort-Object service)) {
        $i = $gridSvc.Rows.Add("$($o.service)", "$($o.status)")
        $gridSvc.Rows[$i].Tag = $o
        if ("$($o.status)" -ne 'serviceOperational') {
            $bad++
            $gridSvc.Rows[$i].DefaultCellStyle.ForeColor = $(if ("$($o.status)" -match 'Degradation|Interruption|Restoring|Extended') { [System.Drawing.Color]::FromArgb(180, 0, 0) } else { [System.Drawing.Color]::FromArgb(190, 110, 0) })
        }
    }

    & $p 'Reading incidents and advisories...' -1
    $iss = @(Get-ServiceIssue -IncludeResolved:$chkResolved.Checked | Sort-Object @{ Expression = { "$($_.lastModifiedDateTime)" }; Descending = $true })
    $script:Issues = $iss
    foreach ($x in $iss) {
        $i = $gridIss.Rows.Add((ConvertTo-LocalGraphTime $x.lastModifiedDateTime), "$($x.classification)", "$($x.service)", "$($x.title)",
                $(if ($x.isResolved) { 'resolved' } else { "$($x.status)" }), "$($x.id)")
        $gridIss.Rows[$i].Tag = $x
        if (-not $x.isResolved -and "$($x.classification)" -eq 'incident') { $gridIss.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0) }
        elseif ($x.isResolved) { $gridIss.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::DimGray }
    }

    & $p 'Reading the message centre...' -1
    $msgs = @(Get-ServiceMessage -Days 30 | Sort-Object @{ Expression = { "$($_.lastModifiedDateTime)" }; Descending = $true })
    if ($chkMajor.Checked) { $msgs = @($msgs | Where-Object { $_.isMajorChange -or "$($_.actionRequiredByDateTime)" }) }
    foreach ($m in $msgs) {
        $i = $gridMsg.Rows.Add((ConvertTo-LocalGraphTime $m.lastModifiedDateTime), "$($m.severity)", "$($m.category)",
                $(if ($m.isMajorChange) { 'yes' } else { '' }), (ConvertTo-LocalGraphTime $m.actionRequiredByDateTime 'yyyy-MM-dd'),
                (@($m.services) -join ', '), "$($m.title)")
        $gridMsg.Rows[$i].Tag = $m
        if ("$($m.severity)" -match 'high|critical') { $gridMsg.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0) }
    }

    $open = @($iss | Where-Object { -not $_.isResolved }).Count
    & $p "$bad service(s) not operational; $open open issue(s); $($msgs.Count) message(s)." 100
    $btnExport.Enabled = ($iss.Count -gt 0)
    if ($gridIss.Rows.Count -gt 0) { $gridIss.Rows[0].Selected = $true }
}

$btnLoad.Add_Click({
    if ($form.Busy) { return }
    $form.Busy = $true
    try { $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; & $load }
    catch { $lblStatus.Text = "Refresh failed: $($_.Exception.Message)"; Show-HandlerError -ErrorRecord $_ -Context 'Service health (Graph serviceAnnouncement)' }
    finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
})
$chkResolved.Add_CheckedChanged({ if (-not $form.Busy -and $gridSvc.Rows.Count) { $btnLoad.PerformClick() } })
$chkMajor.Add_CheckedChanged({ if (-not $form.Busy -and $gridSvc.Rows.Count) { $btnLoad.PerformClick() } })

$gridIss.Add_SelectionChanged({
    if ($gridIss.SelectedRows.Count -eq 0) { return }
    $x = $gridIss.SelectedRows[0].Tag
    if ($x) { $txtDetail.Text = (Get-IssueDetailText $x) }
})
$gridMsg.Add_SelectionChanged({
    if ($gridMsg.SelectedRows.Count -eq 0) { return }
    $m = $gridMsg.SelectedRows[0].Tag
    if ($m) { $txtDetail.Text = (Get-MessageDetailText $m) }
})

$btnExport.Add_Click({
    try {
        if ($script:Issues.Count -eq 0) { $lblStatus.Text = 'Nothing to export. Refresh first.'; return }
        $export = New-Object System.Collections.Generic.List[object]
        foreach ($x in $script:Issues) {
            $export.Add([PSCustomObject]@{
                Id             = "$($x.id)"
                Classification = "$($x.classification)"
                Service        = "$($x.service)"
                Feature        = "$($x.feature)"
                Title          = "$($x.title)"
                Status         = "$($x.status)"
                Resolved       = "$($x.isResolved)"
                Started        = (ConvertTo-LocalGraphTime $x.startDateTime)
                LastModified   = (ConvertTo-LocalGraphTime $x.lastModifiedDateTime)
                Ended          = (ConvertTo-LocalGraphTime $x.endDateTime)
                Impact         = "$($x.impactDescription)"
            })
        }
        if (-not (Test-Path -LiteralPath $ReportDir)) { New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir ("{0}_{1}.csv" -f $ReportPrefix, $stamp)
        $export | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
        $files.SelectPath($csv)
        $lblStatus.Text = "Exported $($export.Count) issue(s) to $csv"
        Write-CompletedLog -Action 'Report' -Target $csv -Result 'Success' -Detail "Service health issues: $($export.Count) row(s)." -Environment 'Cloud'
    } catch {
        $lblStatus.Text = 'Export failed.'
        Show-HandlerError -ErrorRecord $_ -Context 'Export service health'
    }
})

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Service Health' | Out-Null

$form.Add_Shown({
    $gate = Assert-ToolkitGraphScope -Scopes $script:Scopes
    if (-not $gate.Ok) {
        $lblStatus.Text = $gate.Message
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(200, 0, 0)
        $btnLoad.Enabled = $false
        return
    }
    $btnLoad.PerformClick()
})

[void]$form.ShowDialog()
$form.Dispose()
#endregion
