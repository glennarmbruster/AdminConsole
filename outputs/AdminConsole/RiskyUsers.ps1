#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-GRAPH-SCOPES: IdentityRiskyUser.Read.All, IdentityRiskEvent.Read.All
<#
    RiskyUsers.ps1
    Who Entra ID Protection has flagged, how badly, and the detections behind it.

    Author   : Glenn Armbruster
    Created  : 2026-09-17
#>
# TOOLKIT-HELP-RUNS: Reads Identity Protection risky users and risk detections. Changes nothing.
# TOOLKIT-HELP-BEGIN
# The Identity Protection view in one grid: every user currently at risk, their
# risk level (high / medium / low), state (at risk, confirmed compromised,
# remediated, dismissed) and when the risk last changed. Pick a user and the
# lower grid shows the detections behind the score -- leaked credentials,
# anonymous IP, unfamiliar sign-in properties, impossible travel -- with the IP
# and location each one was seen from.
#
# Enter a user at the top to look at one account; leave it blank for the
# whole tenant. "Include remediated and dismissed" widens the top grid to
# users whose risk has already been dealt with.
#
# READ-ONLY. Dismissing risk or confirming a compromise needs
# IdentityRiskyUser.ReadWrite.All, which the tenant's Graph app does not
# hold -- do those in the Entra portal. If that permission is ever added, the
# two actions are a small change here.
#
# NEEDS: Entra ID P2 (Identity Protection), the Graph scopes
# IdentityRiskyUser.Read.All and IdentityRiskEvent.Read.All, and one of
# Global Reader, Security Reader, Security Operator or Security Administrator
# on YOUR account.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

$script:Scopes = @('IdentityRiskyUser.Read.All', 'IdentityRiskEvent.Read.All')
$ReportDir     = Get-ReportPath
$ReportPrefix  = 'RiskyUsers'

#region Data ------------------------------------------------------------------

function Get-RiskyUserList {
    <#
        Learn 2026-09-17 (List riskyUsers): GET /identityProtection/riskyUsers,
        $filter and $select, $top max 500. Properties riskLevel, riskState,
        riskDetail, riskLastUpdatedDateTime, userDisplayName,
        userPrincipalName, isDeleted, isProcessing.

        Filtered server-side to the states that mean something is still
        open, unless the caller asks for everything. 'none' is every user
        Identity Protection has ever looked at, which is the whole directory.
    #>
    param([switch]$IncludeClosed, [string]$UserId)
    $parts = @()
    if ($UserId) { $parts += "id eq '$UserId'" }
    if (-not $IncludeClosed) { $parts += "(riskState eq 'atRisk' or riskState eq 'confirmedCompromised')" }
    else { $parts += "(riskState eq 'atRisk' or riskState eq 'confirmedCompromised' or riskState eq 'remediated' or riskState eq 'dismissed')" }
    $f = $parts -join ' and '
    $uri = "identityProtection/riskyUsers?`$filter=" + [uri]::EscapeDataString($f) + "&`$top=500"
    return @(Invoke-ToolkitGraph -Uri $uri -All -MaxRows 2000)
}

function Get-RiskDetectionList {
    <#
        Learn 2026-09-17 (List riskDetections): GET /identityProtection/
        riskDetections, $filter, $top max 500. Properties riskEventType,
        riskLevel, riskState, riskDetail, detectedDateTime, activityDateTime,
        ipAddress, location, activity, source, userId, userPrincipalName,
        additionalInfo.

        Filtered on userId. Learn does not annotate which properties the
        filter accepts, so a rejected filter falls back to a time-bounded
        pull and a client-side match -- slower, same answer.
    #>
    param([Parameter(Mandatory)][string]$UserId, [int]$Days = 90)
    try {
        $uri = "identityProtection/riskDetections?`$filter=" + [uri]::EscapeDataString("userId eq '$UserId'") + "&`$top=500"
        return @(Invoke-ToolkitGraph -Uri $uri -All -MaxRows 2000)
    } catch {
        $since = (Get-Date).ToUniversalTime().AddDays(-1 * $Days).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $uri = "identityProtection/riskDetections?`$filter=" + [uri]::EscapeDataString("detectedDateTime ge $since") + "&`$top=500"
        $all = @(Invoke-ToolkitGraph -Uri $uri -All -MaxRows 5000)
        return @($all | Where-Object { "$($_.userId)" -eq $UserId })
    }
}

function Get-RiskLocationText {
    param($r)
    $l = $r.location
    if (-not $l) { return '' }
    return (@("$($l.city)", "$($l.state)", "$($l.countryOrRegion)") | Where-Object { $_ }) -join ', '
}

function ConvertTo-RiskLevelRank {
    # For sorting: high first.
    param([string]$Level)
    switch ("$Level".ToLowerInvariant()) { 'high' { 0 } 'medium' { 1 } 'low' { 2 } default { 3 } }
}
#endregion

#region UI --------------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Risky Users'
$form.ClientSize    = New-Object System.Drawing.Size(1180, 720)
$form.StartPosition = 'Manual'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$y = 12
$lblUser          = New-Object System.Windows.Forms.Label
$lblUser.Text     = 'User (blank = whole tenant):'
$lblUser.Location = New-Object System.Drawing.Point(12, ($y + 3))
$lblUser.Size     = New-Object System.Drawing.Size(170, 18)
$form.Controls.Add($lblUser)

$txtUser          = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(184, $y)
$txtUser.Size     = New-Object System.Drawing.Size(300, 24)
$txtUser.Font     = New-Object System.Drawing.Font('Consolas', 9)
$txtUser.Text     = "$(Get-TargetSmtp -Text $SmtpAddress)"
$form.Controls.Add($txtUser)

$chkClosed          = New-Object System.Windows.Forms.CheckBox
$chkClosed.Text     = 'Include remediated and dismissed'
$chkClosed.Location = New-Object System.Drawing.Point(496, ($y + 2))
$chkClosed.Size     = New-Object System.Drawing.Size(230, 20)
$form.Controls.Add($chkClosed)

$btnLoad          = New-Object System.Windows.Forms.Button
$btnLoad.Text     = 'Load'
$btnLoad.Location = New-Object System.Drawing.Point(740, ($y - 2))
$btnLoad.Size     = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnLoad)

$btnExport          = New-Object System.Windows.Forms.Button
$btnExport.Text     = 'Export CSV'
$btnExport.Location = New-Object System.Drawing.Point(848, ($y - 2))
$btnExport.Size     = New-Object System.Drawing.Size(100, 28)
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
$lblStatus.Text      = 'Press Load.'
$form.Controls.Add($lblStatus)
$y += 20

$pb = Add-ReportProgressBar -Form $form -X 12 -Y $y -Width 1156 -Height 12
$y += 18

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix $ReportPrefix `
            -X 12 -Y $y -Width 1156 -StatusLabel $lblStatus
$y += 34

$lblTop          = New-Object System.Windows.Forms.Label
$lblTop.Text     = 'Risky users (pick one for its detections):'
$lblTop.Location = New-Object System.Drawing.Point(12, $y)
$lblTop.Size     = New-Object System.Drawing.Size(600, 18)
$form.Controls.Add($lblTop)
$y += 20

function New-RiskGrid {
    param([object[]]$Columns, [int]$Height, [int]$Top)
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Location            = New-Object System.Drawing.Point(12, $Top)
    $g.Size                = New-Object System.Drawing.Size(1156, $Height)
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

$gridUsers = New-RiskGrid -Top $y -Height 250 -Columns @(
    @('User', 320), @('Name', 220), @('Level', 80), @('State', 170), @('Detail', 220), @('Updated', 118))
$form.Controls.Add($gridUsers)
$y += 258

$lblBottom          = New-Object System.Windows.Forms.Label
$lblBottom.Text     = 'Risk detections for the selected user:'
$lblBottom.Location = New-Object System.Drawing.Point(12, $y)
$lblBottom.Size     = New-Object System.Drawing.Size(600, 18)
$form.Controls.Add($lblBottom)
$y += 20

$gridDet = New-RiskGrid -Top $y -Height (720 - $y - 12) -Columns @(
    @('Detected', 118), @('Type', 230), @('Level', 80), @('State', 150), @('Detail', 190), @('Activity', 90), @('IP', 120), @('Location', 170), @('Source', 90))
$form.Controls.Add($gridDet)

$form | Add-Member -NotePropertyName 'Busy' -NotePropertyValue $false -Force
$script:Users = @()

$load = {
    $gridUsers.Rows.Clear(); $gridDet.Rows.Clear(); $script:Users = @(); $btnExport.Enabled = $false
    $p = New-ProgressReporter -Label $lblStatus -Bar $pb -Activity 'Risky users'
    $uid = ''
    $t = "$($txtUser.Text)".Trim()
    if ($t) {
        & $p "Resolving $t..." -1
        $u = Resolve-GraphUser -Text $t
        if (-not $u) { & $p "No user matches '$t'." 100; return }
        $uid = "$($u.id)"
    }
    & $p 'Reading risky users...' -1
    $list = @(Get-RiskyUserList -IncludeClosed:$chkClosed.Checked -UserId $uid)
    $list = @($list | Sort-Object @{ Expression = { ConvertTo-RiskLevelRank $_.riskLevel } }, @{ Expression = { "$($_.riskLastUpdatedDateTime)" }; Descending = $true })
    $script:Users = $list
    $gridUsers.SuspendLayout()
    try {
        foreach ($r in $list) {
            $i = $gridUsers.Rows.Add("$($r.userPrincipalName)", "$($r.userDisplayName)", "$($r.riskLevel)", "$($r.riskState)", "$($r.riskDetail)",
                    (ConvertTo-LocalGraphTime $r.riskLastUpdatedDateTime))
            $gridUsers.Rows[$i].Tag = $r
            $clr = switch ("$($r.riskLevel)".ToLowerInvariant()) {
                'high'   { [System.Drawing.Color]::FromArgb(180, 0, 0) }
                'medium' { [System.Drawing.Color]::FromArgb(190, 110, 0) }
                default  { $null }
            }
            if ($clr) { $gridUsers.Rows[$i].DefaultCellStyle.ForeColor = $clr }
        }
    } finally { $gridUsers.ResumeLayout() }
    if ($t -and $list.Count -eq 0) {
        & $p "$t is not flagged$(if (-not $chkClosed.Checked) { ' (tick Include remediated and dismissed to see closed risk)' })." 100
    } else {
        & $p "$($list.Count) risky user(s)." 100
    }
    $btnExport.Enabled = ($list.Count -gt 0)
    # SELECT ROW 0 AND LOAD ITS DETECTIONS EXPLICITLY.
    #
    # The grid selects its first row by itself as rows are added, which raises
    # SelectionChanged -- but that handler is guarded on $form.Busy, and this
    # runs INSIDE the busy window. So the guard swallowed the only selection
    # event of the first load and the lower grid stayed empty until the
    # operator clicked a second row.
    if ($gridUsers.Rows.Count -gt 0) {
        $gridUsers.Rows[0].Selected = $true
        & $loadDetections
    }
}

$loadDetections = {
    $gridDet.Rows.Clear()
    if ($gridUsers.SelectedRows.Count -eq 0) { return }
    $r = $gridUsers.SelectedRows[0].Tag
    if (-not $r) { return }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $lblBottom.Text = "Risk detections for $($r.userPrincipalName): reading..."
        $dets = @(Get-RiskDetectionList -UserId "$($r.id)")
        $dets = @($dets | Sort-Object @{ Expression = { "$($_.detectedDateTime)" }; Descending = $true })
        $gridDet.SuspendLayout()
        try {
            foreach ($d in $dets) {
                $i = $gridDet.Rows.Add((ConvertTo-LocalGraphTime $d.detectedDateTime), "$($d.riskEventType)", "$($d.riskLevel)", "$($d.riskState)",
                        "$($d.riskDetail)", "$($d.activity)", "$($d.ipAddress)", (Get-RiskLocationText $d), "$($d.source)")
                $gridDet.Rows[$i].Tag = $d
            }
        } finally { $gridDet.ResumeLayout() }
        $lblBottom.Text = "Risk detections for $($r.userPrincipalName): $($dets.Count)"
    } catch {
        $lblBottom.Text = "Risk detections for $($r.userPrincipalName): failed -- $($_.Exception.Message)"
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

$btnLoad.Add_Click({
    if ($form.Busy) { return }
    $form.Busy = $true
    try { $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; & $load }
    catch { $lblStatus.Text = "Load failed: $($_.Exception.Message)"; Show-HandlerError -ErrorRecord $_ -Context 'Load risky users (Graph identityProtection)' }
    finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
})
$txtUser.Add_KeyDown({ if ($_.KeyCode -eq 'Return') { $btnLoad.PerformClick(); $_.SuppressKeyPress = $true } })
$gridUsers.Add_SelectionChanged({ if (-not $form.Busy) { & $loadDetections } })

$btnExport.Add_Click({
    try {
        if ($script:Users.Count -eq 0) { $lblStatus.Text = 'Nothing to export. Load first.'; return }
        $export = New-Object System.Collections.Generic.List[object]
        foreach ($r in $script:Users) {
            $export.Add([PSCustomObject]@{
                User        = "$($r.userPrincipalName)"
                Name        = "$($r.userDisplayName)"
                RiskLevel   = "$($r.riskLevel)"
                RiskState   = "$($r.riskState)"
                RiskDetail  = "$($r.riskDetail)"
                LastUpdated = (ConvertTo-LocalGraphTime $r.riskLastUpdatedDateTime)
                IsDeleted   = "$($r.isDeleted)"
                IsProcessing = "$($r.isProcessing)"
                Id          = "$($r.id)"
            })
        }
        if (-not (Test-Path -LiteralPath $ReportDir)) { New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir ("{0}_{1}.csv" -f $ReportPrefix, $stamp)
        $export | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
        $files.SelectPath($csv)
        $lblStatus.Text = "Exported $($export.Count) row(s) to $csv"
        Write-CompletedLog -Action 'Report' -Target $csv -Result 'Success' `
            -Detail "Risky users: $($export.Count) row(s)$(if ($chkClosed.Checked) { ' incl. closed' })." -Environment 'Cloud'
    } catch {
        $lblStatus.Text = 'Export failed.'
        Show-HandlerError -ErrorRecord $_ -Context 'Export risky users'
    }
})

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Risky Users' | Out-Null

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
