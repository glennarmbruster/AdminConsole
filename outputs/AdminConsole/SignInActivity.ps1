#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-GRAPH-SCOPES: AuditLog.Read.All, User.Read.All
<#
    SignInActivity.ps1
    One user's recent sign-ins from the Entra sign-in log: when, from where,
    which app, and whether it worked.

    Author   : Glenn Armbruster
    Created  : 2026-09-17
#>
# TOOLKIT-HELP-RUNS: Reads the Entra sign-in log for one user. Changes nothing.
# TOOLKIT-HELP-BEGIN
# The Entra ID sign-in log for one user, most recent first: time, application,
# client, IP address, location, device, result, Conditional Access outcome and
# sign-in risk. Pick a row for the full detail, including the failure reason
# and every Conditional Access policy that was evaluated.
#
# This is the first thing to open for "I can't sign in", "was this account
# used from somewhere it shouldn't be", and "is MFA actually being applied".
#
# INTERACTIVE SIGN-INS ONLY. The v1.0 Graph API returns the interactive log --
# a person typing credentials or answering MFA. Token refreshes and
# non-interactive service sign-ins are a separate log that v1.0 does not expose;
# the summary line at the top shows the last non-interactive time, which is all
# that is available for those.
#
# 7 DAYS BY DEFAULT, up to 30. Entra keeps 30 days of sign-in log; a longer
# window is not available. Results stop at 1,000 rows and the status line says
# so -- narrow the window if that happens.
#
# NEEDS: an Entra ID P1 or P2 licence on the tenant (the sign-in log API
# requires it), the Graph scopes AuditLog.Read.All and User.Read.All, and one
# of these roles on YOUR account: Global Reader, Reports Reader, Security
# Reader, Security Operator or Security Administrator. Without the role every
# call fails with 403 even though the scopes are granted.
#
# Nothing here changes anything.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

$script:Scopes       = @('AuditLog.Read.All', 'User.Read.All')
$ReportDir    = Get-ReportPath
$ReportPrefix = 'SignInActivity'

#region Data ------------------------------------------------------------------

function Get-UserSignInSummary {
    <#
        Learn 2026-09-17 (user resource, signInActivity): must be $selected
        explicitly; properties lastSignInDateTime, lastNonInteractiveSignInDateTime,
        lastSuccessfulSignInDateTime. Needs AuditLog.Read.All and Entra P1/P2.
        Best-effort: a tenant without the licence gets the grid and no summary,
        which is worth more than no tool.
    #>
    param([Parameter(Mandatory)][string]$UserId)
    try {
        return (Invoke-ToolkitGraph -Uri "users/$UserId`?`$select=id,signInActivity")
    } catch { return $null }
}

function Get-UserSignIn {
    <#
        Learn 2026-09-17 (List signIns): GET /auditLogs/signIns, $filter on
        createdDateTime and userId, $top up to 1000, newest first. The time
        filter is what Microsoft recommend to keep the call from timing out.
    #>
    param(
        [Parameter(Mandatory)][string]$UserId,
        [Parameter(Mandatory)][int]$Days,
        [int]$MaxRows = 1000
    )
    $since = (Get-Date).ToUniversalTime().AddDays(-1 * $Days).ToString('yyyy-MM-ddTHH:mm:ssZ')
    $f = "createdDateTime ge $since and userId eq '$UserId'"
    $uri = "auditLogs/signIns?`$filter=" + [uri]::EscapeDataString($f) + "&`$top=500"
    return @(Invoke-ToolkitGraph -Uri $uri -All -MaxRows $MaxRows)
}

function Get-SignInResultText {
    param($SignIn)
    $code = 0
    try { $code = [int]$SignIn.status.errorCode } catch { $code = 0 }
    if ($code -eq 0) { return 'Success' }
    $why = "$($SignIn.status.failureReason)".Trim()
    if ($why) { return "FAILED $code -- $why" }
    return "FAILED $code"
}

function Get-SignInLocationText {
    param($SignIn)
    $l = $SignIn.location
    if (-not $l) { return '' }
    return (@("$($l.city)", "$($l.state)", "$($l.countryOrRegion)") | Where-Object { $_ }) -join ', '
}

function Get-SignInDetailText {
    <#
        Everything on the row, laid out for the detail box. The Conditional
        Access list is the part people actually come here for: "MFA wasn't
        asked for" is answered by which policy applied and what it enforced.
    #>
    param($s)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Time (local)     : $(ConvertTo-LocalGraphTime $s.createdDateTime 'yyyy-MM-dd HH:mm:ss')")
    [void]$sb.AppendLine("User             : $($s.userDisplayName)  <$($s.userPrincipalName)>")
    [void]$sb.AppendLine("Application      : $($s.appDisplayName)  [$($s.appId)]")
    [void]$sb.AppendLine("Resource         : $($s.resourceDisplayName)")
    [void]$sb.AppendLine("Client           : $($s.clientAppUsed)")
    [void]$sb.AppendLine("IP address       : $($s.ipAddress)")
    [void]$sb.AppendLine("Location         : $(Get-SignInLocationText $s)")
    if ($s.deviceDetail) {
        $d = $s.deviceDetail
        [void]$sb.AppendLine("Device           : $($d.displayName)  OS: $($d.operatingSystem)  Browser: $($d.browser)")
        [void]$sb.AppendLine("Device state     : compliant=$($d.isCompliant)  managed=$($d.isManaged)  trust=$($d.trustType)  id=$($d.deviceId)")
    }
    [void]$sb.AppendLine("Result           : $(Get-SignInResultText $s)")
    if ("$($s.status.additionalDetails)".Trim()) {
        [void]$sb.AppendLine("Detail           : $($s.status.additionalDetails)")
    }
    [void]$sb.AppendLine("Interactive      : $($s.isInteractive)")
    [void]$sb.AppendLine("Cond. Access     : $($s.conditionalAccessStatus)")
    [void]$sb.AppendLine("Risk (sign-in)   : $($s.riskLevelDuringSignIn)   aggregated: $($s.riskLevelAggregated)   state: $($s.riskState)   detail: $($s.riskDetail)")
    [void]$sb.AppendLine("Correlation id   : $($s.correlationId)")
    [void]$sb.AppendLine("Request id       : $($s.id)")
    $pols = @()
    if ($s.PSObject.Properties['appliedConditionalAccessPolicies']) { $pols = @($s.appliedConditionalAccessPolicies) }
    if ($pols.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Conditional Access policies evaluated:')
        foreach ($p in $pols) {
            $grant = @($p.enforcedGrantControls) -join ','
            $sess  = @($p.enforcedSessionControls) -join ','
            [void]$sb.AppendLine(("  {0,-18} {1}{2}{3}" -f "$($p.result)", "$($p.displayName)",
                $(if ($grant) { "   grant: $grant" } else { '' }), $(if ($sess) { "   session: $sess" } else { '' })))
        }
    } else {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Conditional Access policies: none returned (either none evaluated, or your role cannot read CA data -- Security Reader or Global Reader can).')
    }
    return $sb.ToString()
}
#endregion

#region UI --------------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Sign-in Activity'
$form.ClientSize    = New-Object System.Drawing.Size(1180, 720)
$form.StartPosition = 'Manual'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$y = 12
$lblUser          = New-Object System.Windows.Forms.Label
$lblUser.Text     = 'User:'
$lblUser.Location = New-Object System.Drawing.Point(12, ($y + 3))
$lblUser.Size     = New-Object System.Drawing.Size(40, 18)
$form.Controls.Add($lblUser)

$txtUser          = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(54, $y)
$txtUser.Size     = New-Object System.Drawing.Size(330, 24)
$txtUser.Font     = New-Object System.Drawing.Font('Consolas', 9)
$txtUser.Text     = "$(Get-TargetSmtp -Text $SmtpAddress)"
$form.Controls.Add($txtUser)

$lblDays          = New-Object System.Windows.Forms.Label
$lblDays.Text     = 'Last:'
$lblDays.Location = New-Object System.Drawing.Point(396, ($y + 3))
$lblDays.Size     = New-Object System.Drawing.Size(36, 18)
$form.Controls.Add($lblDays)

$cmbDays               = New-Object System.Windows.Forms.ComboBox
$cmbDays.DropDownStyle = 'DropDownList'
$cmbDays.Location      = New-Object System.Drawing.Point(434, $y)
$cmbDays.Size          = New-Object System.Drawing.Size(90, 24)
$script:DayChoices     = @(1, 3, 7, 14, 30)
foreach ($d in $script:DayChoices) { [void]$cmbDays.Items.Add("$d day$(if ($d -gt 1) { 's' })") }
$cmbDays.SelectedIndex = 2
$form.Controls.Add($cmbDays)

$chkFail          = New-Object System.Windows.Forms.CheckBox
$chkFail.Text     = 'Failures only'
$chkFail.Location = New-Object System.Drawing.Point(536, ($y + 2))
$chkFail.Size     = New-Object System.Drawing.Size(110, 20)
$form.Controls.Add($chkFail)

$btnLoad          = New-Object System.Windows.Forms.Button
$btnLoad.Text     = 'Load'
$btnLoad.Location = New-Object System.Drawing.Point(654, ($y - 2))
$btnLoad.Size     = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnLoad)

$btnExport          = New-Object System.Windows.Forms.Button
$btnExport.Text     = 'Export CSV'
$btnExport.Location = New-Object System.Drawing.Point(762, ($y - 2))
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

$lblWho           = New-Object System.Windows.Forms.Label
$lblWho.Tag       = 'NoTheme'
$lblWho.Location  = New-Object System.Drawing.Point(12, $y)
$lblWho.Size      = New-Object System.Drawing.Size(1156, 18)
$lblWho.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblWho.ForeColor = [System.Drawing.Color]::DimGray
$lblWho.Text      = ''
$form.Controls.Add($lblWho)
$y += 20

$lblSummary           = New-Object System.Windows.Forms.Label
$lblSummary.Location  = New-Object System.Drawing.Point(12, $y)
$lblSummary.Size      = New-Object System.Drawing.Size(1156, 18)
$lblSummary.ForeColor = [System.Drawing.Color]::DimGray
$lblSummary.Text      = ''
$form.Controls.Add($lblSummary)
$y += 22

$lblStatus           = New-Object System.Windows.Forms.Label
$lblStatus.Location  = New-Object System.Drawing.Point(12, $y)
$lblStatus.Size      = New-Object System.Drawing.Size(1156, 18)
$lblStatus.ForeColor = [System.Drawing.Color]::DimGray
$lblStatus.Text      = 'Enter a user and press Load.'
$form.Controls.Add($lblStatus)
$y += 20

$pb = Add-ReportProgressBar -Form $form -X 12 -Y $y -Width 1156 -Height 12
$y += 18

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix $ReportPrefix `
            -X 12 -Y $y -Width 1156 -StatusLabel $lblStatus
$y += 34

$grid                     = New-Object System.Windows.Forms.DataGridView
$grid.Location            = New-Object System.Drawing.Point(12, $y)
$grid.Size                = New-Object System.Drawing.Size(1156, 330)
$grid.ReadOnly            = $true
$grid.AllowUserToAddRows  = $false
$grid.SelectionMode       = 'FullRowSelect'
$grid.MultiSelect         = $false
$grid.RowHeadersVisible   = $false
$grid.AutoSizeColumnsMode = 'None'
$grid.ScrollBars          = 'Both'
$grid.Font                = New-Object System.Drawing.Font('Consolas', 9)
foreach ($c in @(@('Time', 118), @('App', 200), @('Client', 110), @('IP', 120), @('Location', 170), @('Device', 190), @('Result', 240), @('CA', 100), @('Risk', 70))) {
    $idx = $grid.Columns.Add($c[0], $c[0])
    $grid.Columns[$idx].Width    = $c[1]
    $grid.Columns[$idx].SortMode = 'Automatic'
}
$form.Controls.Add($grid)
$y += 338

$txtDetail            = New-Object System.Windows.Forms.TextBox
$txtDetail.Multiline  = $true
$txtDetail.ReadOnly   = $true
$txtDetail.ScrollBars = 'Both'
$txtDetail.WordWrap   = $false
$txtDetail.Location   = New-Object System.Drawing.Point(12, $y)
$txtDetail.Size       = New-Object System.Drawing.Size(1156, (720 - $y - 12))
$txtDetail.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtDetail.BackColor  = [System.Drawing.Color]::WhiteSmoke
$txtDetail.ForeColor  = [System.Drawing.Color]::Black
$txtDetail.Tag        = 'NoTheme'
$txtDetail.Text       = 'Select a sign-in for the full detail, including Conditional Access policies.'
$form.Controls.Add($txtDetail)

$form | Add-Member -NotePropertyName 'Busy' -NotePropertyValue $false -Force
$script:Rows = @()
$script:User = $null

$load = {
    $grid.Rows.Clear(); $script:Rows = @(); $btnExport.Enabled = $false
    $txtDetail.Text = ''
    $t = "$($txtUser.Text)".Trim()
    if (-not $t) { $lblStatus.Text = 'Enter a user (UPN or email address).'; return }
    $days = $script:DayChoices[$cmbDays.SelectedIndex]

    $p = New-ProgressReporter -Label $lblStatus -Bar $pb -Activity 'Sign-in activity'
    & $p "Resolving $t..." -1
    $u = Resolve-GraphUser -Text $t
    if (-not $u) { & $p "No user matches '$t'." 100; $lblWho.Text = ''; return }
    $script:User = $u
    $lblWho.Text = "$($u.displayName)  <$($u.userPrincipalName)>   account $(if ($u.accountEnabled) { 'ENABLED' } else { 'DISABLED' })   type: $($u.userType)"
    $lblWho.ForeColor = $(if ($u.accountEnabled) { [System.Drawing.Color]::FromArgb(0, 110, 60) } else { [System.Drawing.Color]::FromArgb(200, 0, 0) })

    & $p 'Reading last sign-in summary...' -1
    $sum = Get-UserSignInSummary -UserId $u.id
    if ($sum -and $sum.signInActivity) {
        $a = $sum.signInActivity
        $lblSummary.Text = "Last interactive: $(ConvertTo-LocalGraphTime $a.lastSignInDateTime)    last non-interactive: $(ConvertTo-LocalGraphTime $a.lastNonInteractiveSignInDateTime)    last successful: $(ConvertTo-LocalGraphTime $a.lastSuccessfulSignInDateTime)"
    } else {
        $lblSummary.Text = 'Last sign-in summary not available (needs Entra ID P1/P2 and AuditLog.Read.All).'
    }

    & $p "Reading sign-ins for the last $days day(s)..." -1
    $all = @(Get-UserSignIn -UserId $u.id -Days $days)
    $capped = ($all.Count -ge 1000)
    if ($chkFail.Checked) {
        $all = @($all | Where-Object { (Get-SignInResultText $_) -ne 'Success' })
    }
    $script:Rows = $all
    $grid.SuspendLayout()
    try {
        foreach ($s in $all) {
            $dev = @("$($s.deviceDetail.operatingSystem)", "$($s.deviceDetail.browser)") | Where-Object { $_ }
            $i = $grid.Rows.Add(
                (ConvertTo-LocalGraphTime $s.createdDateTime 'yyyy-MM-dd HH:mm:ss'),
                "$($s.appDisplayName)", "$($s.clientAppUsed)", "$($s.ipAddress)",
                (Get-SignInLocationText $s), ($dev -join ' / '),
                (Get-SignInResultText $s), "$($s.conditionalAccessStatus)", "$($s.riskLevelDuringSignIn)")
            $grid.Rows[$i].Tag = $s
            if ((Get-SignInResultText $s) -ne 'Success') {
                $grid.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
            }
        }
    } finally { $grid.ResumeLayout() }
    $msg = "$($all.Count) sign-in(s) in the last $days day(s)$(if ($chkFail.Checked) { ' (failures only)' })."
    if ($capped) { $msg += '  STOPPED AT 1,000 ROWS -- narrow the window to see everything.' }
    & $p $msg 100
    $btnExport.Enabled = ($all.Count -gt 0)
}

$btnLoad.Add_Click({
    if ($form.Busy) { return }
    $form.Busy = $true
    try { $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor; & $load }
    catch { $lblStatus.Text = "Load failed: $($_.Exception.Message)"; Show-HandlerError -ErrorRecord $_ -Context 'Load sign-ins (Graph auditLogs/signIns)' }
    finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $form.Busy = $false }
})
$txtUser.Add_KeyDown({ if ($_.KeyCode -eq 'Return') { $btnLoad.PerformClick(); $_.SuppressKeyPress = $true } })

$grid.Add_SelectionChanged({
    $r = $null
    if ($grid.SelectedRows.Count -gt 0) { $r = $grid.SelectedRows[0].Tag }
    if ($r) { $txtDetail.Text = (Get-SignInDetailText $r) } else { $txtDetail.Text = '' }
})

$btnExport.Add_Click({
    try {
        if ($script:Rows.Count -eq 0) { $lblStatus.Text = 'Nothing to export. Load first.'; return }
        $export = New-Object System.Collections.Generic.List[object]
        foreach ($s in $script:Rows) {
            $pols = @()
            if ($s.PSObject.Properties['appliedConditionalAccessPolicies']) {
                $pols = @($s.appliedConditionalAccessPolicies | ForEach-Object { "$($_.displayName)=$($_.result)" })
            }
            $export.Add([PSCustomObject]@{
                Time            = (ConvertTo-LocalGraphTime $s.createdDateTime 'yyyy-MM-dd HH:mm:ss')
                User            = "$($s.userPrincipalName)"
                App             = "$($s.appDisplayName)"
                Resource        = "$($s.resourceDisplayName)"
                Client          = "$($s.clientAppUsed)"
                IP              = "$($s.ipAddress)"
                Location        = (Get-SignInLocationText $s)
                OS              = "$($s.deviceDetail.operatingSystem)"
                Browser         = "$($s.deviceDetail.browser)"
                DeviceName      = "$($s.deviceDetail.displayName)"
                Compliant       = "$($s.deviceDetail.isCompliant)"
                Managed         = "$($s.deviceDetail.isManaged)"
                Result          = (Get-SignInResultText $s)
                AdditionalDetail = "$($s.status.additionalDetails)"
                ConditionalAccess = "$($s.conditionalAccessStatus)"
                Policies        = ($pols -join '; ')
                RiskDuringSignIn = "$($s.riskLevelDuringSignIn)"
                RiskState       = "$($s.riskState)"
                CorrelationId   = "$($s.correlationId)"
            })
        }
        if (-not (Test-Path -LiteralPath $ReportDir)) { New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null }
        $who   = ("$($script:User.userPrincipalName)" -replace '[^A-Za-z0-9]', '_')
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir ("{0}_{1}_{2}.csv" -f $ReportPrefix, $who, $stamp)
        $export | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
        $files.SelectPath($csv)
        $lblStatus.Text = "Exported $($export.Count) row(s) to $csv"
        Write-CompletedLog -Action 'Report' -Target $csv -Result 'Success' `
            -Detail "Sign-in activity for $($script:User.userPrincipalName): $($export.Count) row(s)." -Environment 'Cloud'
    } catch {
        $lblStatus.Text = 'Export failed.'
        Show-HandlerError -ErrorRecord $_ -Context 'Export sign-ins'
    }
})

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Sign-in Activity' | Out-Null

$form.Add_Shown({
    # THE GATE. Scopes are checked before anything is asked of Graph, and the
    # answer goes on the status line rather than into an error dialog over a
    # window that has not finished painting.
    $gate = Assert-ToolkitGraphScope -Scopes $script:Scopes
    if (-not $gate.Ok) {
        $lblStatus.Text = $gate.Message
        $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(200, 0, 0)
        $btnLoad.Enabled = $false
        return
    }
    if ("$($txtUser.Text)".Trim()) { $btnLoad.PerformClick() }
})

[void]$form.ShowDialog()
$form.Dispose()
#endregion
