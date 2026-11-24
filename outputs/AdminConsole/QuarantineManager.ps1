#Requires -Version 5.1
<#
    QuarantineManager.ps1
    Find, inspect, release and delete quarantined messages in Exchange Online
    without opening the Defender portal.

    Author   : Glenn Armbruster
    Revised  : 2026-09-12
#>
# TOOLKIT-HELP-RUNS: Exchange Online quarantine. Reads only, until you press Release or Delete.
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# Searches the Exchange Online quarantine by recipient, sender, subject, age
# and quarantine type, and lists what it finds in a grid.
#
# From the grid you can:
#   Preview headers          read the message header without releasing it
#   Release to recipient     deliver to every ORIGINAL recipient (-ReleaseToAll)
#   Release + Report FP      the same, and tell Microsoft it was a false positive
#                            (the service only accepts the report for Spam)
#   Delete                   remove from quarantine (soft delete, not -HardDelete)
#   Export CSV               write the current grid to the reports folder
#
# Every release and delete shows the exact list first and asks. Each row is
# acted on separately, so one failure does not stop the rest, and every
# success and failure is written to the audit log.
#
# The search reads at most 30 days back (the service keeps no more) and stops
# at 1000 results; the summary says so when that cap is hit. It never
# releases to an address that was not an original recipient, and never
# hard-deletes.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# FIRST THING AFTER Common LOADS. See Test-ConsoleLaunch for why a guard that
# runs after the form is built is a guard that already lost.
if (-not (Test-ConsoleLaunch)) { return }

#region Configuration --------------------------------------------------------
$ReportDir = Get-ReportPath

# The service pages at 100 by default and the cap keeps a wide-open search
# from running for minutes; 1000 is also as far as -Page will go.
$PageSize   = 100
$MaxResults = 1000
$MaxDays    = 30

# Dropdown label -> value the QuarantineTypes parameter accepts. The label
# says what the portal says; the value is what the cmdlet documents.
$TypeMap = [ordered]@{
    'All'                 = ''
    'Spam'                = 'Spam'
    'HighConfidencePhish' = 'HighConfPhish'
    'Phish'               = 'Phish'
    'Malware'             = 'Malware'
    'Bulk'                = 'Bulk'
    'TransportRule'       = 'TransportRule'
}
#endregion

#region Data -----------------------------------------------------------------
function Assert-QuarantineCmdlet {
    param([string]$Name)
    if (-not (Get-Command -Name $Name -ErrorAction SilentlyContinue)) {
        throw "$Name is not available. Connect to Exchange Online first."
    }
}

function Format-Stamp {
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return '' }
    try { return ([datetime]$Value).ToString('yyyy-MM-dd HH:mm') } catch { return "$Value" }
}

function Get-QuarantineRows {
    param(
        [string]$Recipient,
        [string]$Sender,
        [string]$Subject,
        [int]$Days,
        [string]$Type,
        [scriptblock]$Progress
    )

    Assert-QuarantineCmdlet -Name 'Get-QuarantineMessage'

    $declared = Get-CmdletParameterNames -Name 'Get-QuarantineMessage'

    $params = @{ ErrorAction = 'Stop' }
    # Bulk-read rule: pass ResultSize only if the cmdlet declares it. This
    # one does not (it pages instead), so the check is what keeps the rule
    # honest rather than a parameter that gets sent.
    if ($declared -contains 'ResultSize') { $params['ResultSize'] = 'Unlimited' }

    if ("$Recipient".Trim()) { $params['RecipientAddress'] = "$Recipient".Trim() }
    if ("$Sender".Trim())    { $params['SenderAddress']    = "$Sender".Trim() }
    if ("$Subject".Trim())   { $params['Subject']          = "$Subject".Trim() }
    if ("$Type".Trim())      { $params['QuarantineTypes']  = "$Type".Trim() }

    if ($Days -lt 1) { $Days = 1 }
    if ($Days -gt $MaxDays) { $Days = $MaxDays }
    $params['StartReceivedDate'] = (Get-Date).AddDays(-$Days)
    $params['EndReceivedDate']   = Get-Date

    $canPage = ($declared -contains 'Page') -and ($declared -contains 'PageSize')

    $rows   = New-Object System.Collections.Generic.List[object]
    $raw    = New-Object System.Collections.Generic.List[object]
    $page   = 1
    $capped = $false

    & $Progress 'Searching quarantine...' 5

    while ($true) {
        $p = @{} + $params
        if ($canPage) {
            $p['Page']     = $page
            $p['PageSize'] = $PageSize
        }
        $batch = @(Get-QuarantineMessage @p)
        foreach ($m in $batch) { $raw.Add($m) }

        & $Progress "Page $page read -- $($raw.Count) message(s) so far..." ([int][math]::Min(90, 5 + 8 * $page))

        if (-not $canPage)              { break }
        if ($batch.Count -lt $PageSize) { break }
        if ($raw.Count -ge $MaxResults) { $capped = $true; break }
        $page++
    }

    foreach ($m in $raw) {
        if ($rows.Count -ge $MaxResults) { $capped = $true; break }

        $rcpt = @($m.RecipientAddress | Where-Object { $_ }) -join '; '
        $type = "$($m.Type)"
        if (-not $type -and $m.QuarantineTypes) { $type = (@($m.QuarantineTypes) -join ', ') }
        $released = 'No'
        if ($m.Released -eq $true) { $released = 'Yes' }

        $rows.Add([PSCustomObject][ordered]@{
            Received  = Format-Stamp $m.ReceivedTime
            Sender    = "$($m.SenderAddress)"
            Recipient = $rcpt
            Subject   = "$($m.Subject)"
            Type      = $type
            Direction = "$($m.Direction)"
            Released  = $released
            Expires   = Format-Stamp $m.Expires
            Identity  = "$($m.Identity)"
        })
    }

    return [PSCustomObject]@{
        Rows     = @($rows | Sort-Object Received -Descending)
        Pages    = $page
        Capped   = $capped
        CanPage  = $canPage
        Days     = $Days
    }
}

function Format-SearchSummary {
    param($Result, [hashtable]$Filters)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Quarantined messages found : $(@($Result.Rows).Count)")
    [void]$sb.AppendLine("Days searched              : $($Result.Days)")
    [void]$sb.AppendLine("Pages read                 : $($Result.Pages)")
    foreach ($k in @('Recipient', 'Sender', 'Subject', 'Type')) {
        $v = "$($Filters[$k])".Trim()
        if (-not $v) { $v = '(any)' }
        [void]$sb.AppendLine(("{0,-27}: {1}" -f $k, $v))
    }
    if (@($Result.Rows).Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('By type:')
        $Result.Rows | Group-Object Type | Sort-Object Count -Descending | ForEach-Object {
            [void]$sb.AppendLine(("  {0,-24}{1,6}" -f $_.Name, $_.Count))
        }
        $rel = @($Result.Rows | Where-Object { $_.Released -eq 'Yes' })
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("Already released           : $($rel.Count)")
    }
    if ($Result.Capped) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("*** CAPPED at $MaxResults results. Narrow the search (recipient, ***")
        [void]$sb.AppendLine('*** sender, fewer days) to see the rest. This list is INCOMPLETE. ***')
    }
    if (-not $Result.CanPage) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('*** This module build does not declare -Page/-PageSize, so only the ***')
        [void]$sb.AppendLine('*** first page the service returned is shown.                     ***')
    }
    return $sb.ToString()
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Quarantine Manager -- Exchange Online'
$form.Size            = New-Object System.Drawing.Size(1040, 700)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Search the Exchange Online quarantine, then select rows to preview, release or delete. Releases go to the original recipients only.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(1000, 18)
$form.Controls.Add($lblInfo)

# --- Filter row 1 ---
$lblRcpt          = New-Object System.Windows.Forms.Label
$lblRcpt.Text     = 'Recipient:'
$lblRcpt.Location = New-Object System.Drawing.Point(12, 43)
$lblRcpt.Size     = New-Object System.Drawing.Size(70, 18)
$form.Controls.Add($lblRcpt)

$txtRcpt          = New-Object System.Windows.Forms.TextBox
$txtRcpt.Location = New-Object System.Drawing.Point(86, 40)
$txtRcpt.Size     = New-Object System.Drawing.Size(260, 22)
$txtRcpt.Text     = "$SmtpAddress".Trim()
$form.Controls.Add($txtRcpt)

$lblSender          = New-Object System.Windows.Forms.Label
$lblSender.Text     = 'Sender:'
$lblSender.Location = New-Object System.Drawing.Point(370, 43)
$lblSender.Size     = New-Object System.Drawing.Size(60, 18)
$form.Controls.Add($lblSender)

$txtSender          = New-Object System.Windows.Forms.TextBox
$txtSender.Location = New-Object System.Drawing.Point(432, 40)
$txtSender.Size     = New-Object System.Drawing.Size(260, 22)
$form.Controls.Add($txtSender)

$lblDays          = New-Object System.Windows.Forms.Label
$lblDays.Text     = 'Days back:'
$lblDays.Location = New-Object System.Drawing.Point(716, 43)
$lblDays.Size     = New-Object System.Drawing.Size(70, 18)
$form.Controls.Add($lblDays)

$numDays          = New-Object System.Windows.Forms.NumericUpDown
$numDays.Location = New-Object System.Drawing.Point(790, 40)
$numDays.Size     = New-Object System.Drawing.Size(60, 22)
$numDays.Minimum  = 1
$numDays.Maximum  = $MaxDays
$numDays.Value    = 7
$form.Controls.Add($numDays)

$btnSearch          = New-Object System.Windows.Forms.Button
$btnSearch.Text     = 'Search'
$btnSearch.Location = New-Object System.Drawing.Point(902, 40)
$btnSearch.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnSearch)

# --- Filter row 2 ---
$lblSubject          = New-Object System.Windows.Forms.Label
$lblSubject.Text     = 'Subject:'
$lblSubject.Location = New-Object System.Drawing.Point(12, 73)
$lblSubject.Size     = New-Object System.Drawing.Size(70, 18)
$form.Controls.Add($lblSubject)

$txtSubject          = New-Object System.Windows.Forms.TextBox
$txtSubject.Location = New-Object System.Drawing.Point(86, 70)
$txtSubject.Size     = New-Object System.Drawing.Size(260, 22)
$form.Controls.Add($txtSubject)

$lblType          = New-Object System.Windows.Forms.Label
$lblType.Text     = 'Type:'
$lblType.Location = New-Object System.Drawing.Point(370, 73)
$lblType.Size     = New-Object System.Drawing.Size(60, 18)
$form.Controls.Add($lblType)

$cboType               = New-Object System.Windows.Forms.ComboBox
$cboType.Location      = New-Object System.Drawing.Point(432, 70)
$cboType.Size          = New-Object System.Drawing.Size(180, 22)
$cboType.DropDownStyle = 'DropDownList'
foreach ($k in $TypeMap.Keys) { [void]$cboType.Items.Add($k) }
$cboType.SelectedIndex = 0
$form.Controls.Add($cboType)

$lblHint           = New-Object System.Windows.Forms.Label
$lblHint.Text      = 'Subject is a contains match. Max 30 days; results stop at 1000.'
$lblHint.Location  = New-Object System.Drawing.Point(624, 73)
$lblHint.Size      = New-Object System.Drawing.Size(270, 18)
$lblHint.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblHint)

# --- Previous runs / status / progress ---
$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 138)
$lblStatus.Size     = New-Object System.Drawing.Size(1000, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 158 -Width 1000 -Height 12

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'QuarantineManager' `
                -X 12 -Y 104 -Width 1000 -StatusLabel $lblStatus

# --- Grid ---
$grid                        = New-Object System.Windows.Forms.DataGridView
$grid.Location               = New-Object System.Drawing.Point(12, 176)
$grid.Size                   = New-Object System.Drawing.Size(1000, 270)
$grid.ReadOnly               = $true
$grid.AllowUserToAddRows     = $false
$grid.AllowUserToDeleteRows  = $false
$grid.AllowUserToResizeRows  = $false
$grid.SelectionMode          = 'FullRowSelect'
$grid.MultiSelect            = $true
$grid.AutoSizeColumnsMode    = 'Fill'
$grid.RowHeadersVisible      = $false
[void]$grid.Columns.Add('Received',  'Received')
[void]$grid.Columns.Add('Sender',    'Sender')
[void]$grid.Columns.Add('Recipient', 'Recipient')
[void]$grid.Columns.Add('Subject',   'Subject')
[void]$grid.Columns.Add('Type',      'Type')
[void]$grid.Columns.Add('Direction', 'Direction')
[void]$grid.Columns.Add('Released',  'Released')
[void]$grid.Columns.Add('Expires',   'Expires')
[void]$grid.Columns.Add('Identity',  'Identity')
$grid.Columns['Received'].FillWeight  = 12
$grid.Columns['Sender'].FillWeight    = 20
$grid.Columns['Recipient'].FillWeight = 20
$grid.Columns['Subject'].FillWeight   = 26
$grid.Columns['Type'].FillWeight      = 9
$grid.Columns['Direction'].FillWeight = 7
$grid.Columns['Released'].FillWeight  = 6
$grid.Columns['Expires'].FillWeight   = 12
# Two GUIDs never fit a grid cell and mean nothing on sight. It travels in
# the CSV, the header dialog and every log row instead.
$grid.Columns['Identity'].Visible     = $false
$form.Controls.Add($grid)

$lblCount          = New-Object System.Windows.Forms.Label
$lblCount.Text     = 'No search run yet.'
$lblCount.Location = New-Object System.Drawing.Point(12, 452)
$lblCount.Size     = New-Object System.Drawing.Size(1000, 18)
$form.Controls.Add($lblCount)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Location   = New-Object System.Drawing.Point(12, 474)
$txtOut.Size       = New-Object System.Drawing.Size(1000, 130)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

# --- Buttons ---
$btnHeaders          = New-Object System.Windows.Forms.Button
$btnHeaders.Text     = 'Preview headers'
$btnHeaders.Location = New-Object System.Drawing.Point(12, 616)
$btnHeaders.Size     = New-Object System.Drawing.Size(140, 30)
$form.Controls.Add($btnHeaders)

$btnRelease          = New-Object System.Windows.Forms.Button
$btnRelease.Text     = 'Release to recipient'
$btnRelease.Location = New-Object System.Drawing.Point(164, 616)
$btnRelease.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnRelease)

$btnReleaseFP          = New-Object System.Windows.Forms.Button
$btnReleaseFP.Text     = 'Release + Report false positive'
$btnReleaseFP.Location = New-Object System.Drawing.Point(326, 616)
$btnReleaseFP.Size     = New-Object System.Drawing.Size(170, 30)
$form.Controls.Add($btnReleaseFP)

$btnDelete          = New-Object System.Windows.Forms.Button
$btnDelete.Text     = 'Delete'
$btnDelete.Location = New-Object System.Drawing.Point(508, 616)
$btnDelete.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnDelete)

$btnExport          = New-Object System.Windows.Forms.Button
$btnExport.Text     = 'Export CSV'
$btnExport.Location = New-Object System.Drawing.Point(630, 616)
$btnExport.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnExport)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Reports Folder'
$btnOpen.Location = New-Object System.Drawing.Point(752, 616)
$btnOpen.Size     = New-Object System.Drawing.Size(138, 30)
$form.Controls.Add($btnOpen)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(902, 616)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)
#endregion

#region Grid helpers ---------------------------------------------------------
function Update-CountLabel {
    $n = $grid.Rows.Count
    $s = $grid.SelectedRows.Count
    $lblCount.Text = "$n row(s) in grid, $s selected."
}

function Set-GridRows {
    param($Rows)
    $grid.SuspendLayout()
    try {
        $grid.Rows.Clear()
        foreach ($r in @($Rows)) {
            $i = $grid.Rows.Add(@($r.Received, $r.Sender, $r.Recipient, $r.Subject,
                                  $r.Type, $r.Direction, $r.Released, $r.Expires, $r.Identity))
            $grid.Rows[$i].Tag = $r
        }
        $grid.ClearSelection()
    } finally {
        $grid.ResumeLayout()
    }
    Update-CountLabel
}

function Get-SelectedRowObjects {
    # Grid order, not click order, so the preview reads the way the grid does.
    $picked = @()
    foreach ($gr in $grid.SelectedRows) { $picked += $gr }
    return @($picked | Sort-Object Index | ForEach-Object { $_.Tag } | Where-Object { $_ })
}

function Get-GridRowObjects {
    $all = @()
    foreach ($gr in $grid.Rows) { if ($gr.Tag) { $all += $gr.Tag } }
    return @($all)
}

function Set-RowStatus {
    param([string]$Identity, [string]$Released, [switch]$Remove)
    for ($i = $grid.Rows.Count - 1; $i -ge 0; $i--) {
        $gr = $grid.Rows[$i]
        if (-not $gr.Tag -or "$($gr.Tag.Identity)" -ne $Identity) { continue }
        if ($Remove) {
            $grid.Rows.RemoveAt($i)
        } else {
            $gr.Tag.Released = $Released
            $gr.Cells['Released'].Value = $Released
        }
    }
}

function Format-RowLine {
    param($Row)
    $subj = "$($Row.Subject)"
    if ($subj.Length -gt 60) { $subj = $subj.Substring(0, 57) + '...' }
    return "  $($Row.Received)  $($Row.Type)  $($Row.Sender) -> $($Row.Recipient)  ""$subj"""
}
#endregion

#region Write actions --------------------------------------------------------
function Invoke-QuarantineAction {
    <#
        One path for all three writes so the preview, the prompt, the per-row
        try/catch and the logging cannot drift apart between buttons.
        $Mode is 'Release', 'ReleaseFP' or 'Delete'.
    #>
    param([string]$Mode)

    $rows = Get-SelectedRowObjects
    if ($rows.Count -eq 0) {
        $txtOut.Text = 'Select one or more rows in the grid first.'
        return
    }

    switch ($Mode) {
        'Release' {
            $cmdlet = 'Release-QuarantineMessage'; $logAction = 'ReleaseQuarantine'
            $verb = 'Release'; $title = 'Release to original recipients'
            $what = 'RELEASE each message below to ALL of its original recipients.'
        }
        'ReleaseFP' {
            $cmdlet = 'Release-QuarantineMessage'; $logAction = 'ReleaseQuarantine'
            $verb = 'Release'; $title = 'Release and report false positive'
            $what = 'RELEASE each message below to ALL of its original recipients AND report it to Microsoft as a false positive.'
        }
        'Delete' {
            $cmdlet = 'Delete-QuarantineMessage'; $logAction = 'DeleteQuarantine'
            $verb = 'Delete'; $title = 'Delete from quarantine'
            $what = 'DELETE each message below from quarantine (soft delete; the service may still allow recovery).'
        }
        default { throw "Unknown mode '$Mode'." }
    }

    Assert-QuarantineCmdlet -Name $cmdlet

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("This will $what")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("$($rows.Count) message(s):")
    $shown = 0
    foreach ($r in $rows) {
        $shown++
        if ($shown -gt 15) { [void]$sb.AppendLine("  ... and $($rows.Count - 15) more."); break }
        [void]$sb.AppendLine((Format-RowLine -Row $r))
    }
    $already = @($rows | Where-Object { $_.Released -eq 'Yes' })
    if ($verb -eq 'Release' -and $already.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("$($already.Count) of these show as already released. The service skips recipients it")
        [void]$sb.AppendLine('already delivered to; -Force is passed for those rows so the request is accepted.')
    }
    if ($Mode -eq 'ReleaseFP') {
        $notSpam = @($rows | Where-Object { $_.Type -ne 'Spam' })
        if ($notSpam.Count -gt 0) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine("$($notSpam.Count) of these are not Spam. The false-positive report is only accepted for")
            [void]$sb.AppendLine('Spam; the service may reject those rows, and each rejection is listed afterwards.')
        }
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Continue?')

    $answer = Show-ToolDialog($sb.ToString(), $title, 'YesNo', 'Warning', 'Button2')
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
        $txtOut.Text = "Cancelled -- nothing was changed."
        return
    }

    $ok     = New-Object System.Collections.Generic.List[string]
    $failed = New-Object System.Collections.Generic.List[string]
    $n = 0
    $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity $title
    try {
        foreach ($r in $rows) {
            $n++
            & $progress "$verb $n of $($rows.Count)..." ([int](100.0 * $n / $rows.Count))
            $detail = "$($r.Sender) -> $($r.Recipient): $($r.Subject)"
            try {
                $p = @{ Identity = $r.Identity; ErrorAction = 'Stop'; Confirm = $false }
                if ($verb -eq 'Release') {
                    $p['ReleaseToAll'] = $true
                    if ($Mode -eq 'ReleaseFP')  { $p['ReportFalsePositive'] = $true }
                    if ($r.Released -eq 'Yes')  { $p['Force'] = $true }
                    Release-QuarantineMessage @p | Out-Null
                    Set-RowStatus -Identity $r.Identity -Released 'Yes'
                } else {
                    Delete-QuarantineMessage @p | Out-Null
                    Set-RowStatus -Identity $r.Identity -Remove
                }
                Write-CompletedLog -Action $logAction -Target $r.Identity -Result 'Success' -Detail $detail
                $ok.Add((Format-RowLine -Row $r))
            } catch {
                Write-FailedLog -Action $logAction -Target $r.Identity -Detail $detail -ErrorRecord $_
                $failed.Add((Format-RowLine -Row $r) + [Environment]::NewLine +
                            "      $($_.Exception.Message)")
            }
        }
    } finally {
        Complete-ProgressReporter -Activity $title -Bar $pbar
        Update-CountLabel
    }

    $out = New-Object System.Text.StringBuilder
    [void]$out.AppendLine("$title -- $($ok.Count) succeeded, $($failed.Count) failed.")
    if ($ok.Count -gt 0) {
        [void]$out.AppendLine('')
        [void]$out.AppendLine('Succeeded:')
        foreach ($l in $ok) { [void]$out.AppendLine($l) }
    }
    if ($failed.Count -gt 0) {
        [void]$out.AppendLine('')
        [void]$out.AppendLine('FAILED:')
        foreach ($l in $failed) { [void]$out.AppendLine($l) }
    }
    $txtOut.Text = $out.ToString()
    $lblStatus.Text = "$title done: $($ok.Count) ok, $($failed.Count) failed."
}
#endregion

#region Header preview -------------------------------------------------------
function Show-HeaderDialog {
    param($Row)

    Assert-QuarantineCmdlet -Name 'Get-QuarantineMessageHeader'

    $lblStatus.Text = 'Reading message header...'
    $h = Get-QuarantineMessageHeader -Identity $Row.Identity -ErrorAction Stop
    $text = ''
    if ($null -ne $h) {
        # The cmdlet returns an object whose Header member carries the raw
        # RFC 5322 text on current builds; fall back to the formatted object
        # on any build where the member is named differently.
        if ($h.PSObject.Properties['Header']) { $text = "$($h.Header)" }
        else                                  { $text = ($h | Out-String) }
    }
    if (-not "$text".Trim()) { $text = '(The service returned no header text for this message.)' }
    $lblStatus.Text = 'Idle.'

    $dlg                 = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Headers -- $($Row.Subject)"
    $dlg.Size            = New-Object System.Drawing.Size(860, 600)
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false
    $dlg.StartPosition   = 'CenterParent'
    $dlg.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $lblId          = New-Object System.Windows.Forms.Label
    $lblId.Text     = "From $($Row.Sender) to $($Row.Recipient) -- received $($Row.Received) -- Identity $($Row.Identity)"
    $lblId.Location = New-Object System.Drawing.Point(12, 12)
    $lblId.Size     = New-Object System.Drawing.Size(820, 18)
    $dlg.Controls.Add($lblId)

    $txtHdr            = New-Object System.Windows.Forms.TextBox
    $txtHdr.Multiline  = $true
    $txtHdr.ReadOnly   = $true
    $txtHdr.ScrollBars = 'Both'
    $txtHdr.WordWrap   = $false
    $txtHdr.Location   = New-Object System.Drawing.Point(12, 36)
    $txtHdr.Size       = New-Object System.Drawing.Size(820, 476)
    $txtHdr.Font       = New-Object System.Drawing.Font('Consolas', 9)
    $txtHdr.Text       = ("$text" -replace "(?<!`r)`n", "`r`n")
    $dlg.Controls.Add($txtHdr)

    $btnHdrClose          = New-Object System.Windows.Forms.Button
    $btnHdrClose.Text     = 'Close'
    $btnHdrClose.Location = New-Object System.Drawing.Point(722, 522)
    $btnHdrClose.Size     = New-Object System.Drawing.Size(110, 30)
    # DialogResult closes the modal form by itself. A Click handler here would
    # run outside this function's scope and not see $dlg.
    $btnHdrClose.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnHdrClose)
    $dlg.CancelButton = $btnHdrClose

    Set-ConsoleTheme -Form $dlg
    [void]$dlg.ShowDialog($form)
    $dlg.Dispose()
}
#endregion

#region Handlers -------------------------------------------------------------
$btnOpen.Add_Click({
    try { Open-ReportFolder -Path $ReportDir }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Open reports folder' }
})

$grid.Add_SelectionChanged({
    try { Update-CountLabel } catch { }
})

$btnSearch.Add_Click({
    $btnSearch.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = ''
    try {
        $filters = @{
            Recipient = $txtRcpt.Text.Trim()
            Sender    = $txtSender.Text.Trim()
            Subject   = $txtSubject.Text.Trim()
            Type      = "$($cboType.SelectedItem)"
        }
        $typeValue = "$($TypeMap[$filters.Type])"

        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity 'Quarantine search'
        $result = Get-QuarantineRows -Recipient $filters.Recipient -Sender $filters.Sender `
                      -Subject $filters.Subject -Days ([int]$numDays.Value) -Type $typeValue `
                      -Progress $progress
        Complete-ProgressReporter -Activity 'Quarantine search' -Bar $pbar

        Set-GridRows -Rows $result.Rows
        $txtOut.Text = Format-SearchSummary -Result $result -Filters $filters

        $note = ''
        if ($result.Capped) { $note = ' -- CAPPED' }
        $lblStatus.Text = "Done. $(@($result.Rows).Count) message(s) found$note."
        Write-CompletedLog -Task "Quarantine search -- $(@($result.Rows).Count) rows$note" -Result 'Success'
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Quarantine search'
    } finally {
        Complete-ProgressReporter -Activity 'Quarantine search'
        $form.Cursor       = $prev
        $btnSearch.Enabled = $true
    }
})

$btnHeaders.Add_Click({
    try {
        $rows = Get-SelectedRowObjects
        if ($rows.Count -eq 0) { $txtOut.Text = 'Select a row first.'; return }
        if ($rows.Count -gt 1) { $txtOut.Text = "Headers are shown for the first selected row only ($($rows.Count) selected)." }
        Show-HeaderDialog -Row $rows[0]
    } catch {
        $lblStatus.Text = 'Failed.'
        Show-HandlerError -ErrorRecord $_ -Context 'Preview headers'
    }
})

$btnRelease.Add_Click({
    try { Invoke-QuarantineAction -Mode 'Release' }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Release to recipient' }
})

$btnReleaseFP.Add_Click({
    try { Invoke-QuarantineAction -Mode 'ReleaseFP' }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Release and report false positive' }
})

$btnDelete.Add_Click({
    try { Invoke-QuarantineAction -Mode 'Delete' }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Delete from quarantine' }
})

$btnExport.Add_Click({
    try {
        $rows = Get-GridRowObjects
        if ($rows.Count -eq 0) { $txtOut.Text = 'Nothing to export -- run a search first.'; return }
        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir "QuarantineManager_$stamp.csv"
        $rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
        $files.SelectPath($csv)
        $lblStatus.Text = "Exported $($rows.Count) row(s)."
        $txtOut.Text = "Exported $($rows.Count) row(s) to:" + [Environment]::NewLine + "  $csv"
        Write-CompletedLog -Task "Quarantine export -- $($rows.Count) rows" -Result 'Success'
    } catch {
        Show-HandlerError -ErrorRecord $_ -Context 'Export CSV'
    }
})
#endregion

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled, and BringToFront'd inside the helper because WinForms
# draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Quarantine Manager' | Out-Null

[void]$form.ShowDialog()
