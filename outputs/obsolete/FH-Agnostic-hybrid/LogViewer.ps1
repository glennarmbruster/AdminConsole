#Requires -Version 5.1
<#
    LogViewer.ps1
    Read-only view over the completed-tasks audit log. Reads any CompletedTasks*.csv in the log folder, including the old Task,DateTime format, and presen

    Reasoning: ScriptNotes/LogViewer.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-07-31
#>
# TOOLKIT-HELP-BEGIN
# A read-only view of the audit log -- every change the toolkit's tools have
# made, who made it, when, and to whom.
#
# This is the answer to "when did that happen and who did it". Search and filter
# across everything, including logs in the older format.
#
# Nothing here changes anything. The log is the record; this reads it.
# TOOLKIT-HELP-END

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

$script:AllRows  = @()
$script:LogFiles = @()

# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------

function Get-LogFileList {
    $folder = Get-LogPath
    if (-not (Test-Path -LiteralPath $folder)) { return @() }
    return @(Get-ChildItem -LiteralPath $folder -Filter 'CompletedTasks*.csv' -File -ErrorAction SilentlyContinue |
             Sort-Object LastWriteTime -Descending)
}

function ConvertTo-ViewRow {
    param($Raw, [string]$SourceFile)

    $isV2 = ($null -ne $Raw.PSObject.Properties['Timestamp'])

    if ($isV2) {
        $ts = "$($Raw.Timestamp)"
        $detail = "$($Raw.Detail)"
        $action = "$($Raw.Action)"
        $target = "$($Raw.Target)"
        $envLbl = "$($Raw.Environment)"
        $result = "$($Raw.Result)"
        $oper   = "$($Raw.Operator)"
        $mach   = "$($Raw.Machine)"
        $scr    = "$($Raw.Script)"
        $run    = "$($Raw.RunId)"
    } else {
        $ts     = "$($Raw.DateTime)"
        $detail = "$($Raw.Task)"
        $parsed = Split-LegacyTask -Task $detail
        $action = $parsed.Action
        $target = $parsed.Target
        $envLbl = $parsed.Environment
        $result = ''
        $oper   = ''
        $mach   = ''
        $scr    = ''
        $run    = ''
    }

    $when = $null
    if ($ts) {
        try { $when = [datetime]::Parse($ts, [Globalization.CultureInfo]::InvariantCulture) } catch { }
        # $null -eq, not -not: PowerShell will not coerce a [datetime] to a bool.
        if ($null -eq $when) {
            try {
                $when = [datetime]::ParseExact($ts, 'yyyy-MM-dd HHmm',
                            [Globalization.CultureInfo]::InvariantCulture)
            } catch { }
        }
    }

    return [PSCustomObject]@{
        When        = $when
        Timestamp   = $ts
        Operator    = $oper
        Machine     = $mach
        Script      = $scr
        Action      = $action
        Target      = $target
        Result      = $result
        Detail      = $detail
        RunId       = $run
        Environment = $envLbl
        SourceFile  = $SourceFile
    }
}

function Import-LogRows {
    param([string[]]$Paths)
    $rows = @()
    foreach ($p in $Paths) {
        if (-not (Test-Path -LiteralPath $p)) { continue }
        try { $raw = @(Import-Csv -LiteralPath $p -ErrorAction Stop) }
        catch {
            [void][System.Windows.Forms.MessageBox]::Show(
                "Could not read $(Split-Path -Leaf $p):`r`n$($_.Exception.Message)",
                'Log Viewer', 'OK', 'Warning')
            continue
        }
        $leaf = Split-Path -Leaf $p
        foreach ($r in $raw) { $rows += (ConvertTo-ViewRow -Raw $r -SourceFile $leaf) }
    }
    return @($rows | Sort-Object -Property @{ Expression = { $_.When }; Descending = $true })
}

# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Audit Log Viewer'
$form.Size = New-Object System.Drawing.Size(1220, 820)
$form.StartPosition = 'Manual'
$form.MinimumSize = New-Object System.Drawing.Size(900, 620)
$form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

# --- top row: file + refresh ------------------------------------------------
$lblFile = New-Object System.Windows.Forms.Label
$lblFile.Text = 'Log file:'
$lblFile.Location = New-Object System.Drawing.Point(12, 15)
$lblFile.Size = New-Object System.Drawing.Size(60, 20)
$form.Controls.Add($lblFile)

$cboFile = New-Object System.Windows.Forms.ComboBox
$cboFile.Location = New-Object System.Drawing.Point(76, 12)
$cboFile.Size = New-Object System.Drawing.Size(320, 22)
$cboFile.DropDownStyle = 'DropDownList'
$form.Controls.Add($cboFile)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = 'Refresh'
$btnRefresh.Location = New-Object System.Drawing.Point(404, 11)
$btnRefresh.Size = New-Object System.Drawing.Size(80, 25)
$form.Controls.Add($btnRefresh)

$btnFolder = New-Object System.Windows.Forms.Button
$btnFolder.Text = 'Open Folder'
$btnFolder.Location = New-Object System.Drawing.Point(490, 11)
$btnFolder.Size = New-Object System.Drawing.Size(95, 25)
$form.Controls.Add($btnFolder)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = 'Export View'
$btnExport.Location = New-Object System.Drawing.Point(591, 11)
$btnExport.Size = New-Object System.Drawing.Size(95, 25)
$form.Controls.Add($btnExport)

$lblPath = New-Object System.Windows.Forms.Label
$lblPath.Location = New-Object System.Drawing.Point(696, 16)
$lblPath.Size = New-Object System.Drawing.Size(380, 18)
$lblPath.AutoEllipsis = $true
$form.Controls.Add($lblPath)

# --- filter row -------------------------------------------------------------
$lblFind = New-Object System.Windows.Forms.Label
$lblFind.Text = 'Search:'
$lblFind.Location = New-Object System.Drawing.Point(12, 50)
$lblFind.Size = New-Object System.Drawing.Size(60, 20)
$form.Controls.Add($lblFind)

$txtFind = New-Object System.Windows.Forms.TextBox
$txtFind.Location = New-Object System.Drawing.Point(76, 47)
$txtFind.Size = New-Object System.Drawing.Size(230, 22)
$form.Controls.Add($txtFind)

$lblResult = New-Object System.Windows.Forms.Label
$lblResult.Text = 'Result:'
$lblResult.Location = New-Object System.Drawing.Point(316, 50)
$lblResult.Size = New-Object System.Drawing.Size(45, 20)
$form.Controls.Add($lblResult)

$cboResult = New-Object System.Windows.Forms.ComboBox
$cboResult.Location = New-Object System.Drawing.Point(363, 47)
$cboResult.Size = New-Object System.Drawing.Size(110, 22)
$cboResult.DropDownStyle = 'DropDownList'
$form.Controls.Add($cboResult)

$lblScript = New-Object System.Windows.Forms.Label
$lblScript.Text = 'Script:'
$lblScript.Location = New-Object System.Drawing.Point(483, 50)
$lblScript.Size = New-Object System.Drawing.Size(42, 20)
$form.Controls.Add($lblScript)

$cboScript = New-Object System.Windows.Forms.ComboBox
$cboScript.Location = New-Object System.Drawing.Point(527, 47)
$cboScript.Size = New-Object System.Drawing.Size(170, 22)
$cboScript.DropDownStyle = 'DropDownList'
$form.Controls.Add($cboScript)

$lblOper = New-Object System.Windows.Forms.Label
$lblOper.Text = 'Operator:'
$lblOper.Location = New-Object System.Drawing.Point(707, 50)
$lblOper.Size = New-Object System.Drawing.Size(58, 20)
$form.Controls.Add($lblOper)

$cboOper = New-Object System.Windows.Forms.ComboBox
$cboOper.Location = New-Object System.Drawing.Point(767, 47)
$cboOper.Size = New-Object System.Drawing.Size(180, 22)
$cboOper.DropDownStyle = 'DropDownList'
$form.Controls.Add($cboOper)

$chkDate = New-Object System.Windows.Forms.CheckBox
$chkDate.Text = 'From'
$chkDate.Location = New-Object System.Drawing.Point(957, 49)
$chkDate.Size = New-Object System.Drawing.Size(52, 20)
$form.Controls.Add($chkDate)

$dtpFrom = New-Object System.Windows.Forms.DateTimePicker
$dtpFrom.Location = New-Object System.Drawing.Point(1009, 47)
$dtpFrom.Size = New-Object System.Drawing.Size(110, 22)
$dtpFrom.Format = 'Short'
$dtpFrom.Value = (Get-Date).AddDays(-30)
$dtpFrom.Enabled = $false
$form.Controls.Add($dtpFrom)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = 'Clear'
$btnClear.Location = New-Object System.Drawing.Point(1125, 46)
$btnClear.Size = New-Object System.Drawing.Size(65, 25)
$form.Controls.Add($btnClear)

# --- grid -------------------------------------------------------------------
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(12, 82)
$grid.Size = New-Object System.Drawing.Size(1178, 420)
$grid.Anchor = 'Top,Left,Right,Bottom'
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToOrderColumns = $true
$grid.RowHeadersVisible = $false
$grid.SelectionMode = 'FullRowSelect'
$grid.MultiSelect = $true
$grid.AutoSizeColumnsMode = 'None'
$form.Controls.Add($grid)

# --- detail pane ------------------------------------------------------------
$lblDetail = New-Object System.Windows.Forms.Label
$lblDetail.Text = 'Selected row:'
$lblDetail.Location = New-Object System.Drawing.Point(12, 512)
$lblDetail.Size = New-Object System.Drawing.Size(200, 18)
$lblDetail.Anchor = 'Left,Bottom'
$form.Controls.Add($lblDetail)

$txtDetail = New-Object System.Windows.Forms.TextBox
$txtDetail.Location = New-Object System.Drawing.Point(12, 532)
$txtDetail.Size = New-Object System.Drawing.Size(1178, 210)
$txtDetail.Multiline = $true
$txtDetail.ReadOnly = $true
$txtDetail.ScrollBars = 'Vertical'
$txtDetail.Anchor = 'Left,Right,Bottom'
$form.Controls.Add($txtDetail)

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Location = New-Object System.Drawing.Point(12, 750)
$lblCount.Size = New-Object System.Drawing.Size(900, 20)
$lblCount.Anchor = 'Left,Bottom'
$form.Controls.Add($lblCount)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(1110, 746)
$btnClose.Size = New-Object System.Drawing.Size(80, 26)
$btnClose.Anchor = 'Right,Bottom'
$form.Controls.Add($btnClose)

# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------

function Set-GridColumns {
    $grid.Columns.Clear()
    $spec = @(
        @{ N = 'Timestamp';   W = 160 }
        @{ N = 'Result';      W = 65  }
        @{ N = 'Action';      W = 80  }
        @{ N = 'Target';      W = 210 }
        @{ N = 'Script';      W = 150 }
        @{ N = 'Operator';    W = 150 }
        @{ N = 'Environment'; W = 90  }
        @{ N = 'Machine';     W = 100 }
        @{ N = 'RunId';       W = 90  }
        @{ N = 'Detail';      W = 400 }
        @{ N = 'SourceFile';  W = 150 }
    )
    foreach ($s in $spec) {
        $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $c.Name = $s.N; $c.HeaderText = $s.N; $c.Width = $s.W
        [void]$grid.Columns.Add($c)
    }
}

function Update-Filters {
    $keepR = "$($cboResult.SelectedItem)"
    $keepS = "$($cboScript.SelectedItem)"
    $keepO = "$($cboOper.SelectedItem)"

    $cboResult.Items.Clear(); [void]$cboResult.Items.Add('(all)')
    foreach ($v in (@($script:AllRows | ForEach-Object { $_.Result }) | Where-Object { $_ } | Sort-Object -Unique)) {
        [void]$cboResult.Items.Add($v)
    }
    $cboScript.Items.Clear(); [void]$cboScript.Items.Add('(all)')
    foreach ($v in (@($script:AllRows | ForEach-Object { $_.Script }) | Where-Object { $_ } | Sort-Object -Unique)) {
        [void]$cboScript.Items.Add($v)
    }
    $cboOper.Items.Clear(); [void]$cboOper.Items.Add('(all)')
    foreach ($v in (@($script:AllRows | ForEach-Object { $_.Operator }) | Where-Object { $_ } | Sort-Object -Unique)) {
        [void]$cboOper.Items.Add($v)
    }

    if ($cboResult.Items.Contains($keepR)) { $cboResult.SelectedItem = $keepR } else { $cboResult.SelectedIndex = 0 }
    if ($cboScript.Items.Contains($keepS)) { $cboScript.SelectedItem = $keepS } else { $cboScript.SelectedIndex = 0 }
    if ($cboOper.Items.Contains($keepO))   { $cboOper.SelectedItem   = $keepO } else { $cboOper.SelectedIndex   = 0 }
}

function Update-Grid {
    $find = "$($txtFind.Text)".Trim()
    $fr = "$($cboResult.SelectedItem)"
    $fs = "$($cboScript.SelectedItem)"
    $fo = "$($cboOper.SelectedItem)"

    $rows = @($script:AllRows)
    if ($fr -and $fr -ne '(all)') { $rows = @($rows | Where-Object { $_.Result   -eq $fr }) }
    if ($fs -and $fs -ne '(all)') { $rows = @($rows | Where-Object { $_.Script   -eq $fs }) }
    if ($fo -and $fo -ne '(all)') { $rows = @($rows | Where-Object { $_.Operator -eq $fo }) }
    if ($chkDate.Checked) {
        $from = $dtpFrom.Value.Date
        $rows = @($rows | Where-Object { $_.When -and $_.When -ge $from })
    }
    if ($find) {
        $rows = @($rows | Where-Object {
            "$($_.Target)$($_.Detail)$($_.Action)$($_.Operator)$($_.Script)$($_.RunId)" `
                -like "*$find*"
        })
    }

    $grid.SuspendLayout()
    $grid.Rows.Clear()
    foreach ($r in $rows) {
        $i = $grid.Rows.Add($r.Timestamp, $r.Result, $r.Action, $r.Target, $r.Script,
                            $r.Operator, $r.Environment, $r.Machine, $r.RunId,
                            $r.Detail, $r.SourceFile)
        switch ($r.Result) {
            'Failed'  { $grid.Rows[$i].Cells['Result'].Style.ForeColor = [System.Drawing.Color]::Firebrick }
            'Partial' { $grid.Rows[$i].Cells['Result'].Style.ForeColor = [System.Drawing.Color]::DarkOrange }
        }
    }
    $grid.ResumeLayout()

    $failed  = @($rows | Where-Object { $_.Result -eq 'Failed'  }).Count
    $partial = @($rows | Where-Object { $_.Result -eq 'Partial' }).Count
    $msg = "$(@($rows).Count) of $(@($script:AllRows).Count) rows"
    if ($failed -or $partial) { $msg += "   --   $failed failed, $partial partial" }
    $lblCount.Text = $msg
}

function Import-SelectedFile {
    $sel = "$($cboFile.SelectedItem)"
    if (-not $sel) { return }
    $folder = Get-LogPath
    if ($sel -eq '(all files)') {
        $paths = @($script:LogFiles | ForEach-Object { $_.FullName })
        $lblPath.Text = $folder
    } else {
        $paths = @(Join-Path $folder $sel)
        $lblPath.Text = $paths[0]
    }
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        # @() at the ASSIGNMENT: 'return @(...)' unrolls a single-element array
        # to a scalar, and .Count on a lone PSCustomObject is $null.
        $script:AllRows = @(Import-LogRows -Paths $paths)
        Update-Filters
        Update-Grid
    } finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
}

function Update-FileList {
    $script:LogFiles = Get-LogFileList
    $cboFile.Items.Clear()
    if ($script:LogFiles.Count -eq 0) {
        $lblPath.Text = "No log files in $(Get-LogPath)"
        $grid.Rows.Clear()
        $lblCount.Text = '0 rows'
        return
    }
    foreach ($f in $script:LogFiles) { [void]$cboFile.Items.Add($f.Name) }
    if ($script:LogFiles.Count -gt 1) { [void]$cboFile.Items.Add('(all files)') }
    $cboFile.SelectedIndex = 0
}

$cboFile.Add_SelectedIndexChanged({ Import-SelectedFile })
$txtFind.Add_TextChanged({ Update-Grid })
$cboResult.Add_SelectedIndexChanged({ Update-Grid })
$cboScript.Add_SelectedIndexChanged({ Update-Grid })
$cboOper.Add_SelectedIndexChanged({ Update-Grid })
$chkDate.Add_CheckedChanged({ $dtpFrom.Enabled = $chkDate.Checked; Update-Grid })
$dtpFrom.Add_ValueChanged({ if ($chkDate.Checked) { Update-Grid } })

$btnClear.Add_Click({
    $txtFind.Text = ''
    $cboResult.SelectedIndex = 0
    $cboScript.SelectedIndex = 0
    $cboOper.SelectedIndex = 0
    $chkDate.Checked = $false
})

$btnRefresh.Add_Click({ Update-FileList })

$btnFolder.Add_Click({
    $folder = Get-LogPath
    if (Test-Path -LiteralPath $folder) { Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$folder`"" }
})

$btnExport.Add_Click({
    if ($grid.Rows.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('Nothing to export.', 'Log Viewer', 'OK', 'Information'); return
    }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'CSV (*.csv)|*.csv'
    $dlg.FileName = "AuditView_$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    $dlg.InitialDirectory = Get-ReportPath
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    try {
        $out = @()
        foreach ($row in $grid.Rows) {
            $o = [ordered]@{}
            foreach ($c in $grid.Columns) { $o[$c.Name] = "$($row.Cells[$c.Name].Value)" }
            $out += [PSCustomObject]$o
        }
        $out | Export-Csv -LiteralPath $dlg.FileName -NoTypeInformation -Encoding UTF8
        [void][System.Windows.Forms.MessageBox]::Show("Exported $($out.Count) rows.", 'Log Viewer', 'OK', 'Information')
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show("Export failed:`r`n$($_.Exception.Message)",
            'Log Viewer', 'OK', 'Warning')
    }
})

$grid.Add_SelectionChanged({
    if ($grid.SelectedRows.Count -eq 0) { $txtDetail.Text = ''; return }
    $r = $grid.SelectedRows[0]
    $nl = [Environment]::NewLine
    $sb = New-Object System.Text.StringBuilder
    $order = @('Timestamp', 'Result', 'Action', 'Target', 'Script', 'Operator',
               'Environment', 'Machine', 'RunId', 'SourceFile', 'Detail')
    foreach ($n in $order) {
        if (-not $grid.Columns.Contains($n)) { continue }
        $v = "$($r.Cells[$n].Value)"
        if ($v) { [void]$sb.Append(("{0,-12}: {1}{2}" -f $n, $v, $nl)) }
    }
    $txtDetail.Text = $sb.ToString()
})

$grid.Add_CellDoubleClick({
    if ($grid.SelectedRows.Count -eq 0) { return }
    $run = "$($grid.SelectedRows[0].Cells['RunId'].Value)"
    if (-not $run) { return }
    $txtFind.Text = $run
})

$btnClose.Add_Click({ $form.Close() })

# ---------------------------------------------------------------------------
Set-GridColumns
Update-FileList

$script:Theme = Get-ConsoleTheme
if ($script:Theme) { Set-ConsoleTheme -Form $form -Theme $script:Theme }
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Log Viewer' | Out-Null
Add-FormLogo -Form $form -Corner 'TopRight' -MaxHeight 28 -Margin 6 | Out-Null
# BottomLeft: TopLeft belongs to AdminConsole.ps1 itself.
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'

[void]$form.ShowDialog()
