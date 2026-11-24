#Requires -Version 5.1
<#
    ReportBrowser.ps1
    Browses every CSV in the reports folder, across all report types. Author : Glenn Armbruster Revised : 2026-07-29

    Reasoning: ScriptNotes/ReportBrowser.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-07-31
#>
# TOOLKIT-HELP-BEGIN
# Opens any CSV the reporting tools have written, without going and finding it
# in Explorer.
#
# Every report writes a timestamped file to the reports folder. This lists them
# all, newest first, across every report type, so you can re-read last month's
# without re-running it.
#
# Read-only.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Configuration --------------------------------------------------------
$ReportDir = Get-ReportPath

#endregion

#region UI -------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Report Browser'
$form.Size            = New-Object System.Drawing.Size(900, 620)
$form.FormBorderStyle = 'Sizable'
$form.MinimumSize     = New-Object System.Drawing.Size(760, 460)
$form.MaximizeBox     = $true
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = "Every CSV in the reports folder. No connection required."
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(600, 18)
$form.Controls.Add($lblInfo)

$lblPath           = New-Object System.Windows.Forms.Label
$lblPath.Text      = $ReportDir
$lblPath.Location  = New-Object System.Drawing.Point(12, 32)
$lblPath.Size      = New-Object System.Drawing.Size(860, 18)
$lblPath.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblPath)

$lblFilter          = New-Object System.Windows.Forms.Label
$lblFilter.Text     = 'Report:'
$lblFilter.Location = New-Object System.Drawing.Point(12, 62)
$lblFilter.Size     = New-Object System.Drawing.Size(50, 20)
$form.Controls.Add($lblFilter)

$cboFilter               = New-Object System.Windows.Forms.ComboBox
$cboFilter.DropDownStyle = 'DropDownList'
$cboFilter.Location      = New-Object System.Drawing.Point(64, 59)
$cboFilter.Size          = New-Object System.Drawing.Size(240, 22)
$form.Controls.Add($cboFilter)

$btnOpenFile          = New-Object System.Windows.Forms.Button
$btnOpenFile.Text     = 'Open'
$btnOpenFile.Location = New-Object System.Drawing.Point(316, 57)
$btnOpenFile.Size     = New-Object System.Drawing.Size(90, 26)
$form.Controls.Add($btnOpenFile)

$btnFolder          = New-Object System.Windows.Forms.Button
$btnFolder.Text     = 'Open Folder'
$btnFolder.Location = New-Object System.Drawing.Point(412, 57)
$btnFolder.Size     = New-Object System.Drawing.Size(110, 26)
$form.Controls.Add($btnFolder)

$btnDelete          = New-Object System.Windows.Forms.Button
$btnDelete.Text     = 'Delete...'
$btnDelete.Location = New-Object System.Drawing.Point(528, 57)
$btnDelete.Size     = New-Object System.Drawing.Size(90, 26)
$form.Controls.Add($btnDelete)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(762, 57)
$btnClose.Size     = New-Object System.Drawing.Size(110, 26)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$lv                = New-Object System.Windows.Forms.ListView
$lv.Location       = New-Object System.Drawing.Point(12, 92)
$lv.Size           = New-Object System.Drawing.Size(860, 442)
$lv.View           = 'Details'
$lv.FullRowSelect  = $true
$lv.GridLines      = $true
$lv.MultiSelect    = $true
$lv.HideSelection  = $false
[void]$lv.Columns.Add('Run at',  130)
[void]$lv.Columns.Add('Report',  200)
[void]$lv.Columns.Add('Rows',     70)
[void]$lv.Columns.Add('Size',     80)
[void]$lv.Columns.Add('File',    360)
$form.Controls.Add($lv)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = ''
$lblStatus.Location = New-Object System.Drawing.Point(12, 542)
$lblStatus.Size     = New-Object System.Drawing.Size(860, 18)
$form.Controls.Add($lblStatus)

$lv.Anchor          = 'Top,Left,Bottom,Right'
$lblStatus.Anchor   = 'Bottom,Left,Right'
$lblPath.Anchor     = 'Top,Left,Right'
$btnClose.Anchor    = 'Top,Right'
#endregion

#region Behaviour ------------------------------------------------------------
$script:AllFiles  = @()
$script:FolderSig = $null

function Update-FilterChoices {
    $keep = "$($cboFilter.SelectedItem)"
    $cboFilter.Items.Clear()
    [void]$cboFilter.Items.Add('(all reports)')
    foreach ($r in (@($script:AllFiles) | Select-Object -ExpandProperty Report -Unique | Sort-Object)) {
        [void]$cboFilter.Items.Add($r)
    }
    $i = $cboFilter.Items.IndexOf($keep)
    $cboFilter.SelectedIndex = if ($i -ge 0) { $i } else { 0 }
}

function Update-List {
    $sel = "$($cboFilter.SelectedItem)"
    $show = if (-not $sel -or $sel -eq '(all reports)') {
                @($script:AllFiles)
            } else {
                @($script:AllFiles | Where-Object { $_.Report -eq $sel })
            }

    $lv.BeginUpdate()
    try {
        $lv.Items.Clear()
        foreach ($f in $show) {
            $rowsText = '-'
            if ($f.SizeKB -le 20480) {
                $n = Get-ReportRowCount -Path $f.Path
                if ($null -ne $n) { $rowsText = "$n" }
            }
            $mark = if ($f.StampFromName) { '' } else { '~' }

            $it = New-Object System.Windows.Forms.ListViewItem(
                    ($mark + $f.RunAt.ToString('yyyy-MM-dd HH:mm')))
            [void]$it.SubItems.Add($f.Report)
            [void]$it.SubItems.Add($rowsText)
            [void]$it.SubItems.Add($(if ($f.SizeKB -ge 1024) {
                        ('{0:N1} MB' -f ($f.SizeKB / 1024))
                    } else {
                        ('{0:N1} KB' -f $f.SizeKB)
                    }))
            [void]$it.SubItems.Add($f.Name)
            $it.Tag = $f.Path
            [void]$lv.Items.Add($it)
        }
    } finally {
        $lv.EndUpdate()
    }

    $total = (@($show) | Measure-Object SizeKB -Sum).Sum
    $lblStatus.Text = if (@($script:AllFiles).Count -eq 0) {
        'No reports yet. Run one, or use Open Folder to check the location.'
    } else {
        ('{0} file(s) shown of {1} in the folder -- {2:N1} MB' -f
            @($show).Count, @($script:AllFiles).Count, (([double]$total) / 1024))
    }
}

function Get-FolderSignature {
    $f = @(Get-ReportFileList -Directory $ReportDir)
    if (@($f).Count -eq 0) { return '0' }
    $sum = (@($f) | Measure-Object SizeKB -Sum).Sum
    return ('{0}|{1}|{2}' -f @($f).Count, $sum, $f[0].RunAt.Ticks)
}

function Update-All {
    param([switch]$Force)

    # Refresh on Activated so files written elsewhere appear -- but rebuilding
    # re-counts every row, so the folder signature short-circuits the common case.
    $sig = Get-FolderSignature
    if (-not $Force -and $sig -eq $script:FolderSig) { return }
    $script:FolderSig = $sig

    $script:AllFiles = @(Get-ReportFileList -Directory $ReportDir)
    Update-FilterChoices
    Update-List
}

function Get-SelectedPaths {
    return @($lv.SelectedItems | ForEach-Object { "$($_.Tag)" } | Where-Object { $_ })
}

$cboFilter.Add_SelectedIndexChanged({ Update-List })

$btnOpenFile.Add_Click({
    $paths = Get-SelectedPaths
    if (@($paths).Count -eq 0) { $lblStatus.Text = 'Select a report first.'; return }
    if (@($paths).Count -gt 5) {
        $lblStatus.Text = "$(@($paths).Count) files selected -- open five or fewer at a time."
        return
    }
    foreach ($p in $paths) {
        if (Test-Path -LiteralPath $p) { Start-Process $p }
    }
})

$lv.Add_DoubleClick({
    $paths = Get-SelectedPaths
    if (@($paths).Count -ge 1 -and (Test-Path -LiteralPath $paths[0])) { Start-Process $paths[0] }
})

$btnFolder.Add_Click({ Open-ReportFolder -Path $ReportDir })

$btnDelete.Add_Click({
    $paths = Get-SelectedPaths
    if (@($paths).Count -eq 0) { $lblStatus.Text = 'Select a report first.'; return }

    # GUARD: the only destructive control in the reporting suite. Named files in
    # the prompt, not just a count, because "delete 7 reports?" tells you
    # nothing about whether they were the seven you meant.
    $names = (@($paths) | ForEach-Object { Split-Path $_ -Leaf } | Select-Object -First 12) -join "`r`n  "
    $more  = if (@($paths).Count -gt 12) { "`r`n  ... and $(@($paths).Count - 12) more" } else { '' }
    $answer = [System.Windows.Forms.MessageBox]::Show(
        "Delete $(@($paths).Count) report file(s) permanently?`r`n`r`n  $names$more",
        'Delete reports', 'YesNo', 'Warning', 'Button2')
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $failed = @()
    foreach ($p in $paths) {
        try { Remove-Item -LiteralPath $p -Force -ErrorAction Stop }
        catch { $failed += "$(Split-Path $p -Leaf): $($_.Exception.Message)" }
    }
    Update-All -Force
    $lblStatus.Text = if ($failed.Count -gt 0) {
        "Deleted $(@($paths).Count - $failed.Count); $($failed.Count) failed -- $($failed[0])"
    } else {
        "Deleted $(@($paths).Count) file(s)."
    }
})

$form.Add_Activated({ Update-All })
#endregion

Update-All -Force

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Report Browser' | Out-Null
[void]$form.ShowDialog()
