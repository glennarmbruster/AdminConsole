#Requires -Version 5.1
# TOOLKIT-VARIANT: Any
<#
    ClearAdditions.ps1
    Clear the Config\Additions-*.json files that Add-ConsoleAdditions reads,
    once their tools are on buttons.

    Author   : Glenn Armbruster
    Created  : 2026-09-19
#>
# TOOLKIT-HELP-RUNS: nothing against any tenant. It reads Config\ and, on Clear, moves the ticked Additions-*.json files to Config\Backup\<stamp>\ and removes them from Config\.
# TOOLKIT-HELP-BEGIN
# Clears the additions files so Publish-Toolkit stops asking about tools that
# are already on buttons.
#
# WHAT AN ADDITIONS FILE IS. A delivery can come with Config\Additions-*.json
# -- a curated list saying 'put this script on a button in this section'.
# Add-ConsoleAdditions reads it (Publish-Toolkit runs that at the end of every
# publish) and adds whatever is missing. Once the tools are on buttons the
# file has done its job, but it stays in Config\ and every later publish reads
# it again and asks. This window is the one-click way to be rid of it.
#
# EACH FILE IS SHOWN WITH ITS ENTRIES, and each entry says whether the script
# is ALREADY on a button in Tools.json. A file whose every entry says 'on a
# button' is the stale one; it is ticked for you. A file with entries that are
# NOT yet on a button is listed unticked, because clearing it would lose the
# instruction to add them -- read it first.
#
# NOTHING IS DELETED. Clear moves the ticked files to Config\Backup\<stamp>\,
# the same place Add-ConsoleAdditions keeps its own backups. Put one back by
# copying it up a level.
#
# Additions-Ignore.json IS LISTED SEPARATELY AND UNTICKED. It is not a delivery
# file: it is the list of scripts you declined, kept so you are not asked
# about them again. Clear it only if you want to be asked afresh.
# TOOLKIT-HELP-END

. (Join-Path $PSScriptRoot 'Common.ps1')

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:ConfigDir = Join-Path $PSScriptRoot 'Config'

#region Read --------------------------------------------------------------------
function Read-AdditionsJson {
    param([string]$Path)
    try { return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { return $null }
}

function Get-ButtonScriptSet {
    # Every script Tools.json already points at, lower-cased leaf name.
    $h = @{}
    $t = Read-AdditionsJson -Path (Join-Path $script:ConfigDir 'Tools.json')
    $items = @($(if ($t -and $t.PSObject.Properties['Tools']) { $t.Tools } else { $t }))
    foreach ($x in $items) {
        $s = "$($x.Script)".Trim()
        if ($s) { $h[(Split-Path -Leaf $s).ToLower()] = $true }
    }
    return $h
}

function Get-AdditionsFileList {
    <#
        One row per Additions-*.json in Config\, with its entries and a
        verdict: Stale (every entry already on a button), Pending (some not),
        Empty, Invalid, or Ignore for the declined-list file.
    #>
    $onButton = Get-ButtonScriptSet
    $rows = @()
    $files = @(Get-ChildItem -LiteralPath $script:ConfigDir -Filter 'Additions-*.json' -File -ErrorAction SilentlyContinue | Sort-Object Name)
    foreach ($f in $files) {
        $r = [PSCustomObject]@{ Name = $f.Name; Path = $f.FullName; Kind = ''; Entries = @(); Pending = 0; Detail = '' }
        if ($f.Name -like '*Ignore*') {
            $j = Read-AdditionsJson -Path $f.FullName
            $r.Kind = 'Ignore'
            $n = $(if ($j) { @($j.Ignore).Count } else { 0 })
            $r.Detail = "declined list -- $n script(s) you said no to; clearing it means being asked again"
            $rows += $r; continue
        }
        $j = Read-AdditionsJson -Path $f.FullName
        if (-not $j) { $r.Kind = 'Invalid'; $r.Detail = 'not valid JSON -- Add-ConsoleAdditions skips it'; $rows += $r; continue }
        $tools = @($j.Tools)
        $themes = @($(if ($j.PSObject.Properties['Themes']) { $j.Themes } else { @() }))
        if ($tools.Count -eq 0 -and $themes.Count -eq 0) { $r.Kind = 'Empty'; $r.Detail = 'no tools or themes listed'; $rows += $r; continue }
        $entries = @()
        foreach ($t in $tools) {
            $leaf = (Split-Path -Leaf "$($t.Script)").ToLower()
            $have = [bool]$onButton[$leaf]
            $exists = Test-Path -LiteralPath (Join-Path $PSScriptRoot (Split-Path -Leaf "$($t.Script)"))
            $entries += [PSCustomObject]@{
                Name = "$($t.Name)"; Script = "$($t.Script)"; Section = "$($t.Section)"
                State = $(if ($have) { 'on a button' } elseif (-not $exists) { 'NOT on a button, and the file is not in the folder' } else { 'NOT on a button yet' })
                Done = $have
            }
        }
        $r.Entries = $entries
        $r.Pending = @($entries | Where-Object { -not $_.Done }).Count
        $r.Kind = $(if ($r.Pending -eq 0) { 'Stale' } else { 'Pending' })
        $r.Detail = ('{0} tool(s){1}: {2} on a button, {3} not' -f $tools.Count,
                     $(if ($themes.Count) { ", $($themes.Count) theme(s)" } else { '' }),
                     ($tools.Count - $r.Pending), $r.Pending)
        $rows += $r
    }
    return $rows
}
#endregion

#region Form --------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text          = 'Clear Additions File'
$form.ClientSize    = New-Object System.Drawing.Size(820, 560)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblWhy = New-Object System.Windows.Forms.Label
$lblWhy.Location = New-Object System.Drawing.Point(12, 10)
$lblWhy.Size     = New-Object System.Drawing.Size(796, 34)
$lblWhy.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblWhy.Text     = 'Additions files tell Add-ConsoleAdditions which scripts to put on buttons. Once they are on buttons the file is stale, and every publish asks about it again until it is cleared.'
$form.Controls.Add($lblWhy)

$lblWhere = New-Object System.Windows.Forms.Label
$lblWhere.Location  = New-Object System.Drawing.Point(12, 46)
$lblWhere.Size      = New-Object System.Drawing.Size(796, 18)
$lblWhere.ForeColor = (Get-ToolStatusColour -Kind Muted)
$lblWhere.Text      = "In: $($script:ConfigDir)     Cleared files go to Config\Backup\<stamp>\ -- nothing is deleted."
$form.Controls.Add($lblWhere)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location            = New-Object System.Drawing.Point(12, 70)
$grid.Size                = New-Object System.Drawing.Size(796, 170)
$grid.ReadOnly            = $false
$grid.AllowUserToAddRows  = $false
$grid.SelectionMode       = 'FullRowSelect'
$grid.MultiSelect         = $false
$grid.RowHeadersVisible   = $false
$grid.AutoSizeColumnsMode = 'None'
$grid.Font                = New-Object System.Drawing.Font('Consolas', 9)
$colPick = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colPick.HeaderText = 'Clear'; $colPick.Name = 'Clear'; $colPick.Width = 50
[void]$grid.Columns.Add($colPick)
foreach ($c in @(@('File', 230), @('Verdict', 70), @('Detail', 430))) {
    $i = $grid.Columns.Add($c[0], $c[0]); $grid.Columns[$i].Width = $c[1]; $grid.Columns[$i].ReadOnly = $true
}
$form.Controls.Add($grid)

$lblEntries = New-Object System.Windows.Forms.Label
$lblEntries.Location = New-Object System.Drawing.Point(12, 248)
$lblEntries.Size     = New-Object System.Drawing.Size(796, 18)
$lblEntries.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblEntries.Text     = 'WHAT THE SELECTED FILE SAYS'
$form.Controls.Add($lblEntries)

$txtEntries = New-Object System.Windows.Forms.TextBox
$txtEntries.Multiline  = $true
$txtEntries.ReadOnly   = $true
$txtEntries.ScrollBars = 'Vertical'
$txtEntries.WordWrap   = $false
$txtEntries.Location   = New-Object System.Drawing.Point(12, 268)
$txtEntries.Size       = New-Object System.Drawing.Size(796, 220)
$txtEntries.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtEntries.Tag        = 'NoTheme'
$txtEntries.BackColor  = [System.Drawing.Color]::WhiteSmoke
$txtEntries.ForeColor  = [System.Drawing.Color]::Black
$form.Controls.Add($txtEntries)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text     = 'Clear ticked'
$btnClear.Location = New-Object System.Drawing.Point(12, 500)
$btnClear.Size     = New-Object System.Drawing.Size(130, 30)
$form.Controls.Add($btnClear)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text     = 'Refresh'
$btnRefresh.Location = New-Object System.Drawing.Point(150, 500)
$btnRefresh.Size     = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnRefresh)

$btnOpen = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Config folder'
$btnOpen.Location = New-Object System.Drawing.Point(258, 500)
$btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnOpen)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location  = New-Object System.Drawing.Point(420, 506)
$lblStatus.Size      = New-Object System.Drawing.Size(270, 18)
$lblStatus.ForeColor = (Get-ToolStatusColour -Kind Muted)
$form.Controls.Add($lblStatus)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(698, 500)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

function Set-ClearStatus {
    param([string]$Text, [switch]$Bad)
    $lblStatus.Text = $Text
    $lblStatus.ForeColor = $(if ($Bad) { (Get-ToolStatusColour -Kind Bad) } else { (Get-ToolStatusColour -Kind Muted) })
}

function Update-FileGrid {
    $grid.Rows.Clear()
    $txtEntries.Text = ''
    $rows = @(Get-AdditionsFileList)
    foreach ($r in $rows) {
        # STALE IS TICKED FOR YOU; everything else is a decision.
        $i = $grid.Rows.Add(($r.Kind -eq 'Stale'), $r.Name, $r.Kind, $r.Detail)
        $grid.Rows[$i].Tag = $r
        switch ($r.Kind) {
            'Stale'   { $grid.Rows[$i].Cells[2].Style.ForeColor = (Get-ToolStatusColour -Kind Good) }
            'Pending' { $grid.Rows[$i].Cells[2].Style.ForeColor = (Get-ToolStatusColour -Kind Warn) }
            'Invalid' { $grid.Rows[$i].Cells[2].Style.ForeColor = (Get-ToolStatusColour -Kind Bad) }
        }
    }
    $grid.ClearSelection()
    if ($rows.Count -eq 0) {
        $txtEntries.Text = 'No Additions-*.json files in Config\. There is nothing for Publish-Toolkit to ask about.'
        Set-ClearStatus -Text 'Nothing to clear.'
    } else {
        $stale = @($rows | Where-Object { $_.Kind -eq 'Stale' }).Count
        Set-ClearStatus -Text ('{0} file(s); {1} stale and ticked.' -f $rows.Count, $stale)
        if ($grid.Rows.Count) { $grid.Rows[0].Selected = $true }
    }
    $btnClear.Enabled = ($rows.Count -gt 0)
}

function Update-EntriesBox {
    $sel = $grid.SelectedRows
    if ($sel.Count -eq 0) { $txtEntries.Text = ''; return }
    $r = $sel[0].Tag
    if (-not $r) { $txtEntries.Text = ''; return }
    $L = New-Object System.Collections.ArrayList
    [void]$L.Add("$($r.Name)  --  $($r.Kind): $($r.Detail)")
    [void]$L.Add('')
    if ($r.Kind -eq 'Ignore') {
        $j = Read-AdditionsJson -Path $r.Path
        foreach ($n in @($(if ($j) { $j.Ignore } else { @() }))) { [void]$L.Add("   declined   $n") }
    } elseif ($r.Entries.Count) {
        foreach ($e in $r.Entries) {
            [void]$L.Add(('   {0,-14} {1,-28} {2,-30} {3}' -f $(if ($e.Done) { 'on a button' } else { 'NOT YET' }), $e.Name, (Split-Path -Leaf $e.Script), $(if ($e.Section) { "[$($e.Section)]" } else { '' })))
        }
        if ($r.Pending -gt 0) {
            [void]$L.Add('')
            [void]$L.Add('Some entries are NOT on a button. Clearing this file drops the instruction to add them;')
            [void]$L.Add('run Add-ConsoleAdditions.ps1 -Apply first if you want them, then clear.')
        }
    }
    $txtEntries.Text = ($L -join "`r`n")
}

$grid.Add_SelectionChanged({ Update-EntriesBox })
$grid.Add_CellContentClick({ if ($_.ColumnIndex -eq 0) { $grid.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) } })

$btnRefresh.Add_Click({ Update-FileGrid })
$btnOpen.Add_Click({ try { Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$($script:ConfigDir)`"" } catch { } })

$btnClear.Add_Click({
    $picked = @()
    foreach ($row in $grid.Rows) { if ([bool]$row.Cells[0].Value) { $picked += $row.Tag } }
    if ($picked.Count -eq 0) { Set-ClearStatus -Text 'Tick a file first.' -Bad; return }
    $pending = @($picked | Where-Object { $_.Kind -eq 'Pending' })
    $ignore  = @($picked | Where-Object { $_.Kind -eq 'Ignore' })
    $msg = "Move $($picked.Count) file(s) out of Config\ into Config\Backup\?" + "`r`n`r`n" +
           (($picked | ForEach-Object { "  $($_.Name)  ($($_.Kind))" }) -join "`r`n")
    if ($pending.Count) { $msg += "`r`n`r`n" + "$($pending.Count) of these still list tools that are NOT on a button. Those instructions are lost with the file." }
    if ($ignore.Count)  { $msg += "`r`n`r`n" + 'The declined list is ticked: you will be asked about every declined script again.' }
    $msg += "`r`n`r`n" + 'Nothing is deleted; the files are in the backup folder if you need them.'
    $a = Show-ToolDialog($msg, 'Clear additions files', 'OKCancel', $(if ($pending.Count -or $ignore.Count) { 'Warning' } else { 'Question' }), 'Button2')
    if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $bak = Join-Path (Join-Path $script:ConfigDir 'Backup') $stamp
    $ok = @(); $bad = @()
    try {
        if (-not (Test-Path -LiteralPath $bak)) { New-Item -ItemType Directory -Path $bak -Force | Out-Null }
        foreach ($p in $picked) {
            try {
                Move-Item -LiteralPath $p.Path -Destination (Join-Path $bak $p.Name) -Force -ErrorAction Stop
                $ok += $p.Name
            } catch { $bad += "$($p.Name): $($_.Exception.Message)" }
        }
    } catch { $bad += "backup folder: $($_.Exception.Message)" }
    try {
        Write-CompletedLog -Action 'ClearAdditions' -Target 'Config' `
            -Result $(if ($bad.Count -and $ok.Count) { 'Partial' } elseif ($bad.Count) { 'Failed' } else { 'Success' }) `
            -Detail ("Moved to Backup\$stamp`: $($ok -join ', ')" + $(if ($bad.Count) { "; FAILED: $($bad -join '; ')" } else { '' }))
    } catch { }
    Update-FileGrid
    if ($bad.Count) { Set-ClearStatus -Text "Cleared $($ok.Count); failed $($bad.Count) -- $($bad[0])" -Bad }
    else { Set-ClearStatus -Text "Cleared $($ok.Count) file(s) to Backup\$stamp. Publish-Toolkit will not ask about them again." }
})
#endregion

try { Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Clear Additions File' | Out-Null } catch { }
try { Set-ConsoleTheme -Form $form } catch { }
$form.Add_Shown({ Update-FileGrid })
[void]$form.ShowDialog()
try { $form.Dispose() } catch { }
