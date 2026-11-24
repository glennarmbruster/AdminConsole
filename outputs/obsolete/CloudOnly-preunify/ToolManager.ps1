#Requires -Version 5.1
<#
    ToolManager.ps1
    Editor for Tools.json -- add, remove, rename, reorder, and repoint the
    tools AdminConsole.ps1 displays, with no code changes required.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-RUNS: nothing. It edits a configuration file on this machine.
# TOOLKIT-HELP-BEGIN
# Edits Tools.json -- which buttons appear on the console, what they are
# called, what order they sit in, and which script each one runs.
#
# Add, remove, rename, reorder and repoint tools without touching any code.
#
# Each entry carries:
#
#   Name          the button text
#   Script        the .ps1 it launches, relative to the toolkit folder
#   Section       which group it appears under
#   Order         position within that section
#   ObjectTypes   which lookup results enable it -- leave empty for always on
#
# A button pointing at a script that is not there says nothing until someone
# clicks it. Test-Toolkit checks every entry against the folder and names any
# that do not resolve -- worth running after editing here.
#
# Tools.json is per-engineer and is deliberately NOT covered by the version
# manifest, so your button layout is yours and will not report as drift.
#
# Local only. It writes one JSON file and connects to nothing.
# TOOLKIT-HELP-END

param(
    [string]$ToolsConfigPath = ''
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# Config now lives in Config\. Resolved after Common.ps1 loads, because the
# resolver is defined there -- a param default cannot call it.
if (-not $ToolsConfigPath) { $ToolsConfigPath = Get-ConfigFilePath -Name 'Tools.json' }

$validObjectTypes = @('Mailbox', 'UnifiedGroup', 'OnpremGroup')

# Sections are user-defined now and live in Sections.json alongside Tools.json.
# Each section is ButtonsPerRow wide and as many rows as it declares, so its
# capacity is simply rows x width -- nothing here is hardcoded to a name.
$SectionsConfigPath = Get-ConfigFilePath -Name 'Sections.json'

function Read-SectionsJson {
    $default = [PSCustomObject]@{
        ButtonsPerRow = 5
        Sections = @(
            [PSCustomObject]@{ Name = 'Connect Scripts'; Rows = 1 }
            [PSCustomObject]@{ Name = 'Admin Tools';     Rows = 4 }
            [PSCustomObject]@{ Name = 'Daily Scripts';   Rows = 1 }
        )
    }
    if (-not (Test-Path -LiteralPath $SectionsConfigPath)) { return $default }
    try {
        $cfg = Get-Content -LiteralPath $SectionsConfigPath -Raw | ConvertFrom-Json
        if (-not $cfg.Sections -or @($cfg.Sections).Count -eq 0) { return $default }
        if (-not $cfg.ButtonsPerRow) {
            Add-Member -InputObject $cfg -NotePropertyName ButtonsPerRow -NotePropertyValue 5 -Force
        }
        return $cfg
    } catch { return $default }
}

function Write-SectionsJson {
    param($Config)
    $Config | ConvertTo-Json -Depth 6 |
        Set-Content -LiteralPath $SectionsConfigPath -Encoding UTF8
}

$script:SectionCfg = Read-SectionsJson

function Get-ValidSections {
    return @(@($script:SectionCfg.Sections) | ForEach-Object { "$($_.Name)" })
}

function Get-SectionCapacityTable {
    $w = [int]$script:SectionCfg.ButtonsPerRow
    if ($w -lt 1) { $w = 5 }
    $h = @{}
    foreach ($s in @($script:SectionCfg.Sections)) {
        $r = [int]$s.Rows; if ($r -lt 1) { $r = 1 }
        $h["$($s.Name)"] = $r * $w
    }
    return $h
}

$validSections   = Get-ValidSections
$script:sectionCapacity = Get-SectionCapacityTable

function Get-NormalizedSection {
    param([string]$Value)
    $s = "$Value".Trim()
    foreach ($n in (Get-ValidSections)) { if ($s -eq $n) { return $n } }
    return (Get-ValidSections)[0]
}

function Show-SectionEditor {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Sections'
    $dlg.Size = New-Object System.Drawing.Size(460, 400)
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.StartPosition = 'CenterParent'
    $dlg.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Sections appear in the console top to bottom, in this order. Every section is $($script:SectionCfg.ButtonsPerRow) buttons across."
    $lbl.Location = New-Object System.Drawing.Point(12, 10)
    $lbl.Size = New-Object System.Drawing.Size(420, 34)
    $dlg.Controls.Add($lbl)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(12, 50)
    $grid.Size = New-Object System.Drawing.Size(320, 250)
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.AutoSizeColumnsMode = 'Fill'

    $cN = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $cN.Name = 'Name'; $cN.HeaderText = 'Section name'; $cN.FillWeight = 70
    [void]$grid.Columns.Add($cN)
    $cR = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $cR.Name = 'Rows'; $cR.HeaderText = 'Rows'; $cR.FillWeight = 30
    [void]$grid.Columns.Add($cR)
    $dlg.Controls.Add($grid)

    foreach ($s in @($script:SectionCfg.Sections)) {
        [void]$grid.Rows.Add("$($s.Name)", "$([int]$s.Rows)")
    }

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = 'Add'; $btnAdd.Location = New-Object System.Drawing.Point(342, 50)
    $btnAdd.Size = New-Object System.Drawing.Size(90, 26)
    $btnAdd.Add_Click({ [void]$grid.Rows.Add('New Section', '1') })
    $dlg.Controls.Add($btnAdd)

    $btnDel = New-Object System.Windows.Forms.Button
    $btnDel.Text = 'Remove'; $btnDel.Location = New-Object System.Drawing.Point(342, 82)
    $btnDel.Size = New-Object System.Drawing.Size(90, 26)
    $btnDel.Add_Click({
        if ($grid.Rows.Count -le 1) {
            [void][System.Windows.Forms.MessageBox]::Show(
                'At least one section is required.', 'Sections', 'OK', 'Warning'); return
        }
        if ($grid.SelectedRows.Count -gt 0) { $grid.Rows.Remove($grid.SelectedRows[0]) }
    })
    $dlg.Controls.Add($btnDel)

    $btnUp = New-Object System.Windows.Forms.Button
    $btnUp.Text = 'Move Up'; $btnUp.Location = New-Object System.Drawing.Point(342, 120)
    $btnUp.Size = New-Object System.Drawing.Size(90, 26)
    $btnUp.Add_Click({
        if ($grid.SelectedRows.Count -eq 0) { return }
        $r = $grid.SelectedRows[0]; $i = $r.Index
        if ($i -le 0) { return }
        $grid.Rows.Remove($r); $grid.Rows.Insert($i - 1, $r); $r.Selected = $true
    })
    $dlg.Controls.Add($btnUp)

    $btnDown = New-Object System.Windows.Forms.Button
    $btnDown.Text = 'Move Down'; $btnDown.Location = New-Object System.Drawing.Point(342, 152)
    $btnDown.Size = New-Object System.Drawing.Size(90, 26)
    $btnDown.Add_Click({
        if ($grid.SelectedRows.Count -eq 0) { return }
        $r = $grid.SelectedRows[0]; $i = $r.Index
        if ($i -ge $grid.Rows.Count - 1) { return }
        $grid.Rows.Remove($r); $grid.Rows.Insert($i + 1, $r); $r.Selected = $true
    })
    $dlg.Controls.Add($btnDown)

    $lblW = New-Object System.Windows.Forms.Label
    $lblW.Text = 'Buttons per row:'
    $lblW.Location = New-Object System.Drawing.Point(12, 310)
    $lblW.Size = New-Object System.Drawing.Size(100, 20)
    $dlg.Controls.Add($lblW)

    $numW = New-Object System.Windows.Forms.NumericUpDown
    $numW.Location = New-Object System.Drawing.Point(116, 308)
    $numW.Size = New-Object System.Drawing.Size(50, 22)
    $numW.Minimum = 1; $numW.Maximum = 8
    $numW.Value = [int]$script:SectionCfg.ButtonsPerRow
    $dlg.Controls.Add($numW)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Save'; $btnOk.Location = New-Object System.Drawing.Point(232, 306)
    $btnOk.Size = New-Object System.Drawing.Size(90, 28)
    $dlg.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'; $btnCancel.Location = New-Object System.Drawing.Point(330, 306)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 28)
    $btnCancel.Add_Click({ $dlg.Close() })
    $dlg.Controls.Add($btnCancel)

    $btnOk.Add_Click({
        $list = @()
        $seen = @{}
        foreach ($row in $grid.Rows) {
            $n = "$($row.Cells['Name'].Value)".Trim()
            if (-not $n) { continue }
            if ($seen.ContainsKey($n)) {
                [void][System.Windows.Forms.MessageBox]::Show(
                    "Duplicate section name: '$n'. Names must be unique.",
                    'Sections', 'OK', 'Warning'); return
            }
            $seen[$n] = $true
            $r = 1
            [int]::TryParse("$($row.Cells['Rows'].Value)", [ref]$r) | Out-Null
            if ($r -lt 1) { $r = 1 }
            $list += [PSCustomObject]@{ Name = $n; Rows = $r }
        }
        if ($list.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show(
                'At least one section is required.', 'Sections', 'OK', 'Warning'); return
        }
        $script:SectionCfg = [PSCustomObject]@{
            ButtonsPerRow = [int]$numW.Value
            Sections      = $list
        }
        try {
            Write-SectionsJson -Config $script:SectionCfg
            $dlg.Tag = 'saved'
            $dlg.Close()
        } catch {
            [void][System.Windows.Forms.MessageBox]::Show(
                "Could not save Sections.json:`n$($_.Exception.Message)",
                'Sections', 'OK', 'Error')
        }
    })

    # This dialog was never themed -- it opened in default Windows colours while
    # its parent was dark. It is safe to theme now that DataGridView cell styles
    # are handled properly; before that fix it would have inherited the same
    # dark-text-on-dark-background problem the main grid had.
    foreach ($col in @($grid.Columns)) {
        $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
    }
    $grid.RowTemplate.Height = 26
    Set-ConsoleTheme -Form $dlg
    Add-ToolHelpButton -Form $dlg -ScriptPath $PSCommandPath -Title 'ToolManager' | Out-Null

    [void]$dlg.ShowDialog()
    return ($dlg.Tag -eq 'saved')
}

function Read-ToolsJson {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    if ("$raw".Trim() -eq '') { return @() }
    $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
    return @($parsed)
}

function Write-ToolsJson {
    param([string]$Path, $ToolObjects)
    $tools = @($ToolObjects)
    $json = $tools | ConvertTo-Json -Depth 4
    if ($tools.Count -le 1) { $json = "[`r`n$json`r`n]" }
    Set-Content -LiteralPath $Path -Value $json -Encoding UTF8
}

#region UI ---------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Tool Manager -- $ToolsConfigPath"
$form.Size = New-Object System.Drawing.Size(980, 820)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

# Resizable, because 32 tools in a fixed 380px viewport is the actual problem --
# no amount of contrast fixes having to scroll a list you are trying to compare
# against itself. Anchors below keep the grid growing and the buttons pinned.
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MinimumSize     = New-Object System.Drawing.Size(900, 600)
$form.MaximizeBox     = $true

$lblNote = New-Object System.Windows.Forms.Label
$lblNote.Text = "Use Add Tool / Delete Selected to add or remove rows, edit cells directly to rename or repoint a tool, then click Save. AdminConsole.ps1 reads this file every time it starts. Order (leftmost column) is read-only -- it reflects row position; use Move Up / Move Down to change it."
$lblNote.Location = New-Object System.Drawing.Point(12, 10)
$lblNote.Size = New-Object System.Drawing.Size(860, 34)
$form.Controls.Add($lblNote)

$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Location = New-Object System.Drawing.Point(12, 48)
$dgv.Size = New-Object System.Drawing.Size(860, 380)
$dgv.AllowUserToAddRows = $true
$dgv.AllowUserToDeleteRows = $true
$dgv.AutoGenerateColumns = $false
$dgv.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dgv.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dgv.Font = New-Object System.Drawing.Font("Consolas", 10)
$dgv.RowTemplate.Height = 26
$dgv.AllowUserToResizeRows = $false
$dgv.ColumnHeadersHeight = 28
$dgv.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$dgv.RowHeadersWidth = 24
$dgv.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
$form.Controls.Add($dgv)

$colOrder = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colOrder.Name = 'Order'
$colOrder.HeaderText = 'Order'
$colOrder.ReadOnly = $true
$colOrder.FillWeight = 7
[void]$dgv.Columns.Add($colOrder)
$colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.Name = 'Name'
$colName.HeaderText = 'Name'
$colName.FillWeight = 22
[void]$dgv.Columns.Add($colName)
$colScript = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colScript.Name = 'Script'
$colScript.HeaderText = 'Script'
$colScript.FillWeight = 26
[void]$dgv.Columns.Add($colScript)
$colObjectTypes = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colObjectTypes.Name = 'ObjectTypes'
$colObjectTypes.HeaderText = 'ObjectTypes'
# Was 22, which truncated 'UnifiedGroup, OnpremGroup' to
# 'UnifiedGroup, Onpre...' -- the longest legitimate value in the column, so
# the one case the width had to accommodate. Taken from Order and Tier, which
# hold two digits and one short word.
$colObjectTypes.FillWeight = 25
[void]$dgv.Columns.Add($colObjectTypes)
$colSection = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$colSection.Name = 'Section'
$colSection.HeaderText = 'Section'
$colSection.FillWeight = 14
$colSection.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$validSections | ForEach-Object { [void]$colSection.Items.Add($_) }
[void]$dgv.Columns.Add($colSection)
$colTier = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colTier.Name = 'Tier'
$colTier.HeaderText = 'Tier'
$colTier.FillWeight = 8
[void]$dgv.Columns.Add($colTier)

<#
    Sorting is disabled on every column, and this is a data-safety fix rather
    than a cosmetic one.

    Text columns default to SortMode = Automatic, so clicking a header sorted
    the grid. Order is not stored data -- Set-OrderColumn renumbers it from ROW
    POSITION, and Save writes those numbers. So one click on 'Name' followed by
    Save would silently renumber all 32 tools into alphabetical order and
    rearrange every button in the console. Nothing would have looked wrong at
    the time.

    Move Up / Move Down remain the only way to reorder, which is what the note
    at the top of the window already tells you.
#>
foreach ($col in @($dgv.Columns)) {
    $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
}

function Set-OrderColumn {
    $n = 0
    foreach ($row in $dgv.Rows) {
        if ($row.IsNewRow) { continue }
        $n++
        $row.Cells['Order'].Value = "$n"
    }
}

function Update-ToolsGrid {
    param($ToolObjects)
    $dgv.Rows.Clear()
    foreach ($t in @($ToolObjects)) {
        $typesVal = (@($t.ObjectTypes) -join ', ')
        $sectionVal = Get-NormalizedSection -Value "$($t.Section)"
        [void]$dgv.Rows.Add('', "$($t.Name)", "$($t.Script)", $typesVal, $sectionVal, "$($t.Tier)")
    }
    Set-OrderColumn
}

$dgv.Add_DefaultValuesNeeded({
    param($sender, $e)
    $e.Row.Cells['Section'].Value = (Get-ValidSections)[0]
})

function Get-ToolsFromGrid {
    $result = @()
    $n = 0
    foreach ($row in $dgv.Rows) {
        if ($row.IsNewRow) { continue }
        $name = "$($row.Cells['Name'].Value)".Trim()
        $script = "$($row.Cells['Script'].Value)".Trim()
        if ($name -eq '' -and $script -eq '') { continue }
        $n++
        $types = @(("$($row.Cells['ObjectTypes'].Value)" -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
        $section = Get-NormalizedSection -Value "$($row.Cells['Section'].Value)"
        $result += [PSCustomObject]@{
            Order       = $n
            Name        = $name
            Script      = $script
            ObjectTypes = $types
            Section     = $section
            Tier        = "$($row.Cells['Tier'].Value)".Trim()
        }
    }
    return $result
}

$script:CurrentTools = @(Read-ToolsJson -Path $ToolsConfigPath)
Update-ToolsGrid -ToolObjects $script:CurrentTools

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = "ObjectTypes: comma-separated combination of Mailbox, UnifiedGroup, OnpremGroup. Section: any section defined under Sections... -- defaults to the first section if left blank. Tier is free text for your own documentation -- AdminConsole ignores it."
$lblHint.Location = New-Object System.Drawing.Point(12, 432)
$lblHint.Size = New-Object System.Drawing.Size(860, 20)
$lblHint.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Italic)
$lblHint.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblHint)

$btnAddRow = New-Object System.Windows.Forms.Button
$btnAddRow.Text = "Add Tool"
$btnAddRow.Location = New-Object System.Drawing.Point(12, 460)
$btnAddRow.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnAddRow)
$btnDeleteRow = New-Object System.Windows.Forms.Button
$btnDeleteRow.Text = "Delete Selected"
$btnDeleteRow.Location = New-Object System.Drawing.Point(120, 460)
$btnDeleteRow.Size = New-Object System.Drawing.Size(120, 28)
$form.Controls.Add($btnDeleteRow)

$btnMoveUp = New-Object System.Windows.Forms.Button
$btnMoveUp.Text = "Move Up"
$btnMoveUp.Location = New-Object System.Drawing.Point(12, 496)
$btnMoveUp.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnMoveUp)
$btnMoveDown = New-Object System.Windows.Forms.Button
$btnMoveDown.Text = "Move Down"
$btnMoveDown.Location = New-Object System.Drawing.Point(120, 496)
$btnMoveDown.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnMoveDown)

$btnReload = New-Object System.Windows.Forms.Button
$btnReload.Text = "Reload from Disk"
$btnReload.Location = New-Object System.Drawing.Point(228, 496)
$btnReload.Size = New-Object System.Drawing.Size(120, 28)
$form.Controls.Add($btnReload)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "$($script:CurrentTools.Count) tool(s) loaded from $ToolsConfigPath"
$lblStatus.Location = New-Object System.Drawing.Point(360, 502)
$lblStatus.Size = New-Object System.Drawing.Size(300, 20)
$form.Controls.Add($lblStatus)

$btnSections = New-Object System.Windows.Forms.Button
$btnSections.Text = "Sections..."
$btnSections.Location = New-Object System.Drawing.Point(12, 532)
$btnSections.Size = New-Object System.Drawing.Size(100, 28)
$btnSections.Add_Click({
    if (Show-SectionEditor) {
        # Rebuild the Section dropdown so newly-added sections are selectable
        # immediately, and warn about tools stranded on a deleted section.
        $valid = Get-ValidSections
        $colSection.Items.Clear()
        $valid | ForEach-Object { [void]$colSection.Items.Add($_) }
        $script:sectionCapacity = Get-SectionCapacityTable
        $stranded = @()
        foreach ($row in $dgv.Rows) {
            if ($row.IsNewRow) { continue }
            $cur = "$($row.Cells['Section'].Value)".Trim()
            if ($valid -notcontains $cur) {
                $stranded += "$($row.Cells['Name'].Value) (was '$cur')"
                $row.Cells['Section'].Value = $valid[0]
            }
        }
        $lblStatus.Text = "Sections saved. Reopen the Admin Console to see the new layout."
        if ($stranded.Count -gt 0) {
            [void][System.Windows.Forms.MessageBox]::Show(
                "These tools were in a section that no longer exists and have been moved to '$($valid[0])'. Click Save to keep that change:`n`n$($stranded -join "`n")",
                'Sections Changed', [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }
})
$form.Controls.Add($btnSections)

$btnTheme = New-Object System.Windows.Forms.Button
$btnTheme.Text = "Theme..."
$btnTheme.Location = New-Object System.Drawing.Point(120, 532)
$btnTheme.Size = New-Object System.Drawing.Size(100, 28)
$btnTheme.Add_Click({
    $p = Join-Path $PSScriptRoot 'ThemeManager.ps1'
    if (Test-Path -LiteralPath $p) {
        & $p
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "ThemeManager.ps1 was not found next to this script.",
            'Tool Manager', 'OK', 'Warning') | Out-Null
    }
})
$form.Controls.Add($btnTheme)

#region Settings dialog -------------------------------------------------------
function Show-SettingsEditor {
    <#
        Edits Settings.json: the mail domain used to complete a bare alias, and
        an optional domain controller.

        The window states what the console will ACTUALLY use and where that came
        from, live, as you type. That is here because of how the theme editor
        went wrong: FollowTenant silently overrode ActiveTheme, "Set as active"
        appeared to do nothing, and the fix was to show the resolved answer
        rather than to explain the precedence in a comment. Same problem shape,
        same solution.

        Returns $true if the file was written.
    #>
    $dlg                 = New-Object System.Windows.Forms.Form
    $dlg.Text            = 'Settings'
    $dlg.Size            = New-Object System.Drawing.Size(600, 526)
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false
    $dlg.StartPosition   = 'CenterParent'
    $dlg.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $current = Get-ToolkitSettings

    $lblInfo          = New-Object System.Windows.Forms.Label
    $lblInfo.Text     = 'Set a mail domain and a bare alias completes to a full address, so you can type jsmith instead of jsmith@domain.'
    $lblInfo.Location = New-Object System.Drawing.Point(12, 12)
    $lblInfo.Size     = New-Object System.Drawing.Size(564, 32)
    $dlg.Controls.Add($lblInfo)

    $lblDom          = New-Object System.Windows.Forms.Label
    $lblDom.Text     = 'Mail domain:'
    $lblDom.Location = New-Object System.Drawing.Point(12, 56)
    $lblDom.Size     = New-Object System.Drawing.Size(110, 20)
    $dlg.Controls.Add($lblDom)

    # Editable combo: pick from the tenant's real accepted domains when Exchange
    # Online is connected, type freely when it is not. A transposed character in
    # a hand-typed domain breaks every alias lookup and looks like the tool is
    # broken, so removing the chance to typo is worth the dropdown.
    $cboDom               = New-Object System.Windows.Forms.ComboBox
    $cboDom.DropDownStyle = 'DropDown'
    $cboDom.Location      = New-Object System.Drawing.Point(126, 54)
    $cboDom.Size          = New-Object System.Drawing.Size(280, 22)
    $cboDom.Text          = $current.MailDomain
    $dlg.Controls.Add($cboDom)

    $lblSrc           = New-Object System.Windows.Forms.Label
    $lblSrc.Location  = New-Object System.Drawing.Point(126, 80)
    $lblSrc.Size      = New-Object System.Drawing.Size(410, 16)
    $lblSrc.ForeColor = [System.Drawing.Color]::DimGray
    $dlg.Controls.Add($lblSrc)

    $accepted = @()
    try {
        if (Get-Command -Name 'Get-AcceptedDomain' -ErrorAction SilentlyContinue) {
            $accepted = @(Get-AcceptedDomainList | Where-Object { $_ -notlike '*.onmicrosoft.com' })
        }
    } catch { }
    if (@($accepted).Count -gt 0) {
        $accepted | Sort-Object | ForEach-Object { [void]$cboDom.Items.Add($_) }
        $lblSrc.Text = "$(@($accepted).Count) accepted domain(s) from the connected tenant -- or type your own."
    } else {
        $lblSrc.Text = 'Not connected to Exchange Online, so no list to offer. Type the domain.'
    }

    $lblCloud           = New-Object System.Windows.Forms.Label
    $lblCloud.Text      = 'No domain controller setting: this toolkit is cloud-only and never queries Active Directory.'
    $lblCloud.Location  = New-Object System.Drawing.Point(12, 106)
    $lblCloud.Size      = New-Object System.Drawing.Size(564, 16)
    $lblCloud.ForeColor = [System.Drawing.Color]::DimGray
    $dlg.Controls.Add($lblCloud)

    # ---- folders ------------------------------------------------------------
    $grpPaths          = New-Object System.Windows.Forms.GroupBox
    $grpPaths.Text     = 'Folders  --  blank means beside this toolkit'
    $grpPaths.Location = New-Object System.Drawing.Point(12, 134)
    $grpPaths.Size     = New-Object System.Drawing.Size(564, 100)
    $dlg.Controls.Add($grpPaths)

    $newFolderRow = {
        param($Parent, $LabelText, $Y, $Value, $Description)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $LabelText
        $l.Location = New-Object System.Drawing.Point(12, ($Y + 3))
        $l.Size = New-Object System.Drawing.Size(64, 20)
        $Parent.Controls.Add($l)

        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Location = New-Object System.Drawing.Point(80, $Y)
        $tb.Size = New-Object System.Drawing.Size(378, 22)
        $tb.Text = $Value
        $Parent.Controls.Add($tb)

        $b = New-Object System.Windows.Forms.Button
        $b.Text = 'Browse...'
        $b.Location = New-Object System.Drawing.Point(464, ($Y - 1))
        $b.Size = New-Object System.Drawing.Size(88, 24)
        $b.Add_Click({
            $fb = New-Object System.Windows.Forms.FolderBrowserDialog
            $fb.Description = $Description
            if ($tb.Text -and (Test-Path -LiteralPath $tb.Text)) { $fb.SelectedPath = $tb.Text }
            if ($fb.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $tb.Text = $fb.SelectedPath }
        }.GetNewClosure())
        $Parent.Controls.Add($b)
        return $tb
    }

    $txtReports = & $newFolderRow $grpPaths 'Reports:' 24 $current.ReportsPath 'Where report CSVs are written'
    $txtLogs    = & $newFolderRow $grpPaths 'Logs:'    58 $current.LogPath     'Where CompletedTasks.csv is written'

    # ---- browser profiles ---------------------------------------------------
    $grpBrowser          = New-Object System.Windows.Forms.GroupBox
    $grpBrowser.Text     = 'Edge profiles  --  the profile NAME as shown in Edge, not "Profile 1"'
    $grpBrowser.Location = New-Object System.Drawing.Point(12, 242)
    $grpBrowser.Size     = New-Object System.Drawing.Size(564, 104)
    $dlg.Controls.Add($grpBrowser)

    $existing = Get-BrowserProfileMap

    $lblAdmP          = New-Object System.Windows.Forms.Label
    $lblAdmP.Text     = 'Admin:'
    $lblAdmP.Location = New-Object System.Drawing.Point(12, 27)
    $lblAdmP.Size     = New-Object System.Drawing.Size(64, 20)
    $grpBrowser.Controls.Add($lblAdmP)

    $cboAdmP               = New-Object System.Windows.Forms.ComboBox
    $cboAdmP.DropDownStyle = 'DropDown'
    $cboAdmP.Location      = New-Object System.Drawing.Point(80, 24)
    $cboAdmP.Size          = New-Object System.Drawing.Size(220, 22)
    $cboAdmP.Text          = "$($existing['Admin'])"
    $grpBrowser.Controls.Add($cboAdmP)

    $lblStdP          = New-Object System.Windows.Forms.Label
    $lblStdP.Text     = 'Standard:'
    $lblStdP.Location = New-Object System.Drawing.Point(12, 59)
    $lblStdP.Size     = New-Object System.Drawing.Size(64, 20)
    $grpBrowser.Controls.Add($lblStdP)

    $cboStdP               = New-Object System.Windows.Forms.ComboBox
    $cboStdP.DropDownStyle = 'DropDown'
    $cboStdP.Location      = New-Object System.Drawing.Point(80, 56)
    $cboStdP.Size          = New-Object System.Drawing.Size(220, 22)
    $cboStdP.Text          = "$($existing['Standard'])"
    $grpBrowser.Controls.Add($cboStdP)

    # Offer the profiles Edge actually has. A name that does not exist falls
    # back to the last-used profile at run time -- which usually works, so the
    # mistake surfaces as the wrong account rather than as an error.
    $lblProfWarn           = New-Object System.Windows.Forms.Label
    $lblProfWarn.Location  = New-Object System.Drawing.Point(310, 24)
    $lblProfWarn.Size      = New-Object System.Drawing.Size(244, 66)
    $lblProfWarn.ForeColor = [System.Drawing.Color]::DimGray
    $grpBrowser.Controls.Add($lblProfWarn)

    $edgeProfiles = @()
    try { $edgeProfiles = @(Get-EdgeProfiles | ForEach-Object { "$($_.DisplayName)" } | Where-Object { $_ }) } catch { }
    if (@($edgeProfiles).Count -gt 0) {
        foreach ($cb in @($cboAdmP, $cboStdP)) {
            $edgeProfiles | Sort-Object -Unique | ForEach-Object { [void]$cb.Items.Add($_) }
        }
        $lblProfWarn.Text = "Found on this machine: $((@($edgeProfiles) | Sort-Object -Unique) -join ', ')"
    } else {
        $lblProfWarn.Text = 'No Edge profiles could be read on this machine. Type the names as they appear in Edge.'
    }

    # --- PIM countdown -----------------------------------------------------
    #
    # Off by default and deliberately opt-in. Where admin rights are standing
    # rather than activated on demand, a countdown is a clock that is always
    # wrong -- and a wrong clock beside a connection status is worse than no
    # clock, because it invites belief.
    $grpPim          = New-Object System.Windows.Forms.GroupBox
    $grpPim.Text     = 'Privileged access (PIM)'
    $grpPim.Location = New-Object System.Drawing.Point(12, 372)
    $grpPim.Size     = New-Object System.Drawing.Size(564, 78)

    $chkPim          = New-Object System.Windows.Forms.CheckBox
    $chkPim.Text     = 'Show the activation countdown on the console'
    $chkPim.Location = New-Object System.Drawing.Point(12, 22)
    $chkPim.Size     = New-Object System.Drawing.Size(330, 22)
    $chkPim.Checked  = $(if ($current.PSObject.Properties['PimTimerEnabled']) { [bool]$current.PimTimerEnabled } else { $false })
    $grpPim.Controls.Add($chkPim)

    $lblPimH          = New-Object System.Windows.Forms.Label
    $lblPimH.Text     = 'Activation lasts (hours):'
    $lblPimH.Location = New-Object System.Drawing.Point(352, 24)
    $lblPimH.Size     = New-Object System.Drawing.Size(140, 20)
    $grpPim.Controls.Add($lblPimH)

    $numPimH          = New-Object System.Windows.Forms.NumericUpDown
    $numPimH.Location = New-Object System.Drawing.Point(494, 21)
    $numPimH.Size     = New-Object System.Drawing.Size(56, 22)
    $numPimH.Minimum  = 1
    $numPimH.Maximum  = 24
    $numPimH.Value    = $(if ($current.PSObject.Properties['PimDurationHours'] -and $current.PimDurationHours) { [int]$current.PimDurationHours } else { 4 })
    $grpPim.Controls.Add($numPimH)

    $lblPimNote          = New-Object System.Windows.Forms.Label
    $lblPimNote.Text     = 'The countdown is local -- it starts when you press Activate, and is not read back from the directory.'
    $lblPimNote.Location = New-Object System.Drawing.Point(12, 48)
    $lblPimNote.Size     = New-Object System.Drawing.Size(540, 20)
    $lblPimNote.ForeColor = [System.Drawing.Color]::DimGray
    $grpPim.Controls.Add($lblPimNote)

    $dlg.Controls.Add($grpPim)

    $grpNow          = New-Object System.Windows.Forms.GroupBox
    $grpNow.Text     = 'What the console will use'
    $grpNow.Location = New-Object System.Drawing.Point(12, 354)
    $grpNow.Size     = New-Object System.Drawing.Size(564, 74)
    $dlg.Controls.Add($grpNow)

    $lblResolved          = New-Object System.Windows.Forms.Label
    $lblResolved.Location = New-Object System.Drawing.Point(12, 20)
    $lblResolved.Size     = New-Object System.Drawing.Size(540, 48)
    $grpNow.Controls.Add($lblResolved)

    $updateResolved = {
        $typed = "$($cboDom.Text)".Trim()
        if ($typed) {
            $lblResolved.Text = "Alias completion: jsmith becomes jsmith@$typed" + [Environment]::NewLine +
                                'Source: Settings.json (this dialog), which takes precedence over everything else.'
        } else {
            # Ask for the answer that would apply with this setting empty, so
            # "blank" is a visible outcome rather than an unknown.
            $global:ToolkitSettings = $null
            $auto = $null
            try { $auto = Get-MailDomain } catch { }
            $lblResolved.Text = if ($auto) {
                "Alias completion: jsmith becomes jsmith@$auto" + [Environment]::NewLine +
                'Source: worked out at run time. It will change if you connect to a different tenant.'
            } else {
                'Alias completion: OFF -- nothing can be resolved right now.' + [Environment]::NewLine +
                'Full email addresses will be required until this is set or Exchange Online is connected.'
            }
        }
    }
    $cboDom.Add_TextChanged($updateResolved)
    & $updateResolved

    $btnOK          = New-Object System.Windows.Forms.Button
    $btnOK.Text     = 'Save'
    $btnOK.Location = New-Object System.Drawing.Point(366, 438)
    $btnOK.Size     = New-Object System.Drawing.Size(100, 28)
    $btnOK.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
    $dlg.Controls.Add($btnOK)

    $btnCancel          = New-Object System.Windows.Forms.Button
    $btnCancel.Text     = 'Cancel'
    $btnCancel.Location = New-Object System.Drawing.Point(474, 438)
    $btnCancel.Size     = New-Object System.Drawing.Size(100, 28)
    $btnCancel.Add_Click({ $dlg.Close() })
    $dlg.Controls.Add($btnCancel)

    $btnOK.Add_Click({
        $d = "$($cboDom.Text)".Trim().TrimStart('@').ToLowerInvariant()
        # Blank is legitimate -- it means "go back to working it out".
        if ($d -and $d -notmatch '^[a-z0-9]([a-z0-9-]*[a-z0-9])?(\.[a-z0-9]([a-z0-9-]*[a-z0-9])?)+$') {
            [void][System.Windows.Forms.MessageBox]::Show(
                "'$d' does not look like a domain name.`n`nEnter something like contoso.com, or leave it blank to work it out automatically.",
                'Settings', 'OK', 'Warning')
            return
        }
        try {
            [void](Save-ToolkitSettings -MailDomain $d `
                        -ReportsPath "$($txtReports.Text)".Trim() `
                        -LogPath "$($txtLogs.Text)".Trim() `
                        -PimTimerEnabled ([bool]$chkPim.Checked) `
                        -PimDurationHours ([int]$numPimH.Value) `
                        -BrowserProfiles @{
                            Admin    = "$($cboAdmP.Text)".Trim()
                            Standard = "$($cboStdP.Text)".Trim()
                        })
            $dlg.Tag = 'saved'
            $dlg.Close()
        } catch {
            [void][System.Windows.Forms.MessageBox]::Show(
                "Could not save Settings.json:`n$($_.Exception.Message)",
                'Settings', 'OK', 'Error')
        }
    })

    Set-ConsoleTheme -Form $dlg
    [void]$dlg.ShowDialog()
    return ($dlg.Tag -eq 'saved')
}

$btnSettings          = New-Object System.Windows.Forms.Button
$btnSettings.Text     = 'Settings...'
$btnSettings.Location = New-Object System.Drawing.Point(228, 532)
$btnSettings.Size     = New-Object System.Drawing.Size(100, 28)
$btnSettings.Add_Click({
    if (Show-SettingsEditor) {
        $d = Get-MailDomain
        $lblStatus.Text = if ($d) {
            "Settings saved. A bare alias now completes to @$d. Reopen any open tool to pick it up."
        } else {
            'Settings saved. No mail domain is set, so full email addresses are required.'
        }
    }
})
$form.Controls.Add($btnSettings)
#endregion

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "Save"
$btnSave.Location = New-Object System.Drawing.Point(672, 496)
$btnSave.Size = New-Object System.Drawing.Size(90, 28)
$btnSave.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnSave)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(770, 496)
$btnClose.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClose)

#region Layout and resize behaviour -------------------------------------------
<#
    Positions are recomputed from the CURRENT client size, then anchors are
    applied -- in that order, and the order matters.

    An anchor preserves the distance between a control and the edge it is
    anchored to AS AT THE MOMENT THE ANCHOR IS SET. The buttons were laid out
    for a 640-tall form; anchoring them to Bottom while the form is already 700
    tall would have frozen that old spacing and left a 60px dead strip along
    the bottom. So everything is moved into place first.

    The grid takes all the growth. That is the actual fix for this window: with
    32 tools in a fixed 380px viewport, nine rows were visible at a time, and no
    amount of contrast helps when you are scrolling a list you are trying to
    compare against itself.
#>
$cw = $form.ClientSize.Width
$ch = $form.ClientSize.Height

$rowC = $ch - 40          # Sections... / Theme... / Settings...
$rowB = $rowC - 36        # Move Up / Move Down / Reload
$rowA = $rowB - 36        # Add Tool / Delete Selected
$hintY = $rowA - 24

$lblNote.Width = $cw - 24
$dgv.Width     = $cw - 24
$dgv.Height    = $hintY - 6 - $dgv.Top

$lblHint.Top   = $hintY
$lblHint.Width = $cw - 24

$btnAddRow.Top    = $rowA
$btnDeleteRow.Top = $rowA
$btnMoveUp.Top    = $rowB
$btnMoveDown.Top  = $rowB
$btnReload.Top    = $rowB
$lblStatus.Top    = $rowB + 6
$btnSections.Top  = $rowC
$btnTheme.Top     = $rowC
$btnSettings.Top  = $rowC

# Save and Close stay on the Move row but pinned right, so they keep their
# distance from the edge as the window widens.
$btnSave.Top  = $rowB
$btnClose.Top = $rowB
$btnClose.Left = $cw - 12 - $btnClose.Width
$btnSave.Left  = $btnClose.Left - 8 - $btnSave.Width

$dgv.Anchor       = 'Top,Left,Bottom,Right'
$lblNote.Anchor   = 'Top,Left,Right'
$lblHint.Anchor   = 'Bottom,Left,Right'
foreach ($b in @($btnAddRow, $btnDeleteRow, $btnMoveUp, $btnMoveDown, $btnReload,
                 $btnSections, $btnTheme, $btnSettings)) {
    $b.Anchor = 'Bottom,Left'
}
$lblStatus.Anchor = 'Bottom,Left'
$btnSave.Anchor   = 'Bottom,Right'
$btnClose.Anchor  = 'Bottom,Right'
#endregion

function Move-SelectedRow {
    param([int]$Delta)
    if ($dgv.SelectedRows.Count -eq 0) { return }
    $idx = $dgv.SelectedRows[0].Index
    if ($dgv.Rows[$idx].IsNewRow) { return }
    $targetIdx = $idx + $Delta
    if ($targetIdx -lt 0 -or $targetIdx -ge $dgv.Rows.Count -or $dgv.Rows[$targetIdx].IsNewRow) { return }
    foreach ($colName in @('Name', 'Script', 'ObjectTypes', 'Section', 'Tier')) {
        $a = $dgv.Rows[$idx].Cells[$colName].Value
        $b = $dgv.Rows[$targetIdx].Cells[$colName].Value
        $dgv.Rows[$idx].Cells[$colName].Value = $b
        $dgv.Rows[$targetIdx].Cells[$colName].Value = $a
    }
    Set-OrderColumn
    $dgv.ClearSelection()
    $dgv.Rows[$targetIdx].Selected = $true
    $dgv.CurrentCell = $dgv.Rows[$targetIdx].Cells['Name']
}
$btnMoveUp.Add_Click({ Move-SelectedRow -Delta -1 })
$btnMoveDown.Add_Click({ Move-SelectedRow -Delta 1 })

$btnAddRow.Add_Click({
    if ($dgv.IsCurrentCellInEditMode) { $dgv.EndEdit() }
    $newRowIndex = $dgv.Rows.Count - 1
    $dgv.CurrentCell = $dgv.Rows[$newRowIndex].Cells['Name']
    $dgv.FirstDisplayedScrollingRowIndex = $newRowIndex
    $dgv.BeginEdit($true)
})

$btnDeleteRow.Add_Click({
    if ($dgv.IsCurrentCellInEditMode) { $dgv.EndEdit() }
    if ($dgv.SelectedRows.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Select a row first.", 'Delete Tool', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $idx = $dgv.SelectedRows[0].Index
    if ($dgv.Rows[$idx].IsNewRow) { return }
    $name = "$($dgv.Rows[$idx].Cells['Name'].Value)"
    $label = if ($name -ne '') { "'$name'" } else { 'this row' }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Delete $label from the grid? This isn't written to Tools.json until you click Save.",
        'Confirm Delete', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    $dgv.Rows.RemoveAt($idx)
    Set-OrderColumn
})

$btnReload.Add_Click({
    if ($dgv.IsCurrentCellInEditMode) { $dgv.EndEdit() }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Reload from '$ToolsConfigPath'? Any unsaved edits in the grid will be discarded.",
        'Confirm Reload', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    try {
        $script:CurrentTools = @(Read-ToolsJson -Path $ToolsConfigPath)
        Update-ToolsGrid -ToolObjects $script:CurrentTools
        $lblStatus.Text = "Reloaded $($script:CurrentTools.Count) tool(s) from $ToolsConfigPath at $(Get-Date -Format 'HH:mm:ss')."
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Reload failed: $($_.Exception.Message)",
            'Reload Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$btnSave.Add_Click({
    if ($dgv.IsCurrentCellInEditMode) { $dgv.EndEdit() }
    try {
        $tools = Get-ToolsFromGrid
        $badTypes = @($tools | ForEach-Object { $_.ObjectTypes } | Where-Object { $validObjectTypes -notcontains $_ } | Select-Object -Unique)
        if ($badTypes.Count -gt 0) {
            [void][System.Windows.Forms.MessageBox]::Show(
                "Unrecognized ObjectTypes value(s): $($badTypes -join ', ')`n`nValid values: $($validObjectTypes -join ', ')",
                'Validation Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $blankNames = @($tools | Where-Object { $_.Name -eq '' -or $_.Script -eq '' })
        if ($blankNames.Count -gt 0) {
            [void][System.Windows.Forms.MessageBox]::Show(
                "Every tool needs both a Name and a Script path.",
                'Validation Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $overCapacity = @($script:sectionCapacity.Keys | ForEach-Object {
            $section = $_
            $count = @($tools | Where-Object { $_.Section -eq $section }).Count
            if ($count -gt $script:sectionCapacity[$section]) {
                "$section`: $count tool(s), grid holds $($script:sectionCapacity[$section])"
            }
        } | Where-Object { $_ })
        if ($overCapacity.Count -gt 0) {
            [void][System.Windows.Forms.MessageBox]::Show(
                "This will still save -- AdminConsole's button panels scroll, so nothing is hidden -- but the following section(s) now exceed the grid size they're sized for:`n`n$($overCapacity -join "`n")",
                'Section Over Capacity', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        Write-ToolsJson -Path $ToolsConfigPath -ToolObjects $tools
        $script:CurrentTools = $tools
        Set-OrderColumn
        $lblStatus.Text = "Saved $($tools.Count) tool(s) to $ToolsConfigPath at $(Get-Date -Format 'HH:mm:ss')."
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Save failed: $($_.Exception.Message)",
            'Save Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$btnClose.Add_Click({ $form.Close() })

Set-ConsoleTheme -Form $form
[void]$form.ShowDialog()
#endregion
