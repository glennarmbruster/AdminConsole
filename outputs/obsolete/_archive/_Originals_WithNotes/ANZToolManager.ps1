#Requires -Version 5.1
<#
.SYNOPSIS
    Editor for ANZTools.json -- add, remove, rename, reorder, and repoint the
    tools ANZAdminConsole.ps1 displays, with no code changes required.
.DESCRIPTION
    ANZAdminConsole.ps1 builds its button list entirely from ANZTools.json
    (see Build-ToolButtons there). This utility is the supported way to
    change that file -- it never touches ANZAdminConsole.ps1's code, and
    ANZAdminConsole.ps1 never needs to be edited to add, remove, or reorder
    a tool.

    Built the same way as SRGToolManager.ps1 -- both are ported from the
    original suite's ToolManager.ps1, with the same structural difference:
    NO Section column. ANZAdminConsole.ps1 uses a single "Tools" button
    panel, not the original's Admin Tools / Daily Scripts split (only 3
    tools exist today -- if that changes, add a Section column back the
    same way the original's REVISED 2026-07-21 note describes, and split
    ANZAdminConsole.ps1's single $flowTools panel into two).

    Each row in the grid is one ANZTools.json entry:
      Order        - read-only, shown for reference only. Derived from the
                     row's top-to-bottom position in the grid at Save time --
                     NOT typed in by you. Use Move Up / Move Down to reorder;
                     the Order column updates automatically to match.
      Name         - button text
      Script       - relative path to the tool .ps1 (must live alongside
                     ANZAdminConsole.ps1 / ANZCommon.ps1)
      ObjectTypes  - comma-separated: Mailbox, UnifiedGroup, DistributionGroup
                     -- controls which detected object types show this button
                     (ANZ has no OnpremGroup type -- cloud-only tenant, no
                     synced on-prem AD groups to find)
      Tier         - free-text label (Tier1/Tier2/Tier3 or your own scheme),
                     for your own organization/documentation -- ANZAdminConsole
                     ignores it
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-21 (ported from ToolManager.ps1, same way as
              SRGToolManager.ps1 -- Section column removed)
#>
param(
    [string]$ToolsConfigPath = (Join-Path $PSScriptRoot 'ANZTools.json')
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$validObjectTypes = @('Mailbox', 'UnifiedGroup', 'DistributionGroup')
# Matches ANZAdminConsole.ps1's $flowTools panel sizing (5-wide x 4-tall) --
# see that file's header for the pixel math this number comes from. Same
# "warn, don't block" treatment as the original ToolManager.ps1's
# $sectionCapacity check, just against one panel instead of two.
$toolCapacity = 20

function Read-ToolsJson {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    if ("$raw".Trim() -eq '') { return @() }
    $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
    # @() here matters two ways: if the JSON root is an ARRAY, ConvertFrom-Json
    # already returns an array and @() is a harmless no-op; if the JSON root
    # is a single OBJECT (e.g. someone hand-edited the file down to one entry
    # without keeping the [ ] wrapper), ConvertFrom-Json returns exactly one
    # PSCustomObject, and @() wraps it into a 1-element array so every caller
    # below can treat the result uniformly as a collection either way.
    return @($parsed)
}

function Write-ToolsJson {
    param([string]$Path, $ToolObjects)
    $tools = @($ToolObjects)
    $json = $tools | ConvertTo-Json -Depth 4
    # ConvertTo-Json collapses a single-item array to a bare object -- force
    # array bracket syntax so ANZTools.json is always a JSON array, even with
    # 0 or 1 entries.
    if ($tools.Count -le 1) { $json = "[`r`n$json`r`n]" }
    Set-Content -LiteralPath $Path -Value $json -Encoding UTF8
}

#region UI ---------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "ANZ Tool Manager -- $ToolsConfigPath"
$form.Size = New-Object System.Drawing.Size(900, 560)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblNote = New-Object System.Windows.Forms.Label
$lblNote.Text = "Use Add Tool / Delete Selected to add or remove rows, edit cells directly to rename or repoint a tool, then click Save. ANZAdminConsole.ps1 reads this file every time it starts. Order (leftmost column) is read-only -- it reflects row position; use Move Up / Move Down to change it."
$lblNote.Location = New-Object System.Drawing.Point(12, 10)
$lblNote.Size = New-Object System.Drawing.Size(860, 34)
$form.Controls.Add($lblNote)

$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Location = New-Object System.Drawing.Point(12, 48)
$dgv.Size = New-Object System.Drawing.Size(860, 344)
$dgv.AllowUserToAddRows = $true
$dgv.AllowUserToDeleteRows = $true
$dgv.AutoGenerateColumns = $false
$dgv.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dgv.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dgv.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($dgv)

$colOrder = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colOrder.Name = 'Order'
$colOrder.HeaderText = 'Order'
$colOrder.ReadOnly = $true
$colOrder.FillWeight = 10
[void]$dgv.Columns.Add($colOrder)
$colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.Name = 'Name'
$colName.HeaderText = 'Name'
$colName.FillWeight = 24
[void]$dgv.Columns.Add($colName)
$colScript = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colScript.Name = 'Script'
$colScript.HeaderText = 'Script'
$colScript.FillWeight = 30
[void]$dgv.Columns.Add($colScript)
$colObjectTypes = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colObjectTypes.Name = 'ObjectTypes'
$colObjectTypes.HeaderText = 'ObjectTypes'
$colObjectTypes.FillWeight = 24
[void]$dgv.Columns.Add($colObjectTypes)
$colTier = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colTier.Name = 'Tier'
$colTier.HeaderText = 'Tier'
$colTier.FillWeight = 12
[void]$dgv.Columns.Add($colTier)

# Recomputes the read-only Order column from each row's current
# top-to-bottom position. Called after every load/reload/move so the
# displayed Order always matches what Get-ToolsFromGrid will actually write
# out -- there is no other, independently-typed source of truth for it.
function Set-OrderColumn {
    $n = 0
    foreach ($row in $dgv.Rows) {
        if ($row.IsNewRow) { continue }
        $n++
        $row.Cells['Order'].Value = "$n"
    }
}

# Clears and repopulates the grid from a parsed ANZTools.json array (the
# PSCustomObject[] Read-ToolsJson returns).
function Update-ToolsGrid {
    param($ToolObjects)
    $dgv.Rows.Clear()
    foreach ($t in @($ToolObjects)) {
        $typesVal = (@($t.ObjectTypes) -join ', ')
        [void]$dgv.Rows.Add('', "$($t.Name)", "$($t.Script)", $typesVal, "$($t.Tier)")
    }
    Set-OrderColumn
}

# Reads the grid back into an array of tool objects, in top-to-bottom row
# order, skipping the trailing "type here to add a row" row and any fully
# blank row. Order is assigned from position here too, independent of
# whatever Set-OrderColumn last displayed, so Save is correct even if this
# is ever called without a preceding grid refresh.
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
        $result += [PSCustomObject]@{
            Order       = $n
            Name        = $name
            Script      = $script
            ObjectTypes = $types
            Tier        = "$($row.Cells['Tier'].Value)".Trim()
        }
    }
    return $result
}

$script:CurrentTools = @(Read-ToolsJson -Path $ToolsConfigPath)
Update-ToolsGrid -ToolObjects $script:CurrentTools

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = "ObjectTypes: comma-separated combination of Mailbox, UnifiedGroup, DistributionGroup. Tier is free text for your own documentation -- ANZAdminConsole ignores it."
$lblHint.Location = New-Object System.Drawing.Point(12, 396)
$lblHint.Size = New-Object System.Drawing.Size(860, 20)
$lblHint.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Italic)
$lblHint.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblHint)

$btnAddRow = New-Object System.Windows.Forms.Button
$btnAddRow.Text = "Add Tool"
$btnAddRow.Location = New-Object System.Drawing.Point(12, 424)
$btnAddRow.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnAddRow)
$btnDeleteRow = New-Object System.Windows.Forms.Button
$btnDeleteRow.Text = "Delete Selected"
$btnDeleteRow.Location = New-Object System.Drawing.Point(120, 424)
$btnDeleteRow.Size = New-Object System.Drawing.Size(120, 28)
$form.Controls.Add($btnDeleteRow)

$btnMoveUp = New-Object System.Windows.Forms.Button
$btnMoveUp.Text = "Move Up"
$btnMoveUp.Location = New-Object System.Drawing.Point(12, 460)
$btnMoveUp.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnMoveUp)
$btnMoveDown = New-Object System.Windows.Forms.Button
$btnMoveDown.Text = "Move Down"
$btnMoveDown.Location = New-Object System.Drawing.Point(120, 460)
$btnMoveDown.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnMoveDown)

$btnReload = New-Object System.Windows.Forms.Button
$btnReload.Text = "Reload from Disk"
$btnReload.Location = New-Object System.Drawing.Point(228, 460)
$btnReload.Size = New-Object System.Drawing.Size(120, 28)
$form.Controls.Add($btnReload)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "$($script:CurrentTools.Count) tool(s) loaded from $ToolsConfigPath"
$lblStatus.Location = New-Object System.Drawing.Point(360, 466)
$lblStatus.Size = New-Object System.Drawing.Size(300, 20)
$form.Controls.Add($lblStatus)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "Save"
$btnSave.Location = New-Object System.Drawing.Point(672, 460)
$btnSave.Size = New-Object System.Drawing.Size(90, 28)
$btnSave.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnSave)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(770, 460)
$btnClose.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClose)

# Swaps the selected row's cell VALUES with the adjacent row's, rather than
# trying to physically reorder rows in the grid (DataGridView has no direct
# "move row" API for an unbound grid) -- the net visual effect is identical.
# Set-OrderColumn afterward keeps the read-only Order column in sync.
function Move-SelectedRow {
    param([int]$Delta)
    if ($dgv.SelectedRows.Count -eq 0) { return }
    $idx = $dgv.SelectedRows[0].Index
    if ($dgv.Rows[$idx].IsNewRow) { return }
    $targetIdx = $idx + $Delta
    if ($targetIdx -lt 0 -or $targetIdx -ge $dgv.Rows.Count -or $dgv.Rows[$targetIdx].IsNewRow) { return }
    foreach ($colName in @('Name', 'Script', 'ObjectTypes', 'Tier')) {
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
        "Delete $label from the grid? This isn't written to ANZTools.json until you click Save.",
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
        # Capacity is a layout nicety, not a hard rule (ANZAdminConsole's
        # panel still shows every button via AutoScroll past this point) --
        # warn, don't block Save over it.
        if ($tools.Count -gt $toolCapacity) {
            [void][System.Windows.Forms.MessageBox]::Show(
                "This will still save -- ANZAdminConsole's button panel scrolls, so nothing is hidden -- but $($tools.Count) tool(s) now exceeds the $toolCapacity the panel is sized for.",
                'Over Capacity', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
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

[void]$form.ShowDialog()
#endregion
