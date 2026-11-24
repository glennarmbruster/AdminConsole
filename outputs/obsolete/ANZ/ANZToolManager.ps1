#Requires -Version 5.1
<#
    ANZToolManager.ps1
    Editor for ANZTools.json -- add, remove, rename, reorder, and repoint
    the tools ANZAdminConsole.ps1 displays, with no code changes required.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

param(
    [string]$ToolsConfigPath = (Join-Path $PSScriptRoot 'ANZTools.json')
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$validObjectTypes = @('Mailbox', 'UnifiedGroup', 'DistributionGroup')
$toolCapacity = 20

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
        [void]$dgv.Rows.Add('', "$($t.Name)", "$($t.Script)", $typesVal, "$($t.Tier)")
    }
    Set-OrderColumn
}

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
