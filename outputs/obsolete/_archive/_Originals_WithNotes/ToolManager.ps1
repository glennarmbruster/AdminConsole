#Requires -Version 5.1
<#
.SYNOPSIS
    Editor for Tools.json -- add, remove, rename, reorder, and repoint the
    tools AdminConsole.ps1 displays, with no code changes required.
.DESCRIPTION
    AdminConsole.ps1 builds its button list entirely from Tools.json (see
    Update-ToolButtons there). This utility is the supported way to change
    that file -- it never touches AdminConsole.ps1's code, and
    AdminConsole.ps1 never needs to be edited to add, remove, or reorder a
    tool.

    This is also the mechanism for the tiered-delegation goal: to hand a
    junior admin only Tier 1 tools, copy AdminConsole.ps1 + Common.ps1 +
    only the Tier 1 .ps1 tool scripts + a Tools.json trimmed (via this
    utility) to just those entries. AdminConsole.ps1 doesn't need to know
    tiers exist -- Tools.json only contains what that admin is allowed to
    see, so nothing is hidden by permission checks, only by what's actually
    present in the folder.

    Each row in the grid is one Tools.json entry:
      Order        - read-only, shown for reference only. Derived from the
                     row's top-to-bottom position in the grid at Save time --
                     NOT typed in by you. Use Move Up / Move Down to reorder;
                     the Order column updates automatically to match.
      Name         - button text
      Script       - relative path to the tool .ps1 (must live alongside
                     AdminConsole.ps1 / Common.ps1)
      ObjectTypes  - comma-separated: Mailbox, UnifiedGroup, OnpremGroup --
                     controls which detected object types show this button
      Section      - "Admin Tools", "Daily Scripts", or "Connect Scripts" --
                     which button grid in AdminConsole.ps1 the tool's button
                     renders in. Defaults to "Admin Tools" if left blank or
                     missing (older Tools.json files). AdminConsole sizes
                     Admin Tools for a 5-wide x 4-tall grid (20 buttons),
                     Daily Scripts for 5-wide x 3-tall (15 buttons), and
                     Connect Scripts for a single row, 5-wide x 1-tall (5
                     buttons); Save below warns, but does not block, if any
                     section's row count exceeds its grid's nominal capacity
                     -- AdminConsole still shows every button either way (it
                     scrolls), it just stops being a clean rectangle past
                     that count.
      Tier         - free-text label (Tier1/Tier2/Tier3 or your own scheme),
                     for your own organization/documentation when deciding
                     what to distribute to whom -- AdminConsole ignores it

    REVISED 2026-07-23 (operator request -- third Section value, "Connect
    Scripts"): $validSections and $sectionCapacity both gained a "Connect
    Scripts" entry (capacity 5, matching AdminConsole's new single-row
    panel -- see that file's own header for the pixel math), and
    Get-NormalizedSection now recognizes "Connect Scripts" as its own value
    instead of folding it into the "anything unrecognized defaults to Admin
    Tools" catch-all. No other behavior changed -- the Section column's
    ComboBox picks up the new value automatically since it's populated from
    $validSections, and Save's over-capacity warning already generalizes to
    any number of sections in $sectionCapacity without further changes.

    REVISED 2026-07-21 (operator request -- AdminConsole split into two
    button sections): added the Section column above/described here, a
    ComboBox-style grid column so the value is always exactly "Admin Tools"
    or "Daily Scripts", never a typo. New rows default to "Admin Tools" (see
    the DefaultValuesNeeded handler near the grid setup below) since that's
    where every existing tool currently lives; Save's capacity check is new
    too (Section 21 spirit: warn before the operator finds out the hard way
    that a section grew past what its grid was sized for).

    REVISED 2026-07-11 (field report: "there's no add button"). The grid had
    AllowUserToAddRows = $true, so a new tool could technically be added by
    scrolling to the always-present blank row at the bottom and typing into
    it -- but that's a stock DataGridView affordance with no visual cue, easy
    to miss entirely. Added explicit "Add Tool" (jumps to and opens that
    blank row for editing) and "Delete Selected" (removes the selected row,
    with a confirm prompt) buttons so both actions are discoverable instead
    of implicit. Neither changes what Get-ToolsFromGrid/Save already did --
    they're just explicit entry points into behavior the grid already
    supported.

    REWRITTEN 2026-07-10b (field report: "Tool Manager only shows one script,
    even though Tools.json has 10 tools in it"). The previous version bound
    the grid to a System.Data.DataTable through a
    System.Windows.Forms.BindingSource (DataGridView.DataSource = BindingSource,
    BindingSource.DataSource = DataTable) -- a standard WinForms pattern, but
    one this file is the only place in the whole toolset that used it. Every
    other grid in this project (grdRules in InboxRules.ps1, grdTrace in
    MessageTrace.ps1, grdCal in CalendarPermissions.ps1, etc.) populates an
    UNBOUND DataGridView directly via Columns.Add + Rows.Add, and none of
    those have ever shown this symptom. Rather than keep chasing the DataTable/
    BindingSource path blind (this environment can't run WinForms to
    reproduce it interactively), this file now uses that same
    proven-elsewhere unbound-grid pattern instead: Update-ToolsGrid populates
    it directly from the parsed Tools.json array, and Get-ToolsFromGrid reads
    it back out at Save time. This also removes the Order column's old
    failure mode entirely (a typed-in Order value that could silently
    disagree with the row's actual position) by making Order purely a
    read-only, auto-computed reflection of row position -- Set-OrderColumn
    recalculates it after every load, reload, move, or row add/delete.
.NOTES
    Author  : Glenn Armbruster
    Created : 2026-07-10 (rewritten 2026-07-10b -- see REFACTOR-DESIGN.md)
    Revised : 2026-07-21 1357 (added Section column -- see REVISED note above)
    Revised : 2026-07-23 (added "Connect Scripts" as a third Section value --
              see REVISED note above)
#>
param(
    [string]$ToolsConfigPath = (Join-Path $PSScriptRoot 'Tools.json')
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$validObjectTypes = @('Mailbox', 'UnifiedGroup', 'OnpremGroup')
$validSections = @('Admin Tools', 'Daily Scripts', 'Connect Scripts')
# Matches AdminConsole.ps1's $flowAdminTools (5x4) / $flowDailyScripts (5x3) /
# $flowConnectScripts (5x1) panel sizing -- see the REVISED 2026-07-21 and
# 2026-07-23 notes in AdminConsole.ps1's header for the pixel math these
# numbers come from.
$sectionCapacity = @{ 'Admin Tools' = 20; 'Daily Scripts' = 15; 'Connect Scripts' = 5 }

# Normalizes a raw Section cell value the same permissive way
# AdminConsole.ps1's own Get-ToolSection does -- blank/missing/unrecognized
# all default to 'Admin Tools' so this file and AdminConsole never disagree
# about where a tool with no explicit Section lands.
function Get-NormalizedSection {
    param([string]$Value)
    $s = "$Value".Trim()
    if ($s -eq 'Daily Scripts') { return 'Daily Scripts' }
    if ($s -eq 'Connect Scripts') { return 'Connect Scripts' }
    return 'Admin Tools'
}

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
    # array bracket syntax so Tools.json is always a JSON array, even with 0
    # or 1 entries (Get-ToolDefinitions in AdminConsole.ps1 always wraps its
    # read in @(...), and Read-ToolsJson above now tolerates a bare object
    # too, but keeping the file itself unambiguous is still cleaner).
    if ($tools.Count -le 1) { $json = "[`r`n$json`r`n]" }
    Set-Content -LiteralPath $Path -Value $json -Encoding UTF8
}

#region UI ---------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Tool Manager -- $ToolsConfigPath"
$form.Size = New-Object System.Drawing.Size(900, 596)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblNote = New-Object System.Windows.Forms.Label
$lblNote.Text = "Use Add Tool / Delete Selected to add or remove rows, edit cells directly to rename or repoint a tool, then click Save. AdminConsole.ps1 reads this file every time it starts. Order (leftmost column) is read-only -- it reflects row position; use Move Up / Move Down to change it."
$lblNote.Location = New-Object System.Drawing.Point(12, 10)
$lblNote.Size = New-Object System.Drawing.Size(860, 34)
$form.Controls.Add($lblNote)

# Unbound DataGridView, populated directly via Columns.Add/Rows.Add -- see
# the .DESCRIPTION header above for why this replaced the previous
# DataTable+BindingSource approach.
$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Location = New-Object System.Drawing.Point(12, 48)
$dgv.Size = New-Object System.Drawing.Size(860, 380)
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
$colObjectTypes.FillWeight = 22
[void]$dgv.Columns.Add($colObjectTypes)
$colSection = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$colSection.Name = 'Section'
$colSection.HeaderText = 'Section'
$colSection.FillWeight = 14
$validSections | ForEach-Object { [void]$colSection.Items.Add($_) }
[void]$dgv.Columns.Add($colSection)
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

# Clears and repopulates the grid from a parsed Tools.json array (the
# PSCustomObject[] Read-ToolsJson returns).
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

# DataGridView's built-in blank trailing row (AllowUserToAddRows = $true)
# starts every ComboBox cell unset, which would otherwise let a brand-new
# row reach Get-ToolsFromGrid with a blank Section -- Get-NormalizedSection
# there would still default it to 'Admin Tools' correctly, but pre-filling
# it here means the operator SEES that default immediately while typing,
# instead of it only appearing after Save/Reload.
$dgv.Add_DefaultValuesNeeded({
    param($sender, $e)
    $e.Row.Cells['Section'].Value = 'Admin Tools'
})

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
$lblHint.Text = "ObjectTypes: comma-separated combination of Mailbox, UnifiedGroup, OnpremGroup. Section: Admin Tools (5x4 grid), Daily Scripts (5x3 grid), or Connect Scripts (5x1 grid, single row) in AdminConsole -- defaults to Admin Tools if left blank. Tier is free text for your own documentation -- AdminConsole ignores it."
$lblHint.Location = New-Object System.Drawing.Point(12, 432)
$lblHint.Size = New-Object System.Drawing.Size(860, 20)
$lblHint.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Italic)
$lblHint.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblHint)

# Row 1: row-level edits (Add / Delete). The grid already supports adding a
# row via its built-in blank trailing row (AllowUserToAddRows = $true) and
# deleting via the Delete key (AllowUserToDeleteRows = $true) -- these
# buttons don't add new capability, they just make both actions discoverable
# instead of relying on stock DataGridView affordances with no visual cue.
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

# Row 2: row-order + file-level actions.
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

# Re-reads Tools.json from disk and rebuilds the grid from it, discarding any
# unsaved edits. Useful any time the file on disk might have changed since
# this window opened (a newer Tools.json copied in, another admin's edit,
# etc.) -- also the fastest way to confirm "what does Tools.json on disk
# actually contain right now" if the grid ever looks wrong.
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

# Jumps to and opens the grid's always-present blank trailing row (present
# because AllowUserToAddRows = $true) so a new tool can be typed in
# immediately, instead of the operator having to notice and scroll to it
# themselves.
$btnAddRow.Add_Click({
    if ($dgv.IsCurrentCellInEditMode) { $dgv.EndEdit() }
    $newRowIndex = $dgv.Rows.Count - 1
    $dgv.CurrentCell = $dgv.Rows[$newRowIndex].Cells['Name']
    $dgv.FirstDisplayedScrollingRowIndex = $newRowIndex
    $dgv.BeginEdit($true)
})

# Removes the selected row from the grid (not from Tools.json until Save is
# clicked). Confirms first since there's no undo besides Reload from Disk.
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
        # Capacity is a layout nicety, not a hard rule (AdminConsole's panels
        # still show every button via AutoScroll past this point) -- warn,
        # don't block Save over it.
        $overCapacity = @($sectionCapacity.Keys | ForEach-Object {
            $section = $_
            $count = @($tools | Where-Object { $_.Section -eq $section }).Count
            if ($count -gt $sectionCapacity[$section]) {
                "$section`: $count tool(s), grid holds $($sectionCapacity[$section])"
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

[void]$form.ShowDialog()
#endregion
