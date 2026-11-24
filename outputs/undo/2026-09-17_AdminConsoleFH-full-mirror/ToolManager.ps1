#Requires -Version 5.1
<#
    ToolManager.ps1
    Editor for Tools.json -- add, remove, rename, reorder, and repoint the
    tools AdminConsole.ps1 displays, with no code changes required.

    Author   : Glenn Armbruster
    Revised  : 2026-09-16

    2026-09-16  The Group column is PERSONAL. It writes to Config\Users\<you>.json
                and no longer to Tools.json, so two engineers sharing a folder
                file tools their own way without changing each other's console.
                A Group left in Tools.json is dropped on the next save here.
    2026-09-16  The Sections editor seeds from this operator's own layout, so
                saving cannot overwrite it with the shared file.
#>
# TOOLKIT-HELP-BEGIN
# Decides which buttons the console shows, and what each one runs.
#
# Every row is one button. Edit cells directly, then Save. AdminConsole.ps1
# re-reads this file each time it starts, so changes appear on the next launch.
#
# THE COLUMNS
#   Name         The button caption.
#   Script       The .ps1 it runs, relative to the toolkit folder.
#   ObjectTypes  Which looked-up object the tool applies to -- Mailbox,
#                UnifiedGroup, OnpremGroup, comma-separated. LEAVE BLANK for a
#                tool that is not about one object; blank means "always
#                applies", not "never".
#   Section      Which panel it appears in. Sections... defines them.
#   Group        Optional. Type the same word on several tools and they collapse
#                behind one button that opens a chooser -- 'Rooms' on the room
#                tools, for instance. Blank means the tool stands alone.
#   Tier         Free text for your own notes. The console ignores it.
#   Own window   Runs the tool in its own PowerShell process. REQUIRED for
#                anything touching on-premises Exchange -- see below.
#
# OWN WINDOW IS NOT A PREFERENCE. On-premises Exchange and Exchange Online
# export the same cmdlet names, and in one process the last connection wins. A
# tool without it can act on the wrong directory with no error at all. If in
# doubt, tick it.
#
# Order is row position -- use Move Up / Move Down, not the Order column.
#
# The other buttons: Sections... edits the panels and how many buttons wide
# they are. Theme... picks colours. Settings... holds YOUR paths and email,
# which are per-engineer and not shipped.
# TOOLKIT-HELP-END

param(
    [string]$ToolsConfigPath = ''
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
#
# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost: the window flashes up, work may have started, and
# the operator has seen a tool they were not meant to be in.
if (-not (Test-ConsoleLaunch)) { return }

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

# WHAT THIS OPERATOR ACTUALLY SEES, NOT WHAT THE SHARED FILE SAYS. The editor
# saves to the personal profile when the operator answers 'just me', and the
# save replaces their section list WHOLE -- so seeding the grid from
# Sections.json alone meant opening the editor and saving showed somebody
# else's sections and silently discarded your own rows, accents and ordering.
# The profile wins where it has an opinion; Sections.json is the fallback.
try {
    $uSec = Get-UserLayoutProperty -Name 'Sections'
    if ($uSec -and @($uSec).Count -gt 0) {
        Add-Member -InputObject $script:SectionCfg -NotePropertyName Sections -NotePropertyValue @($uSec) -Force
    }
    $uPer = Get-UserLayoutProperty -Name 'ButtonsPerRow'
    if ($uPer -and [int]$uPer -gt 0) {
        Add-Member -InputObject $script:SectionCfg -NotePropertyName ButtonsPerRow -NotePropertyValue ([int]$uPer) -Force
    }
} catch { }

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
    $dlg.Size = New-Object System.Drawing.Size(640, 470)
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.StartPosition = 'CenterParent'
    $dlg.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $lbl = New-Object System.Windows.Forms.Label
    # The button count used to be baked into this sentence. It was written once
    # at dialog creation and never updated, so the moment the spinner below was
    # touched there were two different numbers on screen claiming to be the
    # same setting. The live one is in the preview at the bottom now.
    $lbl.Text = "Sections appear in the console top to bottom, in this order. Every section is the same number of buttons across."
    $lbl.Location = New-Object System.Drawing.Point(12, 10)
    $lbl.Size = New-Object System.Drawing.Size(600, 34)
    $dlg.Controls.Add($lbl)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(12, 50)
    $grid.Size = New-Object System.Drawing.Size(490, 250)
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.AutoSizeColumnsMode = 'Fill'

    $cN = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $cN.Name = 'Name'; $cN.HeaderText = 'Section name'; $cN.FillWeight = 56
    [void]$grid.Columns.Add($cN)
    $cR = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $cR.Name = 'Rows'; $cR.HeaderText = 'Rows'; $cR.FillWeight = 14
    [void]$grid.Columns.Add($cR)
    # ACCENT IS A COLUMN HERE BECAUSE THIS DIALOG REBUILDS Sections.json. A
    # key the dialog does not know about is a key the next Save deletes -- the
    # same rule Save-ToolkitSettings documents -- so a section colour added by
    # hand would survive exactly until somebody reordered a section. Optional;
    # blank means no accent. Hex only (#RGB or #RRGGBB); the console validates
    # again on load and ignores anything else.
    $cA = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $cA.Name = 'Accent'; $cA.HeaderText = 'Accent'; $cA.FillWeight = 30
    [void]$grid.Columns.Add($cA)
    $dlg.Controls.Add($grid)

    # THE ACCENT CELL SHOWS ITS COLOUR. Typing a hex value was the whole
    # interface, and a hex value is not a colour anyone can see. The cell is
    # painted in the accent with black or white text by luminance, so the
    # grid IS the preview; the hex stays as the value underneath, editable by
    # hand for anyone who prefers it. Blank cells paint like any other.
    $grid.Add_CellFormatting({
        param($s, $e)
        try {
            if ($e.ColumnIndex -lt 0 -or $s.Columns[$e.ColumnIndex].Name -ne 'Accent') { return }
            $hex = "$($e.Value)".Trim()
            if ($hex -notmatch '^#([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$') { return }
            $c = ConvertTo-ThemeColor $hex
            if (-not $c) { return }
            $e.CellStyle.BackColor = $c
            $e.CellStyle.SelectionBackColor = $c
            $lum = (0.299 * $c.R + 0.587 * $c.G + 0.114 * $c.B) / 255
            $fg = if ($lum -lt 0.55) { [System.Drawing.Color]::White } else { [System.Drawing.Color]::Black }
            $e.CellStyle.ForeColor = $fg
            $e.CellStyle.SelectionForeColor = $fg
        } catch { }
    })

    # A colour picker for the selected section: the button, or double-click
    # the Accent cell. The system ColorDialog, full palette open, seeded with
    # the current accent so "a bit darker" starts from the right place.
    $pickAccent = {
        if ($grid.SelectedRows.Count -eq 0) { return }
        $row = $grid.SelectedRows[0]
        $cd = New-Object System.Windows.Forms.ColorDialog
        $cd.FullOpen = $true
        $cd.AnyColor = $true
        $cur = $null
        try { $cur = ConvertTo-ThemeColor "$($row.Cells['Accent'].Value)".Trim() } catch { }
        if ($cur) { $cd.Color = $cur }
        if ($cd.ShowDialog($dlg) -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $row.Cells['Accent'].Value = '#{0:X2}{1:X2}{2:X2}' -f $cd.Color.R, $cd.Color.G, $cd.Color.B
        $grid.InvalidateRow($row.Index)
    }
    $grid.Add_CellDoubleClick({
        param($s, $e)
        try {
            if ($e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) { return }
            if ($s.Columns[$e.ColumnIndex].Name -ne 'Accent') { return }
            $s.ClearSelection(); $s.Rows[$e.RowIndex].Selected = $true
            & $pickAccent
        } catch { }
    })

    foreach ($s in @($script:SectionCfg.Sections)) {
        $acc = ''
        try { if ($s.PSObject.Properties['Accent']) { $acc = "$($s.Accent)".Trim() } } catch { }
        [void]$grid.Rows.Add("$($s.Name)", "$([int]$s.Rows)", $acc)
    }

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = 'Add'; $btnAdd.Location = New-Object System.Drawing.Point(512, 50)
    $btnAdd.Size = New-Object System.Drawing.Size(90, 26)
    $btnAdd.Add_Click({ [void]$grid.Rows.Add('New Section', '1', '') })
    $dlg.Controls.Add($btnAdd)

    $btnDel = New-Object System.Windows.Forms.Button
    $btnDel.Text = 'Remove'; $btnDel.Location = New-Object System.Drawing.Point(512, 82)
    $btnDel.Size = New-Object System.Drawing.Size(90, 26)
    $btnDel.Add_Click({
        if ($grid.Rows.Count -le 1) {
            $null = Show-ToolDialog(
                'At least one section is required.', 'Sections', 'OK', 'Warning'); return
        }
        if ($grid.SelectedRows.Count -gt 0) { $grid.Rows.Remove($grid.SelectedRows[0]) }
    })
    $dlg.Controls.Add($btnDel)

    $btnUp = New-Object System.Windows.Forms.Button
    $btnUp.Text = 'Move Up'; $btnUp.Location = New-Object System.Drawing.Point(512, 120)
    $btnUp.Size = New-Object System.Drawing.Size(90, 26)
    $btnUp.Add_Click({
        if ($grid.SelectedRows.Count -eq 0) { return }
        $r = $grid.SelectedRows[0]; $i = $r.Index
        if ($i -le 0) { return }
        $grid.Rows.Remove($r); $grid.Rows.Insert($i - 1, $r); $r.Selected = $true
    })
    $dlg.Controls.Add($btnUp)

    $btnDown = New-Object System.Windows.Forms.Button
    $btnDown.Text = 'Move Down'; $btnDown.Location = New-Object System.Drawing.Point(512, 152)
    $btnDown.Size = New-Object System.Drawing.Size(90, 26)
    $btnDown.Add_Click({
        if ($grid.SelectedRows.Count -eq 0) { return }
        $r = $grid.SelectedRows[0]; $i = $r.Index
        if ($i -ge $grid.Rows.Count - 1) { return }
        $grid.Rows.Remove($r); $grid.Rows.Insert($i + 1, $r); $r.Selected = $true
    })
    $dlg.Controls.Add($btnDown)

    $btnPick = New-Object System.Windows.Forms.Button
    $btnPick.Text = 'Pick Accent...'; $btnPick.Location = New-Object System.Drawing.Point(512, 190)
    $btnPick.Size = New-Object System.Drawing.Size(100, 26)
    $btnPick.Add_Click({ & $pickAccent })
    $dlg.Controls.Add($btnPick)

    $btnNoAcc = New-Object System.Windows.Forms.Button
    $btnNoAcc.Text = 'No Accent'; $btnNoAcc.Location = New-Object System.Drawing.Point(512, 222)
    $btnNoAcc.Size = New-Object System.Drawing.Size(100, 26)
    $btnNoAcc.Add_Click({
        if ($grid.SelectedRows.Count -eq 0) { return }
        $grid.SelectedRows[0].Cells['Accent'].Value = ''
        $grid.InvalidateRow($grid.SelectedRows[0].Index)
    })
    $dlg.Controls.Add($btnNoAcc)

    $lblW = New-Object System.Windows.Forms.Label
    $lblW.Text = 'Buttons per row:'
    $lblW.Location = New-Object System.Drawing.Point(12, 310)
    $lblW.Size = New-Object System.Drawing.Size(100, 20)
    $dlg.Controls.Add($lblW)

    $numW = New-Object System.Windows.Forms.NumericUpDown
    $numW.Location = New-Object System.Drawing.Point(116, 308)
    $numW.Size = New-Object System.Drawing.Size(50, 22)
    # 12, matching the console's own clamp in Get-SectionDefinitions. These
    # were 8 here and 10 there: the spinner was the lower of the two and the
    # only one visible, so eight looked like the limit while the console would
    # happily have drawn ten. One number, one limit.
    $numW.Minimum = 1; $numW.Maximum = 12
    $numW.Value = [int]$script:SectionCfg.ButtonsPerRow
    $dlg.Controls.Add($numW)

    # Live size preview.
    #
    # Buttons per row used to be a number you changed, saved, relaunched the
    # console, and then judged -- and if it was wrong you did the whole loop
    # again. The console sizes its buttons by measuring the captions now, so
    # the resulting width is knowable HERE, before saving.
    #
    # It is measured with $dlg.Font, which is the same Microsoft Sans Serif 9
    # the console uses. If those two ever diverge this preview becomes fiction,
    # so they are both left alone deliberately.
    $lblPreview = New-Object System.Windows.Forms.Label
    $lblPreview.Location = New-Object System.Drawing.Point(12, 342)
    $lblPreview.Size = New-Object System.Drawing.Size(600, 48)
    $dlg.Controls.Add($lblPreview)

    # Read straight from Tools.json rather than from whatever the grid behind
    # this dialog holds. The console measures the FILE, so previewing anything
    # else would predict a width the console is not going to use.
    $script:PreviewTools = @()
    try {
        $script:PreviewTools = @(Get-Content -LiteralPath $ToolsConfigPath -Raw | ConvertFrom-Json)
    } catch {
        $script:PreviewTools = @()
    }

    $updatePreview = {
        try {
            $per  = [int]$numW.Value
            $caps = @(@($script:PreviewTools) | ForEach-Object { "$($_.Name)" } |
                        Where-Object { "$_".Trim() })
            if ($caps.Count -eq 0) {
                $lblPreview.Text = 'Tools.json could not be read, so the size cannot be previewed.'
                return
            }

            # Sections come from the GRID, not from the saved file. The preview
            # has to describe what Save is about to write -- previewing what is
            # already on disk would mean the number stops moving the moment you
            # change a row count, which reads as the preview being broken.
            $secs = @()
            foreach ($row in $grid.Rows) {
                $n = "$($row.Cells['Name'].Value)".Trim()
                if (-not $n) { continue }
                $r = 1
                [int]::TryParse("$($row.Cells['Rows'].Value)", [ref]$r) | Out-Null
                if ($r -lt 1) { $r = 1 }
                $secs += [PSCustomObject]@{ Name = $n; Rows = $r }
            }

            # An ESTIMATE of the scrollbar case, and knowingly a pessimistic
            # one. The console decides by asking the panels after tenant
            # capabilities have hidden what they are going to hide; this counts
            # every tool in the file, so it can predict an overflow the console
            # will not actually have. Wrong in the safe direction: it warns
            # about a wider window than you may get, never a narrower one.
            $over = Test-ConsoleSectionOverflow -Sections      $secs `
                                                -Tools         $script:PreviewTools `
                                                -ButtonsPerRow $per

            $bw = Get-ConsoleButtonWidth -Captions     $caps `
                                         -Font         $dlg.Font `
                                         -ButtonHeight $script:ConsoleBtnHeight
            $pw = Get-ConsolePanelWidth -ButtonWidth      $bw `
                                        -ButtonPad        $script:ConsoleBtnPad `
                                        -ButtonsPerRow    $per `
                                        -ReserveScrollBar:$over
            $formW   = $pw + 24
            $screenW = [System.Windows.Forms.Screen]::FromControl($dlg).WorkingArea.Width

            $t = "$per across, buttons $bw px wide. Console window: $formW px." +
                 " This screen fits $screenW px."
            if ($over) {
                $t += "  A section holds more tools than its rows allow, so every" +
                      " panel carries scrollbar width."
            }
            if ($formW -gt $screenW) {
                $t += "  Too wide for this screen -- the console will scroll sideways."
            }
            $lblPreview.Text = $t
        } catch {
            $lblPreview.Text = ''
        }
    }
    $numW.Add_ValueChanged($updatePreview)
    $grid.Add_CellValueChanged($updatePreview)
    & $updatePreview

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Save'; $btnOk.Location = New-Object System.Drawing.Point(412, 306)
    $btnOk.Size = New-Object System.Drawing.Size(90, 28)
    $dlg.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'; $btnCancel.Location = New-Object System.Drawing.Point(510, 306)
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
                $null = Show-ToolDialog(
                    "Duplicate section name: '$n'. Names must be unique.",
                    'Sections', 'OK', 'Warning'); return
            }
            $seen[$n] = $true
            $r = 1
            [int]::TryParse("$($row.Cells['Rows'].Value)", [ref]$r) | Out-Null
            if ($r -lt 1) { $r = 1 }
            $acc = "$($row.Cells['Accent'].Value)".Trim()
            if ($acc -and $acc -notmatch '^#([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$') {
                $null = Show-ToolDialog(
                    "The accent for '$n' is not a hex colour. Use #RGB or #RRGGBB, or leave it blank.",
                    'Sections', 'OK', 'Warning'); return
            }
            $entry = [ordered]@{ Name = $n; Rows = $r }
            if ($acc) { $entry['Accent'] = $acc }
            $list += [PSCustomObject]$entry
        }
        if ($list.Count -eq 0) {
            $null = Show-ToolDialog(
                'At least one section is required.', 'Sections', 'OK', 'Warning'); return
        }
        # ButtonWidth is carried through rather than rebuilt.
        #
        # It is not editable in this dialog, so a save that dropped it would
        # wipe a deliberate override the next time anyone merely reordered a
        # section -- and the symptom would be "the console resized itself for
        # no reason", which is not a clue that leads anywhere.
        $new = [ordered]@{ ButtonsPerRow = [int]$numW.Value }
        if ($script:SectionCfg.PSObject.Properties['ButtonWidth']) {
            $new['ButtonWidth'] = $script:SectionCfg.ButtonWidth
        }
        $new['Sections'] = $list
        $script:SectionCfg = [PSCustomObject]$new

        # WHOSE SECTIONS ARE THESE? Sections.json is shared, so editing it here
        # re-shapes the console for everyone who uses the folder. From
        # 2026-09-15 the layout can be per person instead: the same edits go
        # into this operator's profile and nobody else's console moves.
        #
        # ASKED ONCE PER PERSON WHO HAS ARRANGED SOMETHING -- the same rule
        # as Save Arrangement and Theme Manager, so the three cannot disagree.
        #
        # NOT Test-UserProfile ANY MORE. 'Do they have a profile file' and
        # 'have they arranged anything' were the same question only while
        # arranging was the only thing that created a profile. Picking a
        # default view creates one now, and so does nudging A+ inside a view --
        # so somebody who had done neither of those deliberate acts of tidying
        # would have had their FIRST EVER section save go silently to 'just
        # me', never offered 'everyone', never told which happened.
        $mine = $false
        try {
            $mine = ($null -ne (Get-UserLayoutProperty -Name 'Sections')) -or
                    ($null -ne (Get-UserLayoutProperty -Name 'Placement'))
        } catch { }
        if (-not $mine) {
            $a = Show-ToolDialog(
                ("Save these sections for YOU, or for everyone who uses this toolkit folder?" + "`r`n`r`n" +
                 "YES     just me" + "`r`n" +
                 "            (a profile is created; other people's consoles do not change)" + "`r`n" +
                 "NO      everyone" + "`r`n" +
                 "            (writes Sections.json, the shared file)" + "`r`n" +
                 "CANCEL  go back to the grid"),
                'Save sections', 'YesNoCancel', 'Question')
            if ($a -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
            $mine = ($a -eq [System.Windows.Forms.DialogResult]::Yes)
        }

        try {
            if ($mine) {
                # Placement, theme and parking are carried across untouched:
                # Save-UserLayout replaces the block whole, and a section edit
                # must not quietly undo an arrangement or a parked list.
                $layout = Get-UserLayoutOwn
                if (-not $layout) { $layout = [PSCustomObject]@{} }
                Add-Member -InputObject $layout -NotePropertyName 'Sections' -NotePropertyValue $list -Force
                Add-Member -InputObject $layout -NotePropertyName 'ButtonsPerRow' -NotePropertyValue ([int]$numW.Value) -Force
                $up = Save-UserLayout -Layout $layout
                $null = Show-ToolDialog(
                    ("Saved as your layout." + [Environment]::NewLine + [Environment]::NewLine + "  $up" + [Environment]::NewLine + [Environment]::NewLine +
                     "Sections.json is untouched, so nobody else's console changed. Reopen the Admin Console to see it."),
                    'Sections', 'OK', 'Information')
            } else {
                Write-SectionsJson -Config $script:SectionCfg
            }
            $dlg.Tag = 'saved'
            $dlg.Close()
        } catch {
            $null = Show-ToolDialog(
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
# 1180 wide, up from 980.
#
# The grid gained a Group column and the status line reports a full path --
# 'E:\scripts\m365\adminconsole\Config\Tools.json' is 44 characters before the
# count in front of it. At 980 the columns were squeezed to ellipses and the
# path ran out of label.
$form.Size = New-Object System.Drawing.Size(1180, 820)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

# Resizable, because 32 tools in a fixed 380px viewport is the actual problem --
# no amount of contrast fixes having to scroll a list you are trying to compare
# against itself. Anchors below keep the grid growing and the buttons pinned.
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MinimumSize     = New-Object System.Drawing.Size(900, 600)
$form.MaximizeBox     = $true

$lblNote = New-Object System.Windows.Forms.Label
$lblNote.Text = "Use Add Tool / Delete Selected to add or remove rows, edit cells directly to rename or repoint a tool, then click Save. AdminConsole.ps1 reads this file every time it starts. Order (leftmost column) is read-only and is what the console uses; use Move Up / Move Down to change it. Click any column header to sort the view -- the saved order is not affected; click Order to return to it."
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
# Group folds several tools behind ONE console button.
#
# Type the same name on every tool that belongs together -- 'Rooms' on the five
# room tools -- and the console shows a single Rooms button that opens a chooser
# instead of five buttons on the front page. Blank means the tool stands alone,
# which is the default and what almost everything should stay as.
#
# The button takes the position of the FIRST member in Order, so the group lands
# where you would expect rather than wherever the alphabet puts it.
$colGroup = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colGroup.Name = 'Group'
$colGroup.HeaderText = 'Group'
$colGroup.FillWeight = 12
[void]$dgv.Columns.Add($colGroup)

# TIER IS GONE. It was free text nobody read -- ToolManager was the only file
# that mentioned it, and the help text said outright that the console ignores
# it. A column that exists but does nothing invites people to fill it in and
# then wonder why nothing happened.
#
# Existing Tier values in Tools.json are NOT carried through, so they disappear
# on the next Save. That is the intended outcome: a field with no reader is
# noise in a file that is otherwise all meaning. If a per-tool label is wanted
# later it should come back with something actually consuming it.

# THERE IS NO 'NEW WINDOW' COLUMN, DELIBERATELY.
#
# It was the original control for running a tool in its own PowerShell process.
# On 2026-08-06 an on-premises tool ran with it unticked, loaded into a console
# already connected to Exchange Online, and every write went to the cloud --
# so the requirement was moved INTO the scripts as TOOLKIT-REQUIRES: OwnProcess
# and the console obeys that.
#
# The checkbox survived as "I would like this in its own window" and was never
# used again. All 55 rows sat unticked for months with no symptom, because the
# declarations were doing the work. It looked like the control for something it
# no longer controlled.
#
# A tool that needs its own process says so in its own header. Tools.json now
# describes only what the console SHOWS; the script describes what it NEEDS.

<#
    SORTING IS A VIEW. THE ORDER THAT IS SAVED NEVER COMES FROM ROW POSITION.

    Until 2026-09-12 sorting was disabled outright, because Order was
    renumbered from row position and Save wrote those numbers -- one click on
    'Name' followed by Save would have alphabetised every button in the
    console with nothing looking wrong at the time.

    Now every row carries its place in the SAVED order on .Tag. Clicking a
    header re-arranges the rows on screen (the grid's own sort is off --
    SortMode is Programmatic -- so the arrangement is done here, where the
    rule is visible); Order, Save and the JSON all read .Tag and never the
    row index. Clicking 'Order' puts the grid back in saved order, which is
    the default view: the sections in the order Sections.json declares them,
    then the tools in the order they sit on the console.

    Move Up / Move Down only make sense against the saved order, so they are
    enabled only in that view. Everything else -- edit, add, delete, save --
    works in any view.
#>
foreach ($col in @($dgv.Columns)) {
    $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Programmatic
}
$script:GridSort = [PSCustomObject]@{ Column = 'Order'; Descending = $false }

function Test-GridNaturalOrder {
    return (($script:GridSort.Column -eq 'Order') -and (-not $script:GridSort.Descending))
}

function Get-GridRowsInSavedOrder {
    <# The data rows, in the order Save will write them: by .Tag, rows that
       have no Tag yet (typed into the new-row line since the last renumber)
       at the end, in the order they appear. #>
    $rows = @($dgv.Rows | Where-Object { -not $_.IsNewRow })
    return @($rows | Sort-Object -Property @{ Expression = { if ($null -ne $_.Tag) { [int]$_.Tag } else { [int]::MaxValue } } },
                                            @{ Expression = { $_.Index } })
}

function Set-OrderColumn {
    <# Renumber the SAVED order 1..n and show it. From row position only when
       the grid is in saved-order view; otherwise from the Tags, so a renumber
       in a sorted view cannot bake the sort into the file. #>
    $seq = if (Test-GridNaturalOrder) { @($dgv.Rows | Where-Object { -not $_.IsNewRow }) } else { Get-GridRowsInSavedOrder }
    $n = 0
    foreach ($row in $seq) {
        $n++
        $row.Tag = $n
        $row.Cells['Order'].Value = "$n"
    }
}

function Set-GridSort {
    <#
        Arrange the rows by one column. Order = saved order (Tag). Section =
        the order Sections.json declares, not alphabetical, so sorting by
        Section reads the way the console does. Everything else: text,
        case-insensitive, ties broken by saved order so the result is stable.
    #>
    param([Parameter(Mandatory)][string]$Column, [bool]$Descending = $false)
    if ($dgv.IsCurrentCellInEditMode) { $dgv.EndEdit() }
    $selTag = $null
    if ($dgv.SelectedRows.Count -gt 0 -and -not $dgv.SelectedRows[0].IsNewRow) { $selTag = $dgv.SelectedRows[0].Tag }

    $rows = Get-GridRowsInSavedOrder
    if ($Column -ne 'Order') {
        $secIndex = @{}
        $k = 0
        foreach ($s in @(Get-ValidSections)) { $secIndex["$s"] = $k; $k++ }
        $key = {
            $v = "$($_.Cells[$Column].Value)"
            if ($Column -eq 'Section') {
                if ($secIndex.ContainsKey($v)) { '{0:000}' -f [int]$secIndex[$v] } else { '999' + $v.ToLower() }
            } else { $v.ToLower() }
        }
        $rows = @($rows | Sort-Object -Property @{ Expression = $key }, @{ Expression = { [int]$_.Tag } } -Descending:$Descending)
    } elseif ($Descending) {
        [array]::Reverse($rows)
    }

    $dgv.SuspendLayout()
    try {
        $dgv.Rows.Clear()
        if ($rows.Count -gt 0) { $dgv.Rows.AddRange([System.Windows.Forms.DataGridViewRow[]]$rows) }
    } finally { $dgv.ResumeLayout() }

    $script:GridSort.Column = $Column
    $script:GridSort.Descending = $Descending
    foreach ($c in @($dgv.Columns)) {
        $c.HeaderCell.SortGlyphDirection = if ($c.Name -eq $Column) {
            if ($Descending) { [System.Windows.Forms.SortOrder]::Descending } else { [System.Windows.Forms.SortOrder]::Ascending }
        } else { [System.Windows.Forms.SortOrder]::None }
    }
    $natural = Test-GridNaturalOrder
    $btnMoveUp.Enabled = $natural
    $btnMoveDown.Enabled = $natural

    $dgv.ClearSelection()
    if ($null -ne $selTag) {
        foreach ($r in $dgv.Rows) {
            if (-not $r.IsNewRow -and "$($r.Tag)" -eq "$selTag") {
                $r.Selected = $true
                try { $dgv.FirstDisplayedScrollingRowIndex = [Math]::Max(0, $r.Index - 3) } catch { }
                break
            }
        }
    }
}

$dgv.Add_ColumnHeaderMouseClick({
    param($s, $e)
    try {
        if ($e.ColumnIndex -lt 0) { return }
        $name = $dgv.Columns[$e.ColumnIndex].Name
        $desc = $false
        if ($script:GridSort.Column -eq $name) { $desc = -not $script:GridSort.Descending }
        Set-GridSort -Column $name -Descending $desc
    } catch { }
})

function Update-ToolsGrid {
    param($ToolObjects)
    $dgv.Rows.Clear()
    foreach ($t in @($ToolObjects)) {
        $typesVal = (@($t.ObjectTypes) -join ', ')
        $sectionVal = Get-NormalizedSection -Value "$($t.Section)"
        # BY NAME, NOT BY POSITION.
        #
        # This was a positional Rows.Add with one value per column, in column
        # order. Adding the Group column between Section and Tier then shifted
        # every value after it one cell left: Group received Tier, Tier received
        # the NewProcess boolean and rendered it as 'False', and NewProcess --
        # the Own-window flag -- received nothing and came back UNTICKED.
        #
        # Nothing errored. Saving wrote the shifted values back, so a second
        # open compounded it, and the console grew a group button called
        # 'False'. The dangerous part was invisible: Own window silently cleared
        # on every on-premises tool is how a write lands in the wrong directory.
        #
        # Exactly the failure documented on LogViewer's grid. Naming the cells
        # costs four lines and makes column order irrelevant.
        $i = $dgv.Rows.Add()
        $row = $dgv.Rows[$i]
        $row.Cells['Name'].Value         = "$($t.Name)"
        $row.Cells['Script'].Value       = "$($t.Script)"
        $row.Cells['ObjectTypes'].Value  = $typesVal
        $row.Cells['Section'].Value      = $sectionVal
        # THE OPERATOR'S OWN GROUPING, falling back to whatever Tools.json
        # still says. The fallback is what makes the first open after this
        # change show the groups that are already there rather than an empty
        # column -- saving then carries them into the profile.
        $gk = ''
        try { $gk = Get-ParkedKey -Script "$($t.Script)" } catch { }
        $row.Cells['Group'].Value = $(
            if ($gk -and $script:UserGroups.ContainsKey($gk)) { $script:UserGroups[$gk] }
            else { "$($t.Group)".Trim() })
    }
    # A fresh load is the saved order by definition. Forget any sort that was
    # applied to the previous contents, or its glyph and disabled Move buttons
    # would describe an arrangement that no longer exists.
    $script:GridSort.Column = 'Order'
    $script:GridSort.Descending = $false
    foreach ($c in @($dgv.Columns)) { $c.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None }
    try { $btnMoveUp.Enabled = $true; $btnMoveDown.Enabled = $true } catch { }
    Set-OrderColumn
}

$dgv.Add_DefaultValuesNeeded({
    param($sender, $e)
    $e.Row.Cells['Section'].Value = (Get-ValidSections)[0]
})

function Get-ToolsFromGrid {
    $result = @()
    $n = 0
    # SAVED ORDER, NOT SCREEN ORDER. The grid may be sorted by any column;
    # the file is written in the order the console shows the buttons.
    foreach ($row in (Get-GridRowsInSavedOrder)) {
        if ($row.IsNewRow) { continue }
        $name = "$($row.Cells['Name'].Value)".Trim()
        $script = "$($row.Cells['Script'].Value)".Trim()
        if ($name -eq '' -and $script -eq '') { continue }
        $n++
        $types = @(("$($row.Cells['ObjectTypes'].Value)" -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
        $section = Get-NormalizedSection -Value "$($row.Cells['Section'].Value)"
        # Ordered, so the written file keeps a predictable field order and
        # optional fields can be OMITTED rather than written as "". Emitting an
        # empty value on every tool would add a line to all fifty-odd entries
        # the first time anyone opened ToolManager and hit Save -- a whole-file
        # diff for nothing. Config\ is the operator's, and a config editor
        # should change what you changed.
        #
        # ScriptOnPrem used to be written here. Removed 2026-08-13 with the
        # rest of the AD twin feature; any value still in Tools.json is now
        # ignored on read and dropped on the next Save.
        $fields = [ordered]@{
            Order       = $n
            Name        = $name
            Script      = $script
        }
        $fields['ObjectTypes'] = $types
        $fields['Section']     = $section

        # GROUP IS NOT WRITTEN HERE ANY MORE. Tools.json is shared by everyone
        # who opens this folder, and grouping is one person's opinion about
        # where to keep a tool: two engineers filing the same tools differently
        # is the normal case, not a conflict to resolve. It goes to the
        # operator's own profile instead -- see Save-ToolGroups below.
        #
        # A Group key left in Tools.json by an older console is dropped on the
        # next save from this window, which is the migration: whoever saves
        # first takes the existing groups into their profile and stops handing
        # them to everybody else.

        $result += [PSCustomObject]$fields
    }
    return $result
}

# Read once, before the grid is filled: every row asks it for its group.
$script:UserGroups = @{}
try { $script:UserGroups = Get-UserToolGroupMap } catch { $script:UserGroups = @{} }

$script:CurrentTools = @(Read-ToolsJson -Path $ToolsConfigPath)
Update-ToolsGrid -ToolObjects $script:CurrentTools

# THERE IS NO HINT LABEL, DELIBERATELY.
#
# A 700-character paragraph used to sit under the grid explaining every column.
# It was the only documentation this tool had, so it had to carry everything --
# and it did that badly, clipped mid-word in a strip too short for it, and
# silently went stale when the Group column was added without anyone updating
# it.
#
# All of it now lives in the ? button, which is also where every other tool
# keeps its explanation. Leaving a summary on screen as well would defeat that:
# permanent on-screen text is the reason nobody opens help, and the copy that
# gets read is then the copy that nobody maintains.
#
# The space it occupied goes to the grid.

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
# Width is recomputed in the layout block below to fill the gap up to Save.
# AutoEllipsis so that if the path is still too long for the window the line
# ends in '...' -- a visible signal there is more, rather than a sentence that
# stops mid-path and reads as complete.
$lblStatus.AutoEllipsis = $true
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
            $null = Show-ToolDialog(
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
        Show-ToolDialog(
            "ThemeManager.ps1 was not found next to this script.",
            'Tool Manager', 'OK', 'Warning') | Out-Null
    }
})
$form.Controls.Add($btnTheme)

#region Settings dialog -------------------------------------------------------
# Show-SettingsEditor MOVED TO Common.ps1, 2026-08-20.
#
# It was only ever reachable from this window, which made "change my email
# address" a trip through the tool registry. The dashboard now has its own
# Settings button and Settings.ps1 opens the same dialog -- so the dialog had to
# stop belonging to one caller.
#
# COPIED WOULD HAVE BEEN WRONG. Two settings editors drift, and the one you are
# not looking at is the one that writes the file. Common.ps1 already held every
# helper it calls (Get-ToolkitSettings, Save-ToolkitSettings, Get-EdgeProfiles,
# Get-SpoAdminUrl, Get-ExchangeServerUri), so it moved without changing a line.

$btnSettings          = New-Object System.Windows.Forms.Button
$btnSettings.Text     = 'Settings...'
$btnSettings.Location = New-Object System.Drawing.Point(228, 532)
$btnSettings.Size     = New-Object System.Drawing.Size(100, 28)
$btnSettings.Add_Click({
    if (Show-SettingsEditor) {
        $lblStatus.Text = "Settings saved. Reports -> $(Get-ReportPath)"
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

$rowC = $ch - 40          # Sections... / Theme...  and  Save / Close
$rowB = $rowC - 36        # Move Up / Move Down / Reload
$rowA = $rowB - 36        # Add Tool / Delete Selected
# 44, not 24. The hint wraps to two lines at this width and the old figure gave
# The grid now runs all the way down to the button row -- the hint label that
# used to sit between them is gone, so there is nothing to reserve space for.
$lblNote.Width = $cw - 24
$dgv.Width     = $cw - 24
$dgv.Height    = $rowA - 8 - $dgv.Top

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

# The status line fills whatever is left between Reload and Save, rather than a
# fixed 300px that was already too short for a full path and got shorter every
# time the window grew.
$lblStatus.Width = [Math]::Max(200, ($btnSave.Left - 12 - $lblStatus.Left))

$dgv.Anchor       = 'Top,Left,Bottom,Right'
$lblNote.Anchor   = 'Top,Left,Right'
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
    # Only in saved-order view; the buttons are disabled otherwise, this is
    # the belt to that brace.
    if (-not (Test-GridNaturalOrder)) { return }
    $idx = $dgv.SelectedRows[0].Index
    if ($dgv.Rows[$idx].IsNewRow) { return }
    $targetIdx = $idx + $Delta
    if ($targetIdx -lt 0 -or $targetIdx -ge $dgv.Rows.Count -or $dgv.Rows[$targetIdx].IsNewRow) { return }
    # EVERY DATA COLUMN, DERIVED -- NOT A HAND-MAINTAINED LIST.
    #
    # This was a literal list: Name, Script, ObjectTypes, Section, Group. When
    # the ScriptOnPrem column was added nobody added it here, so Move Up/Down
    # carried a tool's identity to the new row and left its AD twin sitting on
    # the OLD row position, where it silently reattached to whatever landed
    # there. Message Trace, DL Send To/As and Connect-ExSE all acquired twins
    # they never had, Save wrote it back, and it compounded every session.
    #
    # The twin feature is gone now (ARCHITECTURE section 30), but the trap was
    # never really about twins: it was a list of column names that had to be
    # updated by hand whenever a column was added, with nothing to enforce it
    # and silent data corruption as the failure mode. The same omission also
    # produced the Group/Tier/NewProcess shift documented in Update-ToolsGrid.
    #
    # Asking the grid for its own columns cannot fall out of date. Order is
    # excluded because Set-OrderColumn recomputes it from row position
    # immediately below -- swapping it would be undone anyway.
    foreach ($col in @($dgv.Columns)) {
        if ($col.Name -eq 'Order') { continue }
        $a = $dgv.Rows[$idx].Cells[$col.Name].Value
        $b = $dgv.Rows[$targetIdx].Cells[$col.Name].Value
        $dgv.Rows[$idx].Cells[$col.Name].Value = $b
        $dgv.Rows[$targetIdx].Cells[$col.Name].Value = $a
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
        $null = Show-ToolDialog(
            "Select a row first.", 'Delete Tool', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $idx = $dgv.SelectedRows[0].Index
    if ($dgv.Rows[$idx].IsNewRow) { return }
    $name = "$($dgv.Rows[$idx].Cells['Name'].Value)"
    $label = if ($name -ne '') { "'$name'" } else { 'this row' }
    $confirm = Show-ToolDialog(
        "Delete $label from the grid? This isn't written to Tools.json until you click Save.",
        'Confirm Delete', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    $dgv.Rows.RemoveAt($idx)
    Set-OrderColumn
})

$btnReload.Add_Click({
    if ($dgv.IsCurrentCellInEditMode) { $dgv.EndEdit() }
    $confirm = Show-ToolDialog(
        "Reload from '$ToolsConfigPath'? Any unsaved edits in the grid will be discarded.",
        'Confirm Reload', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    try {
        $script:CurrentTools = @(Read-ToolsJson -Path $ToolsConfigPath)
        Update-ToolsGrid -ToolObjects $script:CurrentTools
        $lblStatus.Text = "Reloaded $($script:CurrentTools.Count) tool(s) from $ToolsConfigPath at $(Get-Date -Format 'HH:mm:ss')."
    } catch {
        $null = Show-ToolDialog(
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
            $null = Show-ToolDialog(
                "Unrecognized ObjectTypes value(s): $($badTypes -join ', ')`n`nValid values: $($validObjectTypes -join ', ')",
                'Validation Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $blankNames = @($tools | Where-Object { $_.Name -eq '' -or $_.Script -eq '' })
        if ($blankNames.Count -gt 0) {
            $null = Show-ToolDialog(
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
            $null = Show-ToolDialog(
                "This will still save -- AdminConsole's button panels scroll, so nothing is hidden -- but the following section(s) now exceed the grid size they're sized for:`n`n$($overCapacity -join "`n")",
                'Section Over Capacity', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        Write-ToolsJson -Path $ToolsConfigPath -ToolObjects $tools
        $script:CurrentTools = $tools

        # ---- GROUPING GOES TO THIS OPERATOR'S PROFILE, NOT Tools.json -------
        #
        # Read straight off the grid rather than from $tools, because $tools no
        # longer carries Group at all -- that is the point of the change.
        #
        # EVERY ROW IS RECORDED, INCLUDING THE BLANK ONES. A blank has to be
        # written rather than skipped: it is how 'I took this one out of its
        # group' is said, and skipping it would leave the console falling back
        # to whatever Tools.json used to say and the tool would spring back
        # into a group the operator had just removed it from.
        $groupNote = ''
        try {
            $gmap = [ordered]@{}
            foreach ($row in $dgv.Rows) {
                if ($row.IsNewRow) { continue }
                $sc = "$($row.Cells['Script'].Value)".Trim()
                if (-not $sc) { continue }
                $k = Get-ParkedKey -Script $sc
                if (-not $k) { continue }
                $gmap[$k] = "$($row.Cells['Group'].Value)".Trim()
            }
            $layout = Get-UserLayoutOwn
            if (-not $layout) { $layout = [PSCustomObject]@{} }
            Add-Member -InputObject $layout -NotePropertyName 'Groups' -NotePropertyValue ([PSCustomObject]$gmap) -Force
            $up = Save-UserLayout -Layout $layout
            $script:UserGroups = Get-UserToolGroupMap
            $named = @($gmap.Values | Where-Object { $_ }).Count
            $groupNote = " Grouping ($named tool(s)) saved to $(Split-Path -Leaf $up) -- yours only."
        } catch {
            $groupNote = " WARNING: the grouping could not be saved to your profile ($($_.Exception.Message)); Tools.json was written."
        }

        Set-OrderColumn
        $lblStatus.Text = "Saved $($tools.Count) tool(s) to $ToolsConfigPath at $(Get-Date -Format 'HH:mm:ss').$groupNote"
    } catch {
        $null = Show-ToolDialog(
            "Save failed: $($_.Exception.Message)",
            'Save Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$btnClose.Add_Click({ $form.Close() })

# Wrapped, like TenantManager's. A missing or unreadable help block must not
# stop the editor opening -- the tool is how you fix a broken Tools.json, so it
# has to survive a toolkit that is not entirely well.
try { Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Tool Manager' | Out-Null } catch { }

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
[void]$form.ShowDialog()
#endregion
