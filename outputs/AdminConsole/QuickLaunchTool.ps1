#Requires -Version 5.1
<#
    QuickLaunchTool.ps1
    Editor for Config\QuickLaunch.json -- the links QuickLaunch.ps1 shows.
    Add, remove, reorder, and set which Edge profile each link opens in.

    Author   : Glenn Armbruster
    Revised  : 2026-08-06
#>
# TOOLKIT-HELP-BEGIN
# Edits the links that QuickLaunch shows.
#
#   * Section -- any name you like; links are grouped under it, and sections
#     appear in the order they first appear in the list
#   * Move Up / Move Down -- controls both the section order and the order
#     within a section
#   * Edge profile -- picked from the profiles you actually have. Blank means
#     Edge's Default profile.
#   * Admin -- makes the button dark blue, for links needing an elevated account
#   * Test Selected -- opens the highlighted row exactly as QuickLaunch would,
#     so a wrong profile turns up here rather than at a sign-in prompt
#
# Nothing is saved until you press Save.
# TOOLKIT-HELP-END

param(
    [string]$ConfigPath = ''
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

if (-not $ConfigPath) { $ConfigPath = Get-ConfigFilePath -Name 'QuickLaunch.json' }

#region Data --------------------------------------------------------------------
function Read-QuickLaunchJson {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        if ("$raw".Trim() -eq '') { return @() }
        $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
        if ($parsed.PSObject.Properties['Links']) { return @($parsed.Links) }
        return @($parsed)
    } catch {
        $null = Show-ToolDialog(
            "Could not read $Path.`n`n$($_.Exception.Message)",
            'QuickLaunchTool', 'OK', 'Error')
        return @()
    }
}

function Write-QuickLaunchJson {
    param([string]$Path, $Links)
    $items = @($Links)
    $json  = $items | ConvertTo-Json -Depth 4
    # ConvertTo-Json emits a bare object, not an array, when there is exactly
    # one item -- and QuickLaunch.ps1 would then read a single link as an
    # object with Label/Url properties rather than a one-element list. Same
    # wrapping ToolManager does, for the same reason.
    if ($items.Count -le 1) { $json = "[`r`n$json`r`n]" }
    Set-Content -LiteralPath $Path -Value $json -Encoding UTF8
}

function Get-ProfileChoices {
    <#
        Edge's real profiles, plus a blank meaning Default.

        Populated live rather than typed free-hand: a profile display name that
        does not exist resolves to nothing, and Start-UrlInBrowserProfile then
        falls back to the last-used profile. That failure is silent, and for an
        elevated link it means a sign-in prompt at best and the wrong identity
        at worst. A dropdown of names that provably exist cannot be typo'd.
    #>
    $names = @(Get-EdgeProfiles | ForEach-Object { "$($_.DisplayName)" } |
               Where-Object { $_ } | Sort-Object -Unique)
    return @('') + $names
}

$script:ProfileChoices = Get-ProfileChoices
#endregion

#region UI ----------------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = "Quick Launch Tool -- $ConfigPath"
$form.Size          = New-Object System.Drawing.Size(1020, 640)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$dgv                        = New-Object System.Windows.Forms.DataGridView
$dgv.Location               = New-Object System.Drawing.Point(12, 12)
$dgv.Size                   = New-Object System.Drawing.Size(980, 460)
$dgv.AllowUserToAddRows     = $false
$dgv.AllowUserToDeleteRows  = $false
$dgv.AllowUserToResizeRows  = $false
$dgv.RowHeadersVisible      = $false
$dgv.SelectionMode          = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dgv.MultiSelect            = $false
$dgv.AutoSizeColumnsMode    = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
$form.Controls.Add($dgv)

$colSection            = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colSection.Name       = 'Section'
$colSection.HeaderText = 'Section'
$colSection.Width      = 150
[void]$dgv.Columns.Add($colSection)

$colLabel            = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colLabel.Name       = 'Label'
$colLabel.HeaderText = 'Button text'
$colLabel.Width      = 170
[void]$dgv.Columns.Add($colLabel)

$colUrl            = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colUrl.Name       = 'Url'
$colUrl.HeaderText = 'URL'
$colUrl.Width      = 330
[void]$dgv.Columns.Add($colUrl)

$colProfile              = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$colProfile.Name         = 'Profile'
$colProfile.HeaderText   = 'Edge profile'
$colProfile.Width        = 140
# FlatStyle Flat so the cell reads as a dropdown without having to click it.
$colProfile.FlatStyle    = [System.Windows.Forms.FlatStyle]::Flat
foreach ($p in $script:ProfileChoices) { [void]$colProfile.Items.Add($p) }
[void]$dgv.Columns.Add($colProfile)

$colAdmin            = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colAdmin.Name       = 'Admin'
$colAdmin.HeaderText = 'Admin'
$colAdmin.Width      = 55
[void]$dgv.Columns.Add($colAdmin)

$colNote            = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colNote.Name       = 'Note'
$colNote.HeaderText = 'Note (tooltip)'
$colNote.Width      = 120
[void]$dgv.Columns.Add($colNote)

function Update-Grid {
    param($Links)
    $dgv.Rows.Clear()
    foreach ($l in @($Links)) {
        $prof = "$($l.Profile)".Trim()
        # A profile that no longer exists -- renamed or deleted in Edge -- is
        # added to the dropdown for this session rather than silently blanked.
        # Blanking it would change where the link opens without saying so, and
        # for an Admin link that is the whole problem this tool is meant to
        # avoid. It stays visible until it is deliberately changed.
        if ($prof -and -not ($colProfile.Items -contains $prof)) {
            [void]$colProfile.Items.Add($prof)
        }
        [void]$dgv.Rows.Add(
            "$($l.Section)", "$($l.Label)", "$($l.Url)", $prof,
            [bool]$l.Admin, "$($l.Note)")
    }
}

function Get-GridLinks {
    $result = @()
    foreach ($row in $dgv.Rows) {
        if ($row.IsNewRow) { continue }
        $label = "$($row.Cells['Label'].Value)".Trim()
        $url   = "$($row.Cells['Url'].Value)".Trim()
        if ($label -eq '' -and $url -eq '') { continue }
        $section = "$($row.Cells['Section'].Value)".Trim()
        if (-not $section) { $section = 'General' }
        $result += [PSCustomObject]@{
            Section = $section
            Label   = $label
            Url     = $url
            Profile = "$($row.Cells['Profile'].Value)".Trim()
            # An unticked checkbox cell reads back as $null, not $false.
            Admin   = [bool]$row.Cells['Admin'].Value
            Note    = "$($row.Cells['Note'].Value)".Trim()
        }
    }
    return $result
}

function Test-GridValid {
    <#
        Returns a list of complaints, empty when the grid is fit to save.

        Checked here rather than in QuickLaunch.ps1 because this is where the
        mistake is being made and where it can still be corrected in place. The
        launcher skips unusable links quietly, which is right for a tool you
        are trying to use and wrong for a tool you are trying to edit.
    #>
    $problems = @()
    $n = 0
    foreach ($row in $dgv.Rows) {
        if ($row.IsNewRow) { continue }
        $n++
        $label = "$($row.Cells['Label'].Value)".Trim()
        $url   = "$($row.Cells['Url'].Value)".Trim()
        if ($label -eq '' -and $url -eq '') { continue }
        if (-not $label) { $problems += "Row $n has a URL but no button text." }
        if (-not $url)   { $problems += "Row $n ('$label') has no URL." }
        elseif ($url -notmatch '^(?i)https?://') {
            $problems += "Row $n ('$label') is not an http/https URL: $url"
        }
    }
    return $problems
}

$lblHint          = New-Object System.Windows.Forms.Label
$lblHint.Text     = 'Section: any name you like -- links are grouped by it, and sections appear in the order they first appear here. Use Move Up / Move Down to order both. Edge profile: blank means Edge''s Default profile. Admin: dark blue button, for links that need an elevated account.'
$lblHint.Location = New-Object System.Drawing.Point(12, 480)
$lblHint.Size     = New-Object System.Drawing.Size(980, 34)
$lblHint.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 8, [System.Drawing.FontStyle]::Italic)
$lblHint.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblHint)

$btnAdd          = New-Object System.Windows.Forms.Button
$btnAdd.Text     = 'Add Link'
$btnAdd.Location = New-Object System.Drawing.Point(12, 522)
$btnAdd.Size     = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnAdd)

$btnDelete          = New-Object System.Windows.Forms.Button
$btnDelete.Text     = 'Delete Link'
$btnDelete.Location = New-Object System.Drawing.Point(120, 522)
$btnDelete.Size     = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnDelete)

$btnUp          = New-Object System.Windows.Forms.Button
$btnUp.Text     = 'Move Up'
$btnUp.Location = New-Object System.Drawing.Point(236, 522)
$btnUp.Size     = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnUp)

$btnDown          = New-Object System.Windows.Forms.Button
$btnDown.Text     = 'Move Down'
$btnDown.Location = New-Object System.Drawing.Point(344, 522)
$btnDown.Size     = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnDown)

$btnTest          = New-Object System.Windows.Forms.Button
$btnTest.Text     = 'Test Selected'
$btnTest.Location = New-Object System.Drawing.Point(460, 522)
$btnTest.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnTest)

$btnSave          = New-Object System.Windows.Forms.Button
$btnSave.Text     = 'Save'
$btnSave.Location = New-Object System.Drawing.Point(780, 522)
$btnSave.Size     = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnSave)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(890, 522)
$btnClose.Size     = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(12, 560)
$lblStatus.Size     = New-Object System.Drawing.Size(980, 20)
$form.Controls.Add($lblStatus)

$btnAdd.Add_Click({
    $i = $dgv.Rows.Add('General', '', 'https://', '', $false, '')
    $dgv.CurrentCell = $dgv.Rows[$i].Cells['Label']
    $dgv.BeginEdit($true)
    $lblStatus.Text = 'Row added -- not saved yet.'
})

$btnDelete.Add_Click({
    if ($dgv.CurrentRow -and -not $dgv.CurrentRow.IsNewRow) {
        $label = "$($dgv.CurrentRow.Cells['Label'].Value)"
        $dgv.Rows.Remove($dgv.CurrentRow)
        $lblStatus.Text = "Removed '$label' -- not saved yet."
    }
})

function Move-SelectedRow {
    <#
        Moves the row object itself rather than copying cell values.

        Copying values means enumerating every column, and the day a column is
        added is the day one of them stops being copied -- a bug that shows up
        as a single field mysteriously not moving. Remove-and-reinsert has no
        such per-column knowledge.
    #>
    param([int]$Delta)
    $row = $dgv.CurrentRow
    if (-not $row -or $row.IsNewRow) { return }
    $i = $row.Index
    $j = $i + $Delta
    if ($j -lt 0 -or $j -ge $dgv.Rows.Count) { return }
    $dgv.Rows.Remove($row)
    $dgv.Rows.Insert($j, $row)
    $dgv.ClearSelection()
    $dgv.Rows[$j].Selected = $true
    $dgv.CurrentCell = $dgv.Rows[$j].Cells['Label']
    $lblStatus.Text = 'Reordered -- not saved yet.'
}
$btnUp.Add_Click({ Move-SelectedRow -Delta -1 })
$btnDown.Add_Click({ Move-SelectedRow -Delta 1 })

$btnTest.Add_Click({
    # Opens the selected row exactly as QuickLaunch would, so a wrong profile
    # is found here rather than in front of a sign-in prompt later.
    $row = $dgv.CurrentRow
    if (-not $row -or $row.IsNewRow) { $lblStatus.Text = 'Select a row first.'; return }
    $url = "$($row.Cells['Url'].Value)".Trim()
    if ($url -notmatch '^(?i)https?://') {
        $lblStatus.Text = "Not an http/https URL: $url"
        return
    }
    $prof = "$($row.Cells['Profile'].Value)".Trim()
    if (-not $prof) { $prof = 'Default' }
    try {
        $st = Start-UrlInBrowserProfile -Url $url -ProfileName $prof
        $lblStatus.Text = "$st"
    } catch {
        $lblStatus.Text = "Could not open: $($_.Exception.Message)"
    }
})

$btnSave.Add_Click({
    $problems = Test-GridValid
    if ($problems.Count -gt 0) {
        $null = Show-ToolDialog(
            ("Nothing was saved.`n`n" + ($problems -join "`n")),
            'Fix these first', 'OK', 'Warning')
        $lblStatus.Text = "$($problems.Count) problem(s) -- nothing saved."
        return
    }
    try {
        $links = Get-GridLinks
        Write-QuickLaunchJson -Path $ConfigPath -Links $links
        $lblStatus.Text = "Saved $($links.Count) link(s) to $ConfigPath"
    } catch {
        $null = Show-ToolDialog(
            "Could not save.`n`n$($_.Exception.Message)", 'QuickLaunchTool', 'OK', 'Error')
        $lblStatus.Text = 'Save failed.'
    }
})

$btnClose.Add_Click({ $form.Close() })

Update-Grid -Links (Read-QuickLaunchJson -Path $ConfigPath)

if (@(Get-EdgeProfiles).Count -eq 0) {
    $lblStatus.Text = 'No Edge profiles found -- the profile dropdown will be empty and every link will use the default browser.'
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(176, 96, 0)
}

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled, and BringToFront'd inside the helper because WinForms
# draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'QuickLaunch Tool' | Out-Null

[void]$form.ShowDialog()
$form.Dispose()
#endregion
