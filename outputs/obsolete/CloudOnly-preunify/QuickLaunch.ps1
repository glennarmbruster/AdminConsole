#Requires -Version 5.1
<#
    QuickLaunch.ps1
    URL launcher. Sections and links come from Config\QuickLaunch.json, which
    QuickLaunchTool.ps1 edits.

    Author   : Glenn Armbruster
    Revised  : 2026-08-06
#>
# TOOLKIT-HELP-BEGIN
# Your own set of links, as buttons.
#
# Each link opens in the Edge profile it was set up with, so admin portals open
# with the elevated account and everyday links do not. Links marked as admin are
# dark blue, so you can see which is which before you click.
#
# Type in the Find box to narrow a long list.
#
# Edit the links, their order and their sections in QuickLaunch Tool.
#
# Opens links in a browser. Connects to nothing.
# TOOLKIT-HELP-END

param(
    [string]$ConfigPath = ''
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not $ConfigPath) { $ConfigPath = Get-ConfigFilePath -Name 'QuickLaunch.json' }

#region Data --------------------------------------------------------------------
function Read-QuickLaunchJson {
    <#
        Links in file order. Section order is the order each section FIRST
        appears, and order within a section is row order.

        No separate order field, and no separate sections file. Both were
        considered: Tools.json carries an Order property and Sections.json
        declares sections independently, and the cost of that split is that a
        section can exist with no tools, or a tool can name a section that does
        not exist, and both states have to be handled every time either file is
        read.

        Here there is one list. A section exists because something is in it,
        and moving a row is the only way to change any order. Fewer states, and
        the editor's Move Up / Move Down is the whole ordering model.
    #>
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        if ("$raw".Trim() -eq '') { return @() }
        $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
        # Accept both a bare array and { "Links": [...] }. The first is what
        # ConvertTo-Json produces from a list and the second is what someone
        # hand-writing the file tends to reach for.
        if ($parsed.PSObject.Properties['Links']) { return @($parsed.Links) }
        return @($parsed)
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Could not read $Path.`n`n$($_.Exception.Message)`n`nFix it in QuickLaunchTool, or delete the file to start over.",
            'QuickLaunch', 'OK', 'Error')
        return @()
    }
}

function Test-LinkUsable {
    <#
        A link needs a label and an http(s) URL.

        Scheme is checked because Start-Process will happily hand file: and
        anything else to the shell. A launcher that can be pointed at a local
        executable by editing a JSON file is a different and much less pleasant
        kind of tool.
    #>
    param($Link)
    if (-not "$($Link.Label)".Trim()) { return $false }
    $u = "$($Link.Url)".Trim()
    return ($u -match '^(?i)https?://')
}
#endregion

#region Colour ------------------------------------------------------------------
<#
    Two roles for now, both read from the active theme with defaults in code so
    an older Themes.json still renders correctly.

      Admin  -- dark blue on white. These open with an elevated identity and
                the point of the colour is that you can see that BEFORE
                clicking, not after the sign-in prompt tells you.
      Normal -- the same accent green the console already uses, so the
                launcher looks like part of the same toolkit.

    A per-link colour override is the obvious next step and is deliberately not
    here yet: it wants a colour picker in the editor and a decision about
    whether an override beats the Admin role, and neither is worth guessing at.
#>
function Get-LinkPalette {
    $theme = Get-ConsoleTheme

    $p = @{}
    $p.AdminBack   = ConvertTo-ThemeColor $theme.QuickLaunchAdminBack
    $p.AdminText   = ConvertTo-ThemeColor $theme.QuickLaunchAdminText
    $p.AdminHover  = ConvertTo-ThemeColor $theme.QuickLaunchAdminHover
    $p.AdminBorder = ConvertTo-ThemeColor $theme.QuickLaunchAdminBorder
    if (-not $p.AdminBack)   { $p.AdminBack   = [System.Drawing.Color]::FromArgb(18, 45, 92) }
    if (-not $p.AdminText)   { $p.AdminText   = [System.Drawing.Color]::FromArgb(255, 255, 255) }
    if (-not $p.AdminHover)  { $p.AdminHover  = [System.Drawing.Color]::FromArgb(27, 68, 138) }
    if (-not $p.AdminBorder) { $p.AdminBorder = [System.Drawing.Color]::FromArgb(64, 122, 214) }

    $p.Back   = ConvertTo-ThemeColor $theme.ServiceNowBack
    $p.Text   = ConvertTo-ThemeColor $theme.ServiceNowText
    $p.Hover  = ConvertTo-ThemeColor $theme.ServiceNowHover
    $p.Border = ConvertTo-ThemeColor $theme.ServiceNowBorder
    if (-not $p.Back)   { $p.Back   = [System.Drawing.Color]::FromArgb(10, 92, 66) }
    if (-not $p.Text)   { $p.Text   = [System.Drawing.Color]::FromArgb(255, 255, 255) }
    if (-not $p.Hover)  { $p.Hover  = [System.Drawing.Color]::FromArgb(14, 120, 86) }
    if (-not $p.Border) { $p.Border = [System.Drawing.Color]::FromArgb(22, 160, 114) }

    return $p
}

function Set-LinkButtonStyle {
    param($Button, $Link, $Palette)

    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    if ([bool]$Link.Admin) {
        $Button.BackColor = $Palette.AdminBack
        $Button.ForeColor = $Palette.AdminText
        $Button.FlatAppearance.MouseOverBackColor = $Palette.AdminHover
        $Button.FlatAppearance.BorderColor        = $Palette.AdminBorder
    } else {
        $Button.BackColor = $Palette.Back
        $Button.ForeColor = $Palette.Text
        $Button.FlatAppearance.MouseOverBackColor = $Palette.Hover
        $Button.FlatAppearance.BorderColor        = $Palette.Border
    }
}
#endregion

#region UI ----------------------------------------------------------------------
$script:Links   = @(Read-QuickLaunchJson -Path $ConfigPath)
$script:Palette = $null

$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Quick Launch'
$form.Size          = New-Object System.Drawing.Size(720, 760)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblFind          = New-Object System.Windows.Forms.Label
$lblFind.Text     = 'Find:'
$lblFind.Location = New-Object System.Drawing.Point(12, 15)
$lblFind.Size     = New-Object System.Drawing.Size(40, 18)
$form.Controls.Add($lblFind)

$txtFind          = New-Object System.Windows.Forms.TextBox
$txtFind.Location = New-Object System.Drawing.Point(54, 12)
$txtFind.Size     = New-Object System.Drawing.Size(300, 24)
$form.Controls.Add($txtFind)

$btnClearFind          = New-Object System.Windows.Forms.Button
$btnClearFind.Text     = 'Clear'
$btnClearFind.Location = New-Object System.Drawing.Point(362, 11)
$btnClearFind.Size     = New-Object System.Drawing.Size(70, 26)
$form.Controls.Add($btnClearFind)

$lblCount           = New-Object System.Windows.Forms.Label
$lblCount.Tag       = 'NoTheme'
$lblCount.Location  = New-Object System.Drawing.Point(442, 15)
$lblCount.Size      = New-Object System.Drawing.Size(250, 18)
$lblCount.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblCount)

# Scrolling panel, because the number of links is the user's business and a
# form that silently stops showing them past the bottom edge would be the kind
# of quiet truncation that hides things.
$panel            = New-Object System.Windows.Forms.Panel
$panel.Location   = New-Object System.Drawing.Point(12, 44)
$panel.Size       = New-Object System.Drawing.Size(682, 620)
$panel.AutoScroll = $true
$form.Controls.Add($panel)

$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 20000

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Hover a button to see where it goes and which profile it opens in.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 672)
$lblStatus.Size     = New-Object System.Drawing.Size(560, 34)
$form.Controls.Add($lblStatus)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(594, 672)
$btnClose.Size     = New-Object System.Drawing.Size(100, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

function Invoke-LinkClick {
    param($Link)

    # No profile on the link means the DEFAULT profile, named explicitly.
    #
    # 'Default' is Edge's own directory name for it. That is not the same as
    # passing no --profile-directory at all, which gives the LAST USED profile
    # -- a different one depending on what you had open, and the wrong-identity
    # failure Start-UrlInBrowserProfile exists to prevent. Blank here means one
    # specific profile, every time.
    $profileName = "$($Link.Profile)".Trim()
    if (-not $profileName) { $profileName = 'Default' }

    try {
        $st = Start-UrlInBrowserProfile -Url "$($Link.Url)" -ProfileName $profileName
        $lblStatus.Text = "$($Link.Label) -- $st"
    } catch {
        $lblStatus.Text = "Could not open $($Link.Label): $($_.Exception.Message)"
    }
}

function Update-LinkPanel {
    <#
        Rebuilds the panel from $script:Links, filtered by the find box.

        Rebuild rather than hide-and-reflow: a filtered view has to close the
        gaps, and moving existing controls means tracking their original
        positions, which is more state than the controls are worth.
    #>
    param([string]$Filter = '')

    $panel.SuspendLayout()
    # Dispose rather than just Clear. Clear removes controls from the panel but
    # leaves the handles alive, and this runs on every keystroke in the find
    # box.
    foreach ($c in @($panel.Controls)) { $panel.Controls.Remove($c); $c.Dispose() }

    $usable = @($script:Links | Where-Object { Test-LinkUsable -Link $_ })

    $f = "$Filter".Trim()
    if ($f) {
        # Matched against everything visible or semi-visible: label, section,
        # note and the URL itself. Searching for 'sharepoint' should find a
        # link called 'Sites' that points at sharepoint.com.
        $usable = @($usable | Where-Object {
            "$($_.Label)"   -like "*$f*" -or
            "$($_.Section)" -like "*$f*" -or
            "$($_.Note)"    -like "*$f*" -or
            "$($_.Url)"     -like "*$f*"
        })
    }

    $total = @($script:Links | Where-Object { Test-LinkUsable -Link $_ }).Count
    $lblCount.Text = if ($f) { "$($usable.Count) of $total shown" } else { "$total link(s)" }

    if ($usable.Count -eq 0) {
        $lbl          = New-Object System.Windows.Forms.Label
        $lbl.Text     = if ($f) { "Nothing matches '$f'." }
                        else { 'No links yet. Add some in QuickLaunchTool.' }
        $lbl.Location = New-Object System.Drawing.Point(8, 8)
        $lbl.Size     = New-Object System.Drawing.Size(600, 20)
        $panel.Controls.Add($lbl)
        $panel.ResumeLayout()
        return
    }

    if (-not $script:Palette) { $script:Palette = Get-LinkPalette }

    # Section order = order of first appearance in the file. Select -Unique
    # preserves that; Sort-Object would impose alphabetical and quietly discard
    # the ordering the editor exists to let you control.
    $sections = @($usable | ForEach-Object { "$($_.Section)".Trim() } |
                  ForEach-Object { if ($_) { $_ } else { 'General' } } |
                  Select-Object -Unique)

    $y = 8
    foreach ($s in $sections) {
        $items = @($usable | Where-Object {
            $sec = "$($_.Section)".Trim(); if (-not $sec) { $sec = 'General' }
            $sec -eq $s
        })
        if ($items.Count -eq 0) { continue }

        $rows = [Math]::Ceiling($items.Count / 3.0)

        $grp          = New-Object System.Windows.Forms.GroupBox
        $grp.Text     = $s
        $grp.Location = New-Object System.Drawing.Point(8, $y)
        $grp.Size     = New-Object System.Drawing.Size(645, (26 + ($rows * 38)))
        $panel.Controls.Add($grp)

        $i = 0
        foreach ($lnk in $items) {
            $col = $i % 3
            $row = [Math]::Floor($i / 3)

            $b          = New-Object System.Windows.Forms.Button
            $b.Text     = "$($lnk.Label)"
            $b.Location = New-Object System.Drawing.Point((12 + ($col * 208)), (22 + ($row * 38)))
            $b.Size     = New-Object System.Drawing.Size(200, 32)
            # Tag carries the link. Set-ConsoleTheme skips controls whose Tag
            # contains 'NoTheme', which is not usable here because Tag is
            # holding data -- so instead the theme is applied to the form FIRST
            # and these buttons are painted afterwards. Order, not tagging.
            $b.Tag      = $lnk

            $profShown = "$($lnk.Profile)".Trim()
            if (-not $profShown) { $profShown = 'Default' }
            $tip = "$($lnk.Url)`r`nProfile: $profShown"
            if ([bool]$lnk.Admin) { $tip += '  (elevated)' }
            if ("$($lnk.Note)".Trim()) { $tip = "$($lnk.Note)`r`n$tip" }
            $toolTip.SetToolTip($b, $tip)

            $b.Add_Click({ Invoke-LinkClick -Link $this.Tag })

            Set-LinkButtonStyle -Button $b -Link $lnk -Palette $script:Palette
            $grp.Controls.Add($b)
            $i++
        }
        $y += $grp.Height + 8
    }
    $panel.ResumeLayout()
}

$txtFind.Add_TextChanged({ Update-LinkPanel -Filter $txtFind.Text })
$btnClearFind.Add_Click({ $txtFind.Text = ''; $txtFind.Focus() })

# Theme first, links after. Set-ConsoleTheme repaints every Button it finds, so
# painting the link buttons before it would simply be undone.
Set-ConsoleTheme -Form $form
$script:Palette = Get-LinkPalette
Update-LinkPanel

# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled, and BringToFront'd inside the helper because WinForms
# draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'QuickLaunch' | Out-Null

[void]$form.ShowDialog()
$form.Dispose()
#endregion
