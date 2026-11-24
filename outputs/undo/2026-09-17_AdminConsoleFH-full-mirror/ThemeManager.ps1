#Requires -Version 5.1
<#
    ThemeManager.ps1
    Edits Themes.json -- console colour palettes, tenant bindings, and logo. Author : Glenn Armbruster Revised : 2026-07-28

    Reasoning: ScriptNotes/ThemeManager.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-09-16

    2026-09-16  'Set as active theme' followed by Save no longer bypasses the
                scope question. Apply changes $cfg in memory, so Save was
                comparing the new name against itself, seeing no change, and
                writing the shared ActiveTheme with no profile written. The
                comparison is against what was on DISK at load.
    2026-09-16  Per-user themes work on SINGLE-TENANT installs. The whole
                decision sat behind Test-MultiTenantMode, so on a one-tenant
                folder no profile theme was ever written and every save went
                into the shared ActiveTheme -- two operators overwrote each
                other's theme all day. Multi-tenant now only decides where
                'everyone' writes.
    2026-09-16  'Just me' no longer moves ActiveTheme in Themes.json. That value
                is SHARED and is the fallback for anyone whose profile names no
                theme, so a personal pick was changing other people's consoles --
                the scope question is now asked BEFORE the file is written, and
                the 'just me' answer puts the shared pointer back.
#>
# TOOLKIT-HELP-BEGIN
# The console's colours, and which theme each tenant gets.
#
# COLOUR IS NOT DECORATION HERE. One console folder per tenant, all looking
# identical, is how a change lands in the wrong tenant. A distinct theme per
# tenant means you can tell at a glance which one is in front of you, before
# you click anything.
#
# Add shipped themes brings in the bundled palettes. Pick one, Save, relaunch
# the console.
#
# TENANT BINDING ties a theme to a tenant so the right colours load
# automatically. The console reads Theme from Tenant.json; this is where that
# value comes from.
#
# Contrast matters more than taste. A palette where disabled text is nearly the
# same as enabled text hides the grey-out that tells you a tool cannot run --
# the shipped themes are checked for that, hand-rolled ones are not.
#
# The logo is cosmetic and optional. It is the one thing here that is.
#
# Themes.json is per-deployment like the rest of Config\ and is not shipped.
# TOOLKIT-HELP-END

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

#region Configuration --------------------------------------------------------
$ThemeFile = Get-ConfigFilePath -Name 'Themes.json'
$ImageDir  = Get-ConfigImageFolder

$Roles = [ordered]@{
    FormBack     = 'Window background'
    FormText     = 'Window text'
    PanelBack    = 'Panel background'
    GroupText    = 'Group heading text'
    ButtonBack   = 'Button background'
    ButtonText   = 'Button text'
    ButtonBorder = 'Button border'
    ButtonHover  = 'Button hover'
    ButtonDown   = 'Button pressed'
    InputBack    = 'Input background'
    InputText    = 'Input text'
    GradientTo   = 'Gradient end (blank = flat)'
    SelectedText = 'Selected radio/checkbox'
    ServiceNowBack = 'ServiceNow button'
    AccentBack   = 'Primary button (Save/Apply)'
    AccentText   = 'Primary button text'
    AccentHover  = 'Primary button hover'
    DangerColor  = 'Danger button (Delete/Remove)'
}
#endregion

#region Load -----------------------------------------------------------------
$script:IsFirstRun = $false

function Read-ThemeConfig {
    <#
        FIRST RUN IS NOT AN ERROR, AND USED TO BE TREATED AS ONE.

        Config\ ships with Themes.TEMPLATE.json and deliberately WITHOUT
        Themes.json -- the whole point of the folder boundary is that a
        deployment's own config is never overwritten by an upgrade (section
        31). So on a brand new tenant the file legitimately does not exist yet.

        This used to show "Themes.json was not found" and return, which is a
        dead end in the most literal sense: the button that would have created
        it -- Add shipped themes, further down this file -- lives on a form
        that never opens. The one tool whose job is to write Themes.json
        refused to start because Themes.json was missing.

        Tenant Manager already had the right shape for this: detect the first
        run, open anyway, say plainly what state you are in, and let Save
        create the file. This now matches it.

        SEEDED IN MEMORY, NOT WRITTEN HERE. Save writes it, the same as every
        other change in this window, so opening the tool to look around and
        closing it again leaves nothing behind.
    #>
    if (-not (Test-Path -LiteralPath $ThemeFile)) {
        $tpl = Get-ConfigFilePath -Name 'Themes.TEMPLATE.json'

        if ($tpl -and (Test-Path -LiteralPath $tpl)) {
            try {
                $seed = Get-Content -LiteralPath $tpl -Raw | ConvertFrom-Json
                $script:IsFirstRun = $true
                return $seed
            } catch {
                Show-ToolDialog(
                    ("Themes.json does not exist yet, and the template it would be built " +
                     "from could not be parsed.`n`n$tpl`n`n$($_.Exception.Message)"),
                    'Theme Manager', 'OK', 'Error') | Out-Null
                return $null
            }
        }

        # Neither file. Something removed the template, so there is genuinely
        # nothing to start from -- and the message says which file to restore
        # rather than naming the one that was never meant to ship.
        Show-ToolDialog(
            ("No themes are configured, and Themes.TEMPLATE.json is missing from Config\ " +
             "so there is nothing to build them from.`n`nExpected the template at:`n$tpl`n`n" +
             "Restore it from another copy of the toolkit, or from the deployment zip."),
            'Theme Manager', 'OK', 'Error') | Out-Null
        return $null
    }
    try {
        return Get-Content -LiteralPath $ThemeFile -Raw | ConvertFrom-Json
    } catch {
        Show-ToolDialog(
            "Themes.json could not be parsed.`n`n$($_.Exception.Message)",
            'Theme Manager', 'OK', 'Error') | Out-Null
        return $null
    }
}

$cfg = Read-ThemeConfig

# WHAT ActiveTheme SAYS ON DISK, CAPTURED BEFORE ANY BUTTON CAN CHANGE IT.
#
# 'Set as active theme' writes $cfg.ActiveTheme in MEMORY and tells you to
# click Save -- so by the time Save runs, $cfg already holds the new pick and
# cannot be asked what the old one was. Reading it there compared the new value
# with itself, concluded nothing had changed, and took the do-nothing path:
# no scope question, no profile written, and the new name saved into the
# SHARED Themes.json anyway. That is the bug this whole mechanism exists to
# prevent, arriving through the Apply button instead of the Save button.
#
# Updated after a save that genuinely writes ActiveTheme, so a second save in
# the same session compares against what is actually on disk.
$script:ActiveThemeOnDisk = "$($cfg.ActiveTheme)"

# AND WHAT THIS OPERATOR ACTUALLY RESOLVES TO, which is a different question.
#
# ActiveTheme is the SHARED fallback. Somebody with a theme in their own
# profile is not using it, and comparing against it is what made one theme
# unselectable: pick the theme that happens to equal ActiveTheme and the save
# concluded 'nothing changed', wrote nothing to the profile, and left the
# operator's old theme still winning. Every OTHER theme differed from the
# pointer, so every other theme worked -- which is precisely how it was
# reported: "I can set it to anything except Navy."
#
# Captured from disk BEFORE anything in this window edits $cfg, for the same
# reason ActiveThemeOnDisk is: Apply mutates $cfg in memory, so by save time
# the in-memory copy can no longer say what the operator started from.
$script:EffectiveThemeOnDisk = ''
try { $script:EffectiveThemeOnDisk = "$((Resolve-ConsoleThemeName -Config $cfg).Name)" } catch { }

# Set by Apply. Apply means 'make this the one I use', which is the same
# question Save asks -- so having pressed it, the operator is not asked again.
$script:PendingActive = ''
if (-not $cfg) { return }
$toolTipTM = New-Object System.Windows.Forms.ToolTip
$script:Current = $null
#endregion

#region Form -----------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Theme Manager'
$form.Size            = New-Object System.Drawing.Size(880, 706)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblTheme          = New-Object System.Windows.Forms.Label
$lblTheme.Text     = 'Theme:'
$lblTheme.Location = New-Object System.Drawing.Point(12, 15)
$lblTheme.Size     = New-Object System.Drawing.Size(50, 20)
$form.Controls.Add($lblTheme)

$cmbTheme               = New-Object System.Windows.Forms.ComboBox
$cmbTheme.DropDownStyle = 'DropDownList'
$cmbTheme.Location      = New-Object System.Drawing.Point(64, 12)
$cmbTheme.Size          = New-Object System.Drawing.Size(200, 22)
$form.Controls.Add($cmbTheme)

$chkEnabled          = New-Object System.Windows.Forms.CheckBox
$chkEnabled.Text     = 'Enabled'
$chkEnabled.Location = New-Object System.Drawing.Point(276, 13)
$chkEnabled.Size     = New-Object System.Drawing.Size(80, 20)
$form.Controls.Add($chkEnabled)

$chkFollow          = New-Object System.Windows.Forms.CheckBox
$chkFollow.Text     = 'Tenant binding overrides the active theme'
$chkFollow.Location = New-Object System.Drawing.Point(362, 13)
$chkFollow.Size     = New-Object System.Drawing.Size(220, 20)
$chkFollow.Checked  = [bool]$cfg.FollowTenant
$form.Controls.Add($chkFollow)

$lblDesc           = New-Object System.Windows.Forms.Label
$lblDesc.Location  = New-Object System.Drawing.Point(12, 40)
$lblDesc.Size      = New-Object System.Drawing.Size(560, 18)
$lblDesc.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblDesc)

$lblResolves           = New-Object System.Windows.Forms.Label
# Right edge at 810, not 858.
#
# Add-ToolHelpButton pins the '?' to the top-right corner at roughly x=824, and
# this label is RIGHT-aligned -- so its text grew towards the button and ended
# up underneath it. The clash only appeared once help was added to this tool,
# because until then nothing else lived in that corner.
#
# MOVED DOWN A ROW, having been moved sideways once already.
#
# The top row had run out: 'Tenant binding overrides the active' ends at x=582
# and this label started at 540 -- a 42px overlap -- while the help button owns
# the right-hand corner from 824. Sliding left collides with the checkbox and
# sliding right collides with the button, because the row is simply full.
#
# On its own line it clears both, and AutoEllipsis means a long theme name
# truncates visibly instead of disappearing under something.
#
# Original note, still true: the text is 'Console uses: <theme name>' and a
# narrower label would start clipping tenant theme names instead.
$lblResolves.Location  = New-Object System.Drawing.Point(590, 40)
$lblResolves.Size      = New-Object System.Drawing.Size(258, 20)
$lblResolves.AutoEllipsis = $true
$lblResolves.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblResolves.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblResolves)

# --- colour role rows ---
$grpColors          = New-Object System.Windows.Forms.GroupBox
$grpColors.Text     = 'Colours'
$grpColors.Location = New-Object System.Drawing.Point(12, 64)
$grpColors.Size     = New-Object System.Drawing.Size(420, 496)
$form.Controls.Add($grpColors)

$script:RoleControls = @{}
$y = 24
foreach ($key in $Roles.Keys) {
    $lbl          = New-Object System.Windows.Forms.Label
    $lbl.Text     = $Roles[$key]
    $lbl.Location = New-Object System.Drawing.Point(12, ($y + 4))
    $lbl.Size     = New-Object System.Drawing.Size(180, 20)
    $grpColors.Controls.Add($lbl)

    $swatch             = New-Object System.Windows.Forms.Panel
    $swatch.Location    = New-Object System.Drawing.Point(198, $y)
    $swatch.Size        = New-Object System.Drawing.Size(34, 24)
    $swatch.BorderStyle = 'FixedSingle'
    $grpColors.Controls.Add($swatch)

    $txt          = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(238, ($y + 1))
    $txt.Size     = New-Object System.Drawing.Size(80, 22)
    $grpColors.Controls.Add($txt)

    $btn          = New-Object System.Windows.Forms.Button
    $btn.Text     = 'Pick'
    $btn.Location = New-Object System.Drawing.Point(324, $y)
    $btn.Size     = New-Object System.Drawing.Size(50, 24)
    $btn.Tag      = $key
    $grpColors.Controls.Add($btn)

    $clr          = New-Object System.Windows.Forms.Button
    $clr.Text     = 'X'
    $clr.Location = New-Object System.Drawing.Point(378, $y)
    $clr.Size     = New-Object System.Drawing.Size(28, 24)
    $clr.Tag      = $key
    $grpColors.Controls.Add($clr)

    $script:RoleControls[$key] = @{ Swatch = $swatch; Text = $txt }
    $y += 33
}

# --- preview ---
$grpPrev          = New-Object System.Windows.Forms.GroupBox
$grpPrev.Text     = 'Live preview'
$grpPrev.Location = New-Object System.Drawing.Point(444, 64)
$grpPrev.Size     = New-Object System.Drawing.Size(404, 300)
$form.Controls.Add($grpPrev)

$prev             = New-Object System.Windows.Forms.Panel
$prev.Location    = New-Object System.Drawing.Point(12, 22)
$prev.Size        = New-Object System.Drawing.Size(380, 264)
$prev.BorderStyle = 'FixedSingle'
$grpPrev.Controls.Add($prev)

$pvTop          = New-Object System.Windows.Forms.Label
$pvTop.Text     = 'ExO: Connected   SPO: Connected      Tenant: contoso.com'
$pvTop.Location = New-Object System.Drawing.Point(10, 8)
$pvTop.Size     = New-Object System.Drawing.Size(360, 18)
$pvTop.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 8, [System.Drawing.FontStyle]::Bold)
$prev.Controls.Add($pvTop)

$pvGroup          = New-Object System.Windows.Forms.GroupBox
$pvGroup.Text     = 'Admin Tools'
$pvGroup.Location = New-Object System.Drawing.Point(10, 30)
$pvGroup.Size     = New-Object System.Drawing.Size(360, 96)
$prev.Controls.Add($pvGroup)

$px = 12
foreach ($n in @('New User', 'Disable', 'Health', 'Trace')) {
    $b          = New-Object System.Windows.Forms.Button
    $b.Text     = $n
    $b.Location = New-Object System.Drawing.Point($px, 26)
    $b.Size     = New-Object System.Drawing.Size(78, 28)
    $pvGroup.Controls.Add($b)
    $px += 84
}
$px = 12
foreach ($n in @('Perms', 'Rules', 'Calendar', 'Audit')) {
    $b          = New-Object System.Windows.Forms.Button
    $b.Text     = $n
    $b.Location = New-Object System.Drawing.Point($px, 58)
    $b.Size     = New-Object System.Drawing.Size(78, 28)
    $pvGroup.Controls.Add($b)
    $px += 84
}

$pvOut            = New-Object System.Windows.Forms.TextBox
$pvOut.Multiline  = $true
$pvOut.ReadOnly   = $true
$pvOut.ScrollBars = 'Vertical'
$pvOut.Location   = New-Object System.Drawing.Point(10, 132)
$pvOut.Size       = New-Object System.Drawing.Size(360, 118)
$pvOut.Text       = "Results appear here." + [Environment]::NewLine +
                    "----- jsmith  [Directory] -----" + [Environment]::NewLine +
                    "  Mailbox enabled." + [Environment]::NewLine +
                    "  Archive enabled."
$prev.Controls.Add($pvOut)

# --- logo ---
$grpLogo          = New-Object System.Windows.Forms.GroupBox
$grpLogo.Text     = 'Logo'
$grpLogo.Location = New-Object System.Drawing.Point(444, 372)
$grpLogo.Size     = New-Object System.Drawing.Size(404, 122)
$form.Controls.Add($grpLogo)

$lblLogoInfo           = New-Object System.Windows.Forms.Label
$lblLogoInfo.Location  = New-Object System.Drawing.Point(12, 20)
$lblLogoInfo.Size      = New-Object System.Drawing.Size(380, 32)
$lblLogoInfo.ForeColor = [System.Drawing.Color]::DimGray
$lblLogoInfo.Text      = 'Copied into Config\Images. A logo named for the connected tenant is preferred over the generic one. PNG keeps transparency.'
$grpLogo.Controls.Add($lblLogoInfo)

$txtLogo          = New-Object System.Windows.Forms.TextBox
$txtLogo.Location = New-Object System.Drawing.Point(12, 56)
$txtLogo.Size     = New-Object System.Drawing.Size(270, 22)
$txtLogo.ReadOnly = $true
$grpLogo.Controls.Add($txtLogo)

$btnLogo          = New-Object System.Windows.Forms.Button
$btnLogo.Text     = 'Choose...'
$btnLogo.Location = New-Object System.Drawing.Point(290, 55)
$btnLogo.Size     = New-Object System.Drawing.Size(100, 24)
$grpLogo.Controls.Add($btnLogo)

$lblLogoFound           = New-Object System.Windows.Forms.Label
$lblLogoFound.Location  = New-Object System.Drawing.Point(12, 86)
$lblLogoFound.Size      = New-Object System.Drawing.Size(380, 20)
$grpLogo.Controls.Add($lblLogoFound)

# --- buttons ---
$btnApply          = New-Object System.Windows.Forms.Button
$btnApply.Text     = 'Set as active theme'
$btnApply.Location = New-Object System.Drawing.Point(12, 572)
$btnApply.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnApply)

$btnSave          = New-Object System.Windows.Forms.Button
$btnSave.Text     = 'Save'
$btnSave.Location = New-Object System.Drawing.Point(170, 572)
$btnSave.Size     = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnSave)

# Adds the shipped palettes to THIS file without touching what is already here.
#
# The alternative was handing over a replacement Themes.json, which would have
# taken this deployment's own themes and its tenant bindings with it. Config\
# holds per-deployment data; a tool that improves it by overwriting it is not
# an improvement.
#
# Skips by NAME, so a local theme that happens to share a name with a shipped
# one is left alone -- yours is the one you have been looking at, and a tool
# that silently repaints your console has broken trust for a colour scheme.
$btnBuiltIn          = New-Object System.Windows.Forms.Button
$btnBuiltIn.Text     = 'Add shipped themes'
$btnBuiltIn.Location = New-Object System.Drawing.Point(390, 572)
$btnBuiltIn.Size     = New-Object System.Drawing.Size(150, 30)
$btnBuiltIn.Add_Click({
    $tpl = Get-ConfigFilePath -Name 'Themes.TEMPLATE.json'
    if (-not (Test-Path -LiteralPath $tpl)) {
        $lblStatus.Text = "Themes.TEMPLATE.json is not in Config\ -- nothing to add from."
        return
    }
    try { $shipped = Get-Content -LiteralPath $tpl -Raw | ConvertFrom-Json }
    catch { $lblStatus.Text = "Themes.TEMPLATE.json could not be parsed: $($_.Exception.Message)"; return }

    # A theme already here is left alone EXCEPT for keys it does not have
    # yet: a shipped theme that gains a colour (HeaderBack in 2026-09) would
    # otherwise never reach an install that added it earlier. Values the
    # operator has edited are never overwritten -- only absent keys are
    # filled in.
    $added = @(); $skipped = @(); $topped = @()
    foreach ($n in $shipped.Themes.PSObject.Properties.Name) {
        if ($cfg.Themes.PSObject.Properties.Name -contains $n) {
            $mine = $cfg.Themes.$n; $theirs = $shipped.Themes.$n; $new = 0
            foreach ($p in $theirs.PSObject.Properties) {
                if (-not $mine.PSObject.Properties[$p.Name]) { Add-Member -InputObject $mine -NotePropertyName $p.Name -NotePropertyValue $p.Value -Force; $new++ }
            }
            if ($new) { $topped += "$n (+$new)" } else { $skipped += $n }
            continue
        }
        Add-Member -InputObject $cfg.Themes -NotePropertyName $n -NotePropertyValue $shipped.Themes.$n -Force
        $added += $n
    }
    if ($added.Count -eq 0 -and $topped.Count -eq 0) {
        $lblStatus.Text = "Nothing to add -- all $($skipped.Count) shipped theme(s) are already here."
        return
    }
    foreach ($n in $added) { if (-not $cmbTheme.Items.Contains($n)) { [void]$cmbTheme.Items.Add($n) } }
    if ($script:Current -and $cmbTheme.SelectedItem -and ($topped -join ' ') -like "*$($cmbTheme.SelectedItem) (*") {
        # The theme on screen gained keys; reload it so Save keeps them.
        $script:Current = $cfg.Themes.$($cmbTheme.SelectedItem)
    }
    $lblStatus.Text = $(if ($added.Count) { "Added $($added.Count): $($added -join ', ')." } else { '' }) +
        $(if ($topped.Count) { " New keys filled in: $($topped -join ', ')." }) +
        $(if ($skipped.Count) { " Unchanged: $($skipped -join ', ')." }) +
        ' Click Save to write to Themes.json.'
})
$form.Controls.Add($btnBuiltIn)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(278, 572)
$btnClose.Size     = New-Object System.Drawing.Size(100, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$lblStatus           = New-Object System.Windows.Forms.Label
$lblStatus.Location  = New-Object System.Drawing.Point(12, 610)
$lblStatus.Size      = New-Object System.Drawing.Size(836, 40)
$lblStatus.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblStatus)
#endregion

#region Behaviour ------------------------------------------------------------
function Update-ResolvedLabel {
    <#
        WHAT THIS OPERATOR'S CONSOLE OPENS IN, and where that comes from.

        This used to read $cfg.ActiveTheme and the tenant binding, and nothing
        else -- so for anybody with a theme in their own profile it named a
        theme their console was not using, on every visit, in green, as a
        statement of fact. Now it asks the one function that decides.

        THE LABEL SAYS WHOSE ANSWER IT IS, because 'Navy' meaning 'the shared
        fallback, which does not apply to you' and 'Navy' meaning 'your own
        pick' look identical otherwise, and the difference is the whole reason
        per-user themes exist.
    #>
    $r = $null
    try { $r = Resolve-ConsoleThemeName -Config $cfg } catch { }
    if (-not $r) { $r = [PSCustomObject]@{ Name = ''; Source = ''; Detail = 'could not be resolved' } }

    $name = "$($r.Name)"
    $why  = "$($r.Detail)"
    # Apply has been pressed but not saved: show what Save WOULD do, or the
    # label contradicts the status line directly underneath it.
    if ($script:PendingActive) {
        $name = "$script:PendingActive"
        $why  = 'your pick, not saved yet'
    }

    # A NAME THAT WINS IS NOT THE SAME AS A THEME THAT LOADS, and this window
    # is the one place that distinction has to be visible.
    #
    # Resolve-ConsoleThemeName answers 'which name wins' and stops there --
    # deliberately, because Get-ConsoleTheme is what decides whether the name
    # is usable, and two functions both deciding that is how they drift. But it
    # means the winning name can be a theme that is missing from Themes.json or
    # has Enabled=false, and in BOTH of those cases the console runs UNTHEMED
    # while this label confidently names the theme. That is the exact shape of
    # the complaint this label already caused once -- 'it says that is the
    # current theme when it clearly is not' -- so it is checked here.
    #
    # Enabled=false is the quiet one. 'Default' ships that way, so it is a
    # plausible thing for ActiveTheme to be pointing at.
    $why = "$why"
    $usable = $true
    if ($name) {
        if (-not ($cfg.Themes -and ($cfg.Themes.PSObject.Properties.Name -contains $name))) {
            $usable = $false; $why = "'$name' is not in Themes.json -- the console runs UNTHEMED"
        } elseif (-not $cfg.Themes.$name.Enabled) {
            $usable = $false; $why = "'$name' is DISABLED -- the console runs UNTHEMED"
        }
    }

    $lblResolves.Text = $(if (-not $name) { "Console uses: no theme  ($why)" }
                          elseif (-not $usable) { "Console uses: NO THEME  ($why)" }
                          else { "Console uses: $name  ($why)" })
    # Red when nothing loads. Otherwise green only when it is THIS operator's
    # own answer, amber for anything inherited -- a shared fallback or a tenant
    # file -- because that is the case where the name on screen is not a choice
    # this person made.
    $lblResolves.ForeColor = if (-not $usable) {
        [System.Drawing.Color]::Firebrick
    } elseif ($script:PendingActive -or $r.Source -eq 'profile') {
        [System.Drawing.Color]::FromArgb(0,110,0)
    } else {
        [System.Drawing.Color]::FromArgb(180,90,0)
    }
    $toolTipTM.SetToolTip($lblResolves, "Chosen by: $why")
}

function Get-EditedTheme {
    $o = [ordered]@{
        Description = $script:Current.Description
        Enabled     = [bool]$chkEnabled.Checked
        Logo        = $txtLogo.Text
    }
    foreach ($k in $Roles.Keys) { $o[$k] = $script:RoleControls[$k].Text.Text.Trim() }
    # Keys this editor has no control for (FontFamily, MonoFamily, the
    # QuickLaunch colours) survive a save instead of being dropped.
    if ($script:Current) {
        foreach ($p in $script:Current.PSObject.Properties) {
            if (-not $o.Contains($p.Name)) { $o[$p.Name] = $p.Value }
        }
    }
    return [pscustomobject]$o
}

function Update-Preview {
    $t = Get-EditedTheme
    if (-not $t.Enabled) {
        $prev.BackColor = [System.Drawing.SystemColors]::Control
        $prev.ForeColor = [System.Drawing.SystemColors]::ControlText
        foreach ($c in $prev.Controls) {
            $c.BackColor = [System.Drawing.SystemColors]::Control
            $c.ForeColor = [System.Drawing.SystemColors]::ControlText
            foreach ($cc in $c.Controls) {
                $cc.BackColor = [System.Drawing.SystemColors]::Control
                $cc.ForeColor = [System.Drawing.SystemColors]::ControlText
                if ($cc -is [System.Windows.Forms.Button]) {
                    $cc.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
                    $cc.UseVisualStyleBackColor = $true
                }
            }
        }
        $pvOut.BackColor = [System.Drawing.SystemColors]::Window
        return
    }
    Set-ConsoleTheme -Form $prev -Theme $t -NoGradient
    $fb = ConvertTo-ThemeColor $t.FormBack
    if ($fb) { $prev.BackColor = $fb }
}

function Load-Theme {
    param([string]$Name)
    if (-not $Name -or $cfg.Themes.PSObject.Properties.Name -notcontains $Name) { return }
    $script:Current = $cfg.Themes.$Name
    $lblDesc.Text   = $script:Current.Description
    $chkEnabled.Checked = [bool]$script:Current.Enabled
    $txtLogo.Text   = "$($script:Current.Logo)"
    foreach ($k in $Roles.Keys) {
        $v = "$($script:Current.$k)"
        $script:RoleControls[$k].Text.Text = $v
        $c = ConvertTo-ThemeColor $v
        $script:RoleControls[$k].Swatch.BackColor =
            if ($c) { $c } else { [System.Drawing.SystemColors]::Control }
    }
    Update-Preview
}

foreach ($k in $Roles.Keys) {
    $script:RoleControls[$k].Text.Add_TextChanged({
        $key = ($Roles.Keys | Where-Object { $script:RoleControls[$_].Text -eq $this })
        if ($key) {
            $c = ConvertTo-ThemeColor $this.Text
            $script:RoleControls[$key].Swatch.BackColor =
                if ($c) { $c } else { [System.Drawing.SystemColors]::Control }
        }
        Update-Preview
    })
}

foreach ($c in $grpColors.Controls) {
    if ($c -is [System.Windows.Forms.Button] -and $c.Text -eq 'Pick') {
        $c.Add_Click({
            $key = $this.Tag
            $dlg = New-Object System.Windows.Forms.ColorDialog
            $dlg.FullOpen = $true
            $existing = ConvertTo-ThemeColor $script:RoleControls[$key].Text.Text
            if ($existing) { $dlg.Color = $existing }
            if ($dlg.ShowDialog() -eq 'OK') {
                $script:RoleControls[$key].Text.Text =
                    '#{0:X2}{1:X2}{2:X2}' -f $dlg.Color.R, $dlg.Color.G, $dlg.Color.B
            }
        })
    }
    if ($c -is [System.Windows.Forms.Button] -and $c.Text -eq 'X') {
        $c.Add_Click({ $script:RoleControls[$this.Tag].Text.Text = '' })
    }
}

$cmbTheme.Add_SelectedIndexChanged({
    # MOVING THE DROPDOWN CANCELS AN UNSAVED APPLY, and it has to, in memory as
    # well as on the label. Apply writes $cfg.ActiveTheme immediately; without
    # this, Apply-then-change-theme-then-Save carried the abandoned pick into
    # the save -- the label still claimed the old one in green, the switch
    # question quoted a third theme, and whichever way it was answered was
    # wrong. Moving the selector is the operator saying 'not that one'.
    if ($script:PendingActive -and $script:PendingActive -ne "$($cmbTheme.SelectedItem)") {
        $script:PendingActive = ''
        $cfg.ActiveTheme      = $script:ActiveThemeOnDisk
    }
    Load-Theme -Name $cmbTheme.SelectedItem
    Update-ResolvedLabel
})
$chkEnabled.Add_CheckedChanged({ Update-Preview })
# THE LABEL RESOLVES FROM $cfg.FollowTenant, NOT FROM THE BOX. Until this line
# the box was cosmetic between save and save: untick it and the label still
# named the tenant binding it had just been told to ignore. The save writes the
# same value again at the bottom; that is now belt and braces.
$chkFollow.Add_CheckedChanged({ $cfg.FollowTenant = [bool]$chkFollow.Checked; Update-ResolvedLabel })

$btnLogo.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Images (*.png;*.jpg;*.gif;*.bmp)|*.png;*.jpg;*.gif;*.bmp'
    $dlg.Title  = 'Choose a logo'
    if ($dlg.ShowDialog() -ne 'OK') { return }
    try {
        if (-not (Test-Path -LiteralPath $ImageDir)) {
            New-Item -Path $ImageDir -ItemType Directory -Force | Out-Null
        }
    # Copy, not reference: a path into someone's Desktop breaks when it moves.
        $name = [System.IO.Path]::GetFileName($dlg.FileName)
        $dest = Join-Path $ImageDir $name
        Copy-Item -LiteralPath $dlg.FileName -Destination $dest -Force
        $txtLogo.Text  = "Config\Images\$name"
        $lblStatus.Text = "Logo copied to $dest"
    } catch {
        $lblStatus.Text = "Could not copy the logo: $($_.Exception.Message)"
    }
})

$btnApply.Add_Click({
    if (-not $cmbTheme.SelectedItem) { return }
    $cfg.ActiveTheme = "$($cmbTheme.SelectedItem)"
    # Apply has already answered 'make this the one I use'. Save reads this
    # rather than asking the same question a second time.
    $script:PendingActive = "$($cmbTheme.SelectedItem)"
    $note = ''
    if ($chkFollow.Checked) {
        $d = Get-ConnectedTenantDomain
        if ($d -and $cfg.TenantThemes.PSObject.Properties.Name -contains $d) {
    # An explicit pick switches the binding override off, or it would win and
    # the choice would appear to do nothing.
            $chkFollow.Checked = $false
            $note = " Tenant binding switched off, or '$d' would have forced '$($cfg.TenantThemes.$d)'."
        }
    }
    Update-ResolvedLabel
    $lblStatus.Text = "Active theme set to '$($cfg.ActiveTheme)'.$note Click Save to write it to Themes.json."
})

$btnSave.Add_Click({
    if (-not $cmbTheme.SelectedItem) { return }
    $name = "$($cmbTheme.SelectedItem)"

    # WHAT ActiveTheme SAYS ON DISK. ActiveTheme is SHARED -- the last fallback
    # for anyone whose profile does not name a theme -- so writing it changes
    # everyone's console. The 'just me' path puts this value back before the
    # file is written, which is what makes 'just me' true. A plain alias for
    # $script:ActiveThemeOnDisk, which nothing changes until after the write.
    $origActive = "$($script:ActiveThemeOnDisk)"

    # SAVE DID NOT SET ActiveTheme, AND THAT IS WHY PICKING A THEME DID NOTHING.
    #
    # The dropdown chooses which theme is being EDITED; only Apply made it the
    # active one. So the obvious sequence -- pick a theme, press Save -- wrote
    # Themes.json with the PREVIOUS ActiveTheme still in it, reported the save
    # as successful, and left the console exactly as it was. A button called
    # Save that does not save the thing just selected is a trap, and this one
    # caught its own author twice. Reported 2026-09-14.
    #
    # ASKED, NOT ASSUMED. Making Save imply Apply would mean that opening
    # another theme to adjust one colour silently switches the console to it --
    # a different wrong answer. The two cases are indistinguishable from the
    # code, so the person who knows which one this is gets asked.
    #
    # ASKED AGAINST WHAT THIS OPERATOR USES, NOT AGAINST ActiveTheme. That
    # comparison was the bug: ActiveTheme is the SHARED fallback, so for
    # somebody with a theme of their own it is not what their console shows.
    # Selecting the theme that happened to equal ActiveTheme therefore read as
    # 'already active, nothing to do' -- no question, no profile write, old
    # theme still winning. One theme in the list, and only ever one, could not
    # be selected. Reported as "I can set the theme to anything except Navy".
    $effective  = "$($script:EffectiveThemeOnDisk)"
    $wantSwitch = $false
    # The switch question unticks Follow-tenant on Yes. The readability gate
    # below can still bail out AFTER that, and the box would stay unticked --
    # so the NEXT save, even a colours-only one, would write FollowTenant=false
    # and permanently disable a tenant binding on the strength of an answer the
    # operator abandoned.
    $followWas  = [bool]$chkFollow.Checked

    if ($script:PendingActive -eq $name) {
        # Apply was pressed on this theme. That IS the answer to this question;
        # asking again would be the window doubting a button it just obeyed.
        #
        # UNCONDITIONALLY TRUE, not ($name -ne $effective). Comparing NAMES
        # ignores where the name came from, and the two differ in ways that
        # matter: Apply also switches the tenant binding off, so applying the
        # very theme a binding was supplying looked like 'no change' -- and the
        # save then restored ActiveTheme, wrote FollowTenant=false, and flipped
        # the console to a theme nobody asked for. Every write below is
        # idempotent, so saying yes when it already matches costs nothing.
        $wantSwitch = $true
        $cfg.ActiveTheme = $name
    } elseif ($name -ne $effective) {
        $usesNow = $(if ($effective) { "'$effective'" } else { 'no theme' })
        $ans = Show-ToolDialog(
            ("'$name' is being edited. Your console currently uses $usesNow." + [Environment]::NewLine + [Environment]::NewLine +
             "Switch your console to '$name' as well?" + [Environment]::NewLine + [Environment]::NewLine +
             "Yes  -- save the colours AND switch to it." + [Environment]::NewLine +
             "No   -- save the colours only; your console keeps $usesNow."),
            'Active theme', 'YesNoCancel', 'Question')
        if ($ans -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
        if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
            $wantSwitch = $true
            $cfg.ActiveTheme = $name
            # Same reason Apply does this: a tenant binding left on would win
            # and make the explicit choice appear to do nothing.
            if ($chkFollow.Checked) { $chkFollow.Checked = $false }
        }
    }
    # No switch wanted: leave the shared pointer exactly as the file had it.
    # Without this, editing one colour of a theme nobody is using would move
    # everybody's fallback onto it.
    if (-not $wantSwitch) { $cfg.ActiveTheme = $script:ActiveThemeOnDisk }
    $t    = Get-EditedTheme

    # ---- readability gate ---------------------------------------------------
    #
    # CHECKED HERE BECAUSE THIS IS THE LAST MOMENT SOMEONE IS WATCHING.
    #
    # A palette whose text and background are too close produces a console that
    # cannot be read -- and every control for repairing it is inside that
    # console. Refusing to LOAD such a theme would be too late: the file is
    # written, the operator is locked out, and the message would appear in a
    # window they cannot see. At save time the person who made the change is
    # present, knows what they just did, and can undo it with one click.
    #
    # IT WARNS, IT DOES NOT REFUSE. A deliberate low-contrast choice is
    # legitimate -- a deliberately muted disabled state, or a brand colour that
    # is simply what the brand is -- and a tool that will not let an engineer
    # do a thing they meant to do gets worked around rather than obeyed. The
    # numbers are shown, the failing pairs are named, and 'Save anyway' is one
    # button away. Safe mode is the backstop for the case where they were wrong.
    $bad = @()
    try { $bad = @(Test-ThemeReadable -Theme $t) } catch { }
    if ($bad.Count) {
        $lines = @($bad | ForEach-Object { "  {0,-32} {1}:{2}  contrast {3}:1  ({4})" -f $_.Pair, $_.Fore, $_.Back, $_.Ratio, $_.Severity })
        $ans = Show-ToolDialog(
            ("This palette has text that will be hard or impossible to read:" +
             "`r`n`r`n" + ($lines -join "`r`n") +
             "`r`n`r`n4.5:1 is the usual minimum for body text. Below 3:1 the text is effectively invisible." +
             "`r`n`r`nIf you save this and the console becomes unusable, open it with:" +
             "`r`n    .\AdminConsole.ps1 -SafeMode" +
             "`r`nwhich ignores the theme for one run so you can come back here and fix it." +
             "`r`n`r`nSave anyway?"),
            'Readability', 'YesNo', 'Warning', 'Button2')
        if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) {
            # Put back what the switch question changed. Nothing was written,
            # so nothing about this window should look as if it was.
            $chkFollow.Checked = $followWas
            $lblStatus.Text = 'Not saved -- adjust the colours flagged above, or save again and confirm.'
            return
        }
    }

    foreach ($p in $t.PSObject.Properties) {
        $cfg.Themes.$name | Add-Member -NotePropertyName $p.Name -NotePropertyValue $p.Value -Force
    }
    $cfg.FollowTenant = [bool]$chkFollow.Checked

    # ---- whose theme is this? ASKED BEFORE ANYTHING IS WRITTEN ---------------
    #
    # This used to be asked AFTER Themes.json had already been written with the
    # new ActiveTheme in it, so the answer could not affect the shared file --
    # 'just me' wrote the profile as well and everyone else inherited the pick
    # anyway. The question has to come first, because one of its answers means
    # 'do not touch the shared pointer at all'.
    #
    # Two engineers on one folder means 'the tenant theme' is really 'the
    # tenant theme for whoever is at the keyboard'.
    #
    # ASKED ONCE, like Save Arrangement: once a profile exists it is the answer
    # from then on, because a dialog on every colour tweak is a dialog nobody
    # reads.
    # PER-USER IS NOT PER-TENANT, AND TYING THEM TOGETHER WAS THE BUG.
    #
    # This whole decision used to sit inside 'if (Test-MultiTenantMode)'. On a
    # SINGLE-TENANT install -- which is most of them -- that meant the question
    # was never asked, no profile Theme was ever written, and every theme save
    # went straight into ActiveTheme in the shared Themes.json. Two engineers
    # on one folder therefore shared one theme and overwrote each other, which
    # is exactly what per-user themes exist to stop. Reported from FH.
    #
    # Multi-tenant only decides where 'everyone' WRITES: the tenant file when
    # there is one, ActiveTheme when there is not. Whether the pick is personal
    # is a separate question, and it is asked either way.
    # WHAT WAS PICKED, AND WHETHER IT CHANGES ANYTHING FOR THIS PERSON.
    # $picked is the theme the operator selected -- not the shared pointer,
    # which may name something they have never used.
    $picked        = $name
    $activeChanged = $wantSwitch
    $multi = $false
    try { $multi = [bool](Test-MultiTenantMode) } catch { }
    $mine = $false

    if ($activeChanged) {
        # HAVE THEY CHOSEN A THEME BEFORE -- not 'do they have a profile
        # file'. Those were the same question until View Manager shipped: a
        # default view creates a profile, and so does nudging A+ inside a
        # view, so an operator who had done only that would have had their
        # first theme pick go silently to 'just me' without being offered
        # 'everyone'. Same correction as Save Arrangement and Tool Manager,
        # scoped to what THIS dialog writes.
        $mine = $false
        try { $mine = ($null -ne (Get-UserLayoutProperty -Name 'Theme')) } catch { }
        if (-not $mine) {
            $a = Show-ToolDialog(
                ("Apply '$picked' for YOU, or for everyone who uses this toolkit?" + "`r`n`r`n" +
                 "YES     just me" + "`r`n" +
                 "            (a profile is created; nobody else's console changes at all)" + "`r`n" +
                 $(if ($multi) {
                     "NO      everyone in $(if (Get-ActiveTenantLabel) { "'$(Get-ActiveTenantLabel)'" } else { 'this tenant' })" + "`r`n" +
                     "            (writes the tenant file AND the shared ActiveTheme)"
                 } else {
                     "NO      everyone who uses this toolkit folder" + "`r`n" +
                     "            (writes ActiveTheme in Themes.json, the shared file)"
                 }) + "`r`n`r`n" +
                 "The COLOURS are saved to Themes.json either way -- editing a palette is a" + "`r`n" +
                 "shared act. This is only about which theme is in use, and for whom."),
                'Whose theme', 'YesNoCancel', 'Question')
            if ($a -eq [System.Windows.Forms.DialogResult]::Cancel) {
                # Colours yes, active theme no -- which is exactly what the
                # old Cancel reported, except now it is also true.
                $cfg.ActiveTheme = $origActive
                # THE STATUS LINE IS INSIDE THE TRY. It used to sit after an
                # EMPTY catch, so a read-only Themes.json or a dropped share
                # told the operator the colours were saved when nothing had
                # been written -- the one silence the catch at the bottom of
                # this handler exists to prevent, reproduced here.
                try {
                    $cfg | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ThemeFile -Encoding UTF8
                    $global:AdminConsoleTheme = $null
                    $lblStatus.Text = "Colours saved to Themes.json. Nothing changed about which theme is in use."
                } catch {
                    $lblStatus.Text = "Save failed: $($_.Exception.Message)"
                }
                $script:PendingActive = ''
                try { Update-ResolvedLabel } catch { }
                return
            }
            $mine = ($a -eq [System.Windows.Forms.DialogResult]::Yes)
        }

        # THE LINE THAT MAKES 'JUST ME' MEAN IT. The shared pointer goes back
        # to what it was; the pick is carried in $picked and lands only in this
        # person's profile, below.
        if ($mine) { $cfg.ActiveTheme = $origActive }
    }

    try {
        $cfg | ConvertTo-Json -Depth 8 |
            Set-Content -LiteralPath $ThemeFile -Encoding UTF8
    # Clear BOTH caches, including the globals child scripts read from, or an
    # already-open console keeps serving the previous theme.
        Clear-ConnectedTenantDomainCache
        $global:AdminConsoleTheme = $null
        # THE WRITE HALF. Get-ConsoleTheme reads the ACTIVE TENANT's Theme
        # before ActiveTheme, so in multi-tenant mode the pick has to land in
        # the tenant file or it goes somewhere nothing consults -- and the
        # tenant's existing value keeps winning. Shipping the read without this
        # made every theme change appear to do nothing.
        #
        # Single-tenant installs are untouched: ActiveTheme IS the answer there,
        # and a second copy in Tenant.json would be two places to change.
        #
        # WHO gets it was decided above, before the file was written, and is
        # carried here in $mine and $picked. Deciding it here instead was the
        # bug: by this point the shared pointer had already moved.
        $tenantNote = ''
        try {
            if ($activeChanged) {
                if ($mine) {
                    # PLACEMENT AND PARKING COME ALONG UNTOUCHED. Save-UserLayout
                    # replaces the tenant's block whole, so the layout this
                    # person already has must be carried into the new one or a
                    # colour change would silently reset their arrangement.
                    $layout = Get-UserLayoutOwn
                    if (-not $layout) { $layout = [PSCustomObject]@{} }
                    Add-Member -InputObject $layout -NotePropertyName 'Theme' -NotePropertyValue $picked -Force
                    $up = Save-UserLayout -Layout $layout
                    $tenantNote = " Yours only -- written to $(Split-Path -Leaf $up). ActiveTheme in Themes.json is untouched, so nobody else's console changes."
                } elseif (-not $multi) {
                    # SINGLE TENANT: ActiveTheme was left holding the pick a few
                    # lines above, and that IS 'everyone' here. Nothing further
                    # to write -- say so rather than staying silent, because a
                    # save that reports nothing reads as a save that failed.
                    $tenantNote = " Written to Themes.json for everyone who uses this folder."
                    $script:ActiveThemeOnDisk = $picked
                } else {
                    $tp = Get-TenantConfigPath
                    if (Test-Path -LiteralPath $tp) {
                        # One writer for every per-tenant key, so the Parked
                        # block, the band and anything added later survive.
                        $null = Set-TenantFileKey -Name 'Theme' -Value $picked
                        $tenantNote = " Written to $(Split-Path -Leaf $tp) for everyone in this tenant."
                    } else {
                        # Multi-tenant mode with no tenant file to write to --
                        # started with no -Tenant, say. Say so rather than
                        # nothing: a save that reports nothing reads as a save
                        # that failed.
                        $tenantNote = " Written to Themes.json; there is no tenant file here to carry it."
                    }
                    # OUTSIDE THE Test-Path, because ActiveTheme went to disk a
                    # few lines above EITHER WAY. Leaving this inside meant the
                    # no-tenant-file case kept a stale ActiveThemeOnDisk, and
                    # the next colours-only save in the same session wrote the
                    # OLD theme back over the one just saved -- silently.
                    $script:ActiveThemeOnDisk = $picked
                }
            }
        } catch {
            # SAYING SO IS THE WHOLE VALUE OF THIS CATCH. A theme in Themes.json
            # but not in the tenant file is a HALF SAVE -- the console reopens in
            # the old colours and the save reported success. That exact silence
            # is what made this hard to diagnose the first time.
            $tenantNote = " WARNING: could not write it to the tenant file ($($_.Exception.Message)) -- the console will keep its previous theme."
        }

        Get-ConsoleTheme -Refresh | Out-Null

        # RE-READ WHAT THIS OPERATOR NOW RESOLVES TO, so a SECOND save in the
        # same session compares against the file as it is rather than as it was
        # when the window opened. Without this, saving twice hits the same
        # stale-comparison trap the whole fix is about -- the second save would
        # think nothing had changed and write nothing.
        try { $script:EffectiveThemeOnDisk = "$((Resolve-ConsoleThemeName -Config $cfg).Name)" } catch { }
        $script:PendingActive = ''
        try { Update-ResolvedLabel } catch { }

        $lblStatus.Text = $(if ($activeChanged) {
            "Saved. Your console will use '$picked'.$tenantNote Reopen the Admin Console to see it."
        } else {
            "Colours saved to Themes.json. Which theme is in use did not change.$tenantNote"
        })
    } catch {
        $lblStatus.Text = "Save failed: $($_.Exception.Message)"
    }
})
#endregion

#region Start ----------------------------------------------------------------
foreach ($n in $cfg.Themes.PSObject.Properties.Name) { [void]$cmbTheme.Items.Add($n) }
$idx = $cmbTheme.Items.IndexOf("$($cfg.ActiveTheme)")
$cmbTheme.SelectedIndex = if ($idx -ge 0) { $idx } else { 0 }

Update-ResolvedLabel

$logo = Get-TenantLogoPath
$lblLogoFound.Text = if ($logo) { "Currently resolving to: $(Split-Path $logo -Leaf)" }
                     else       { 'No logo found in Config\Images.' }
$lblLogoFound.ForeColor = if ($logo) { [System.Drawing.Color]::FromArgb(0,128,0) }
                          else       { [System.Drawing.Color]::DimGray }

# FIRST RUN SAYS SO, ON THE FACE OF THE WINDOW.
#
# The themes are loaded from the template and look completely normal, so
# without this the operator has no way to tell that nothing is saved yet --
# they would pick a theme, close the window, and find the console unchanged.
# Set last so nothing later overwrites the status line.
if ($script:IsFirstRun) {
    $form.Text = 'Theme Manager -- first run'
    $lblStatus.Text = ("No Themes.json yet, so these $(@($cfg.Themes.PSObject.Properties.Name).Count) " +
                       'themes came from the shipped template. Nothing is saved until you click Save.')
    $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
}

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
try { Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Theme Manager' | Out-Null } catch { }

[void]$form.ShowDialog()
#endregion
