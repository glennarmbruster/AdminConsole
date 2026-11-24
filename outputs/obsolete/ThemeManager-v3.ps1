#Requires -Version 5.1
<#
    ThemeManager.ps1
    Edits Themes.json -- console colour palettes, tenant bindings, and logo. Author : Glenn Armbruster Revised : 2026-07-28

    Reasoning: ScriptNotes/ThemeManager.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-10
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
}
#endregion

#region Load -----------------------------------------------------------------
function Read-ThemeConfig {
    if (-not (Test-Path -LiteralPath $ThemeFile)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Themes.json was not found next to this script.`n`nExpected:`n$ThemeFile",
            'Theme Manager', 'OK', 'Error') | Out-Null
        return $null
    }
    try {
        return Get-Content -LiteralPath $ThemeFile -Raw | ConvertFrom-Json
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Themes.json could not be parsed.`n`n$($_.Exception.Message)",
            'Theme Manager', 'OK', 'Error') | Out-Null
        return $null
    }
}

$cfg = Read-ThemeConfig
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
# Moved rather than shortened: the text is 'Console uses: <theme name>' and a
# narrower label would start clipping tenant theme names instead.
$lblResolves.Location  = New-Object System.Drawing.Point(540, 13)
$lblResolves.Size      = New-Object System.Drawing.Size(270, 20)
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

    $added = @(); $skipped = @()
    foreach ($n in $shipped.Themes.PSObject.Properties.Name) {
        if ($cfg.Themes.PSObject.Properties.Name -contains $n) { $skipped += $n; continue }
        Add-Member -InputObject $cfg.Themes -NotePropertyName $n -NotePropertyValue $shipped.Themes.$n -Force
        $added += $n
    }
    if ($added.Count -eq 0) {
        $lblStatus.Text = "Nothing to add -- all $($skipped.Count) shipped theme(s) are already here."
        return
    }
    foreach ($n in $added) { if (-not $cmbTheme.Items.Contains($n)) { [void]$cmbTheme.Items.Add($n) } }
    $lblStatus.Text = "Added $($added.Count): $($added -join ', ')." +
        $(if ($skipped.Count) { " Left alone (already present): $($skipped -join ', ')." }) +
        ' Click Save to write them to Themes.json.'
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
    $name = "$($cfg.ActiveTheme)"
    $why  = 'active theme'
    if ($chkFollow.Checked) {
        $d = Get-ConnectedTenantDomain
        if ($d -and $cfg.TenantThemes.PSObject.Properties.Name -contains $d) {
            $name = "$($cfg.TenantThemes.$d)"
            $why  = "tenant binding for $d"
        }
    }
    $lblResolves.Text = "Console uses: $name"
    $lblResolves.ForeColor = if ($why -eq 'active theme') {
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

$cmbTheme.Add_SelectedIndexChanged({ Load-Theme -Name $cmbTheme.SelectedItem })
$chkEnabled.Add_CheckedChanged({ Update-Preview })
$chkFollow.Add_CheckedChanged({ Update-ResolvedLabel })

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
    $t    = Get-EditedTheme
    foreach ($p in $t.PSObject.Properties) {
        $cfg.Themes.$name | Add-Member -NotePropertyName $p.Name -NotePropertyValue $p.Value -Force
    }
    $cfg.FollowTenant = [bool]$chkFollow.Checked
    try {
        $cfg | ConvertTo-Json -Depth 8 |
            Set-Content -LiteralPath $ThemeFile -Encoding UTF8
    # Clear BOTH caches, including the globals child scripts read from, or an
    # already-open console keeps serving the previous theme.
        Clear-ConnectedTenantDomainCache
        $global:AdminConsoleTheme = $null
        Get-ConsoleTheme -Refresh | Out-Null
        $lblStatus.Text = "Saved. Active theme is '$($cfg.ActiveTheme)'. Reopen the Admin Console to see it."
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

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
try { Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Theme Manager' | Out-Null } catch { }

[void]$form.ShowDialog()
#endregion
