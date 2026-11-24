#Requires -Version 5.1
<#
    AdminConsole.ps1
    Layer 1 launcher -- Mailbox / Group lookup + dynamic tool menu.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-RUNS: whatever you connect it to. The status row at the top of this window shows what it currently has.
# TOOLKIT-HELP-BEGIN
# The launcher. Look up a mailbox or group, then use the tools below.
#
# ---------------------------------------------------------------------------
# CONNECT FIRST
# ---------------------------------------------------------------------------
#
# Tools that need Exchange Online or SharePoint Online stay greyed out until
# this window has that connection. The status row at the top says which you
# have, and the connection belongs to THIS window.
#
# If your admin rights are activated on demand rather than standing, activate
# them BEFORE connecting. A session made beforehand holds the token it was
# issued with, and that token does not contain access granted afterwards --
# everything then fails as though the activation never happened. Activate,
# then reconnect.
#
# ---------------------------------------------------------------------------
# CHECK THE TENANT BEFORE YOU CHANGE ANYTHING
# ---------------------------------------------------------------------------
#
# This toolkit runs against several tenants and looks identical in all of
# them. The tenant name sits in the top row, and a colour scheme can be bound
# per tenant in ThemeManager -- colour tells you where you are faster than
# reading does.
#
# A (?) beside the tenant name means the connected domain does not match what
# the settings expect. That is a question rather than a fault -- a tenant with
# several accepted domains can disagree legitimately -- but it is worth
# answering before a write.
#
# ---------------------------------------------------------------------------
# WHY SOME BUTTONS ARE GREYED OUT
# ---------------------------------------------------------------------------
#
# Two separate reasons, and the difference matters:
#
#   * WRONG OBJECT TYPE. Some tools only apply to a mailbox, some only to a
#     group. After a lookup, anything that cannot act on what you found is
#     greyed. Look up something else and it comes back.
#
#   * NOT CONNECTED. A tool that needs a service you have not connected is
#     greyed until you connect it.
#
# Hover a greyed button and the tooltip says which of the two it is.
#
# ---------------------------------------------------------------------------
#
# The line at the bottom shows the toolkit version and whether the files still
# match it. Amber means this copy differs from the one it was stamped from --
# run Tester to see which files.
#
# Every tool has its own ? button in its top-right corner explaining what it
# does and which services it touches.
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

# One-time tidy-up for installs that predate Config\. Runs before anything reads
# a config file, and reports what it did rather than moving files silently.
$script:ConfigMigrationNotes = @(Invoke-ConfigMigration)

#region Tools.json load -------------------------------------------------------
function Get-ToolDefinitions {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Tools.json not found at '$Path'. Use ToolManager.ps1 to create one, or copy the seed Tools.json shipped with this console."
    }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        $tools = @($raw | ConvertFrom-Json -ErrorAction Stop)
    } catch {
        throw "Tools.json at '$Path' could not be parsed as JSON: $($_.Exception.Message)"
    }
    return @($tools | Sort-Object { if ($_.PSObject.Properties['Order']) { [int]$_.Order } else { 999 } })
}
#endregion

#region Main window ------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Admin Console"
$form.Size = New-Object System.Drawing.Size(840, 808)
$form.StartPosition = "Manual"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblExo = New-Object System.Windows.Forms.Label
$lblExo.Tag = 'NoTheme'
$lblExo.Location = New-Object System.Drawing.Point(12, 8)
$lblExo.Size = New-Object System.Drawing.Size(150, 20)
$lblExo.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblExo.Text = "ExO: ..."
$form.Controls.Add($lblExo)

$lblSpo = New-Object System.Windows.Forms.Label
$lblSpo.Tag = 'NoTheme'
$lblSpo.Location = New-Object System.Drawing.Point(170, 8)
$lblSpo.Size = New-Object System.Drawing.Size(150, 20)
$lblSpo.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblSpo.Text = "SPO: ..."
$form.Controls.Add($lblSpo)

$lblTenant = New-Object System.Windows.Forms.Label
$lblTenant.Tag = 'NoTheme'
$lblTenant.Location = New-Object System.Drawing.Point(330, 8)
$lblTenant.Size = New-Object System.Drawing.Size(330, 20)
$lblTenant.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblTenant.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblTenant)

$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Tag = 'NoTheme'
$lblWarn.Text = "If not already connected use the connect scripts below to connect."
$lblWarn.Location = New-Object System.Drawing.Point(12, 30)
# 480, not 800: the PIM countdown sits to its right at x=500. Measured against
# the label text, not guessed -- see ARCHITECTURE 14l.
$lblWarn.Size = New-Object System.Drawing.Size(480, 18)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Italic)
$lblWarn.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblWarn)

# --- PIM countdown ---------------------------------------------------------
#
# Created only when Settings.json says this tenant uses on-demand admin rights.
# Where rights are standing, a countdown is a clock that is always wrong, and a
# wrong clock beside a live connection status invites belief. Toggle it in
# ToolManager > Settings.
#
# The value is LOCAL: it starts when the Activate button records it and is not
# read back from the directory. Reading it back needs Graph consent many
# tenants will not give.
$script:PimTimerOn = $false
try { $script:PimTimerOn = Test-PimTimerEnabled } catch { $script:PimTimerOn = $false }

$lblPim = New-Object System.Windows.Forms.Label
$lblPim.Tag = 'NoTheme'
$lblPim.Location = New-Object System.Drawing.Point(500, 30)
$lblPim.Size = New-Object System.Drawing.Size(312, 18)
$lblPim.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblPim.Visible = $false
$form.Controls.Add($lblPim)

function Update-PimLabel {
    if (-not $script:PimTimerOn) { $lblPim.Visible = $false; return }
    try {
        $line = Get-PimStatusLine
        if (-not $line) { $lblPim.Visible = $false; return }
        $lblPim.Text      = $line
        $lblPim.ForeColor = Get-PimStatusColour
        $lblPim.Visible   = $true
    } catch { $lblPim.Visible = $false }
}

function Update-TenantLabel {
    <#
        The label now comes from configuration or from the session's own
        record, never from enumerating accepted domains -- see
        Get-ConnectedTenantDomain. That makes it free to call, which matters
        because Update-ConnectionButtonColors clears the cache and calls it
        again on every repaint.

        Because it can now report a CONFIGURED value rather than a verified
        one, the mismatch check is what keeps it honest: connected to the wrong
        tenant, the label alone would look perfectly correct.
    #>
    $d = Get-ConnectedTenantDomain
    if ($d) {
        $lblTenant.Text = "Tenant: $d"
        $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(0, 128, 0)
    } else {
        $lblTenant.Text = "Tenant: not connected"
        $lblTenant.ForeColor = [System.Drawing.Color]::DimGray
    }

    # Amber, not red, and it does not block anything. A tenant with several
    # accepted domains can disagree with its onmicrosoft prefix legitimately,
    # so this is a question worth asking rather than a fault worth stopping on.
    $mismatch = Test-TenantDomainMismatch
    if ($mismatch) {
        $lblTenant.Text = "$($lblTenant.Text)  (?)"
        $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(176, 96, 0)
        $toolTip.SetToolTip($lblTenant, $mismatch)
    } else {
        $toolTip.SetToolTip($lblTenant, '')
    }
}

$lblInput = New-Object System.Windows.Forms.Label
$domainHint = Get-MailDomain
$lblInput.Text = if ($domainHint) {
    "Enter mailbox or group email address (@$domainHint optional):"
} else {
    'Enter mailbox or group email address:'
}
$lblInput.Location = New-Object System.Drawing.Point(12, 50)
$lblInput.AutoSize = $true
$form.Controls.Add($lblInput)
$txtInput = New-Object System.Windows.Forms.TextBox
$txtInput.Location = New-Object System.Drawing.Point(12, 70)
$txtInput.Size = New-Object System.Drawing.Size(800, 20)
$txtInput.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtInput)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Look Up"
$btnRun.Location = New-Object System.Drawing.Point(12, 98)
$btnRun.Size = New-Object System.Drawing.Size(110, 28)
$form.Controls.Add($btnRun)
$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear"
$btnClear.Location = New-Object System.Drawing.Point(130, 98)
$btnClear.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClear)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(228, 98)
$btnClose.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClose)
$form.CancelButton = $btnClose

$lblResult = New-Object System.Windows.Forms.Label
$lblResult.Text = "Not looked up yet."
$lblResult.Location = New-Object System.Drawing.Point(12, 136)
$lblResult.Size = New-Object System.Drawing.Size(800, 20)
$lblResult.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblResult)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = "Vertical"
$txtOut.Location = New-Object System.Drawing.Point(12, 160)
$txtOut.Size = New-Object System.Drawing.Size(800, 140)
$txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtOut.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.Controls.Add($txtOut)

# Sections are built from Sections.json rather than hardcoded, so a new
# section can be added without touching this script. Geometry is derived:
# every section is ButtonsPerRow buttons wide and as many rows as it declares,
# and the form grows or shrinks to fit whatever total that comes to.
$script:SectionsConfigPath = Get-ConfigFilePath -Name 'Sections.json'
$script:BtnW = 150; $script:BtnH = 32; $script:BtnPad = 4
$script:PanelW = 800

function Get-SectionDefinitions {
    $default = [PSCustomObject]@{
        ButtonsPerRow = 5
        Sections = @(
                        [PSCustomObject]@{ Name = 'Admin Tools';     Rows = 4 }
            [PSCustomObject]@{ Name = 'Daily Scripts';   Rows = 1 }
        )
    }
    if (-not (Test-Path -LiteralPath $script:SectionsConfigPath)) { return $default }
    try {
        $cfg = Get-Content -LiteralPath $script:SectionsConfigPath -Raw | ConvertFrom-Json
        if (-not $cfg.Sections -or @($cfg.Sections).Count -eq 0) { return $default }
        if (-not $cfg.ButtonsPerRow) {
            Add-Member -InputObject $cfg -NotePropertyName ButtonsPerRow -NotePropertyValue 5 -Force
        }
        return $cfg
    } catch { return $default }
}

$script:SectionCfg   = Get-SectionDefinitions
$script:SectionPanels = @{}
$script:SectionNames  = @(@($script:SectionCfg.Sections) | ForEach-Object { "$($_.Name)" })

$y = 308
foreach ($sec in @($script:SectionCfg.Sections)) {
    $rows = [int]$sec.Rows
    if ($rows -lt 1) { $rows = 1 }
    $panelH = ($rows * ($script:BtnH + ($script:BtnPad * 2))) + 12

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "$($sec.Name)"
    $lbl.Location = New-Object System.Drawing.Point(12, $y)
    $lbl.Size = New-Object System.Drawing.Size($script:PanelW, 20)
    $form.Controls.Add($lbl)

    $flow = New-Object System.Windows.Forms.FlowLayoutPanel
    $flow.Location = New-Object System.Drawing.Point(12, ($y + 22))
    $flow.Size = New-Object System.Drawing.Size($script:PanelW, $panelH)
    $flow.AutoScroll = $true
    $flow.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $form.Controls.Add($flow)

    $script:SectionPanels["$($sec.Name)"] = $flow
    $y += 22 + $panelH + 8
}

# Grow the window to whatever the sections actually needed.
# The tooltip provider is created here rather than further down because the
# version footer below needs it. It was originally below, which meant the
# footer called SetToolTip on a null and threw -- and the catch turned that
# ordering mistake into "version unavailable" on every launch. A catch that
# reports the wrong cause is worse than no catch.
$toolTip = New-Object System.Windows.Forms.ToolTip

# Version footer, below the last section panel. Added before ClientSize is
# fixed so the form grows to include it rather than clipping it.
#
# Shows the version AND whether the files still match it -- because the number
# alone is a claim. Green when everything matches, amber when files differ or
# the manifest is stale, grey when the copy was never stamped.
# Width is the form's, less both margins, so the line has the whole footer to
# work with rather than the 600px it started with. AutoEllipsis makes any
# remaining overflow end in "..." -- a visible signal that there is more to
# read, instead of a sentence that silently stops mid-word.
$lblVersion              = New-Object System.Windows.Forms.Label
$lblVersion.Tag          = 'NoTheme'
$lblVersion.Location     = New-Object System.Drawing.Point(12, ($y + 4))
$lblVersion.Size         = New-Object System.Drawing.Size(800, 18)
$lblVersion.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif', 8)
$lblVersion.AutoEllipsis = $true
$form.Controls.Add($lblVersion)
$y += 26

try {
    $vv = Test-ToolkitIntegrity
    if (-not $vv.HasManifest) {
        $lblVersion.Text      = 'Admin Toolkit -- unversioned copy (no Version.json)'
        $lblVersion.ForeColor = [System.Drawing.Color]::DimGray
    } elseif ($vv.Modified.Count -eq 0 -and $vv.Missing.Count -eq 0 -and $vv.Extra.Count -eq 0) {
        $lblVersion.Text      = "Admin Toolkit v$($vv.Version)  --  $($vv.Matched) files verified  --  released $($vv.Released)"
        $lblVersion.ForeColor = [System.Drawing.Color]::FromArgb(80, 220, 130)
    } else {
        # Built from the parts rather than reusing $vv.Summary. Summary already
        # opens with "v1.03 --", so pasting it after the name printed the
        # version twice -- and it is worded for the Tester report, which has
        # room to spare. This line does not.
        #
        # Only non-zero categories are named. "1 modified" says the same thing
        # as "1 modified, 0 missing, 0 extra" in a third of the width, and the
        # zeroes were the padding that pushed this off the edge.
        $bits = @()
        if ($vv.Modified.Count -gt 0) { $bits += "$($vv.Modified.Count) modified" }
        if ($vv.Missing.Count  -gt 0) { $bits += "$($vv.Missing.Count) missing" }
        if ($vv.Extra.Count    -gt 0) { $bits += "$($vv.Extra.Count) extra" }
        $detail = if ($bits.Count -gt 0) { $bits -join ', ' } else { 'differs from the manifest' }
        if ($vv.Stale) { $detail += ' (manifest older -- re-stamp)' }

        $lblVersion.Text      = "Admin Toolkit v$($vv.Version)  --  $detail  --  run Tester for the list"
        $lblVersion.ForeColor = [System.Drawing.Color]::FromArgb(230, 170, 60)
    }
} catch {
    $lblVersion.Text      = 'Admin Toolkit -- version unavailable'
    $lblVersion.ForeColor = [System.Drawing.Color]::DimGray
}

# Outside the try deliberately: a tooltip that fails to attach must not be able
# to overwrite a version line that was worked out correctly.
$toolTip.SetToolTip($lblVersion, "Version and per-file SHA-256 come from Version.json.`r`nRun Tester to see exactly which files differ from the master.")

$form.ClientSize = New-Object System.Drawing.Size(824, ($y + 8))
#endregion

#region State + lookup ----------------------------------------------------------
$script:LastSmtp = $null
$script:LastType = $null
$script:LastADServer = $null
$script:AllTools = @()

try {
    $script:AllTools = Get-ToolDefinitions -Path $ToolsConfigPath
} catch {
    $script:AllTools = @()
    $txtOut.Text = $_.Exception.Message
}


function Start-ToolScript {
    param($Tool)
    $scriptPath = Join-Path $PSScriptRoot $Tool.Script
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "'$($Tool.Name)' points to '$($Tool.Script)', which does not exist under this console's folder.`n`nThis is expected if this is a reduced-tier distribution that doesn't include every tool -- otherwise, check Tools.json (ToolManager.ps1) for a stale/typo'd path.",
            'Tool Script Not Found', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    try {
        & $scriptPath -SmtpAddress $script:LastSmtp -DomainController $script:LastADServer
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context "Launching $($Tool.Name)"
    }
}

$script:KnownObjectTypes = @('Mailbox', 'UnifiedGroup', 'OnpremGroup')

function Get-ToolSection {
    param($Tool)
    $s = "$($Tool.Section)".Trim()
    foreach ($n in $script:SectionNames) {
        if ($s -eq $n) { return $n }
    }
    # Unknown or blank section -- fall back to the first declared section so a
    # tool can never become invisible just because its section was renamed.
    return $script:SectionNames[0]
}

function Add-ToolButtonsToPanel {
    param($Panel, $Tools, [string]$EmptyText)
    if (@($Tools).Count -eq 0) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $EmptyText
        $lbl.AutoSize = $true
        $Panel.Controls.Add($lbl)
        return
    }
    foreach ($tool in $Tools) {
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = $tool.Name
        $btn.Size = New-Object System.Drawing.Size(150, 32)
        $btn.Margin = New-Object System.Windows.Forms.Padding(4)
        $btn.Tag = $tool
        $btn.Add_Click({
            Start-ToolScript -Tool $this.Tag
            # Tools run in THIS process via '&', so a tool that connected to
            # something has changed what every other tool can do. With no
            # Connect Scripts section any more, the moment a tool returns is the
            # reliable re-probe point: connect inside console.ps1, close it, and
            # the buttons update. Full probe including SharePoint -- a tool
            # finishing is infrequent and user-initiated, so the cost is fine
            # here in a way it would not be on every Look Up.
            Update-ConnectionButtonColors
        })
        $Panel.Controls.Add($btn)
    }
}

function Build-ToolButtons {
    foreach ($name in $script:SectionNames) {
        $panel = $script:SectionPanels[$name]
        if (-not $panel) { continue }
        $panel.Controls.Clear()
        $tools = @($script:AllTools | Where-Object { (Get-ToolSection $_) -eq $name })
        Add-ToolButtonsToPanel -Panel $panel -Tools $tools `
            -EmptyText "No tools in '$name'. Use ToolManager.ps1 to add some."
    }
}

#region NEW 2026-07-24 -- colour the Connect Scripts buttons by live state ----
$script:ConnectButtonServices = @(
    [PSCustomObject]@{ Service = 'ExO'; Pattern = 'exo|exchange' }
    [PSCustomObject]@{ Service = 'SPO'; Pattern = 'spo|sharepoint' }
)

$script:ConnClrOk   = [System.Drawing.Color]::FromArgb(198, 239, 206)
$script:ConnClrBad  = [System.Drawing.Color]::FromArgb(255, 199, 206)
$script:ConnClrNone = $null   # $null = leave the themed colour alone
$script:ConnTextClr = [System.Drawing.Color]::FromArgb(20, 20, 20)

function Get-ConnectButtonService {
    param($Tool)
    $name = "$($Tool.Name)"
    foreach ($svc in $script:ConnectButtonServices) {
        if ($name -imatch $svc.Pattern) { return $svc.Service }
    }
    return $null
}

function Update-ConnectionButtonColors {
    param([switch]$SkipSpo)

    $prevCursor = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Clear-ConnectedTenantDomainCache
        Update-TenantLabel

        # Read the capability map rather than probing again. These labels and the
        # tool grey-out have to agree: "ExO: Connected" beside a button greyed
        # for a missing Exchange session would be the console contradicting
        # itself. One probe, one answer, both consumers.
        $map = Get-CapabilityStateMap -Refresh -SkipSpo:$SkipSpo
        $exo = $map['ExO']
        $spo = if ($map['SPO'].Unknown) { $null } else { $map['SPO'] }

        # Only the real Connect Scripts panel gets status colouring. The old
        # fallback to the first section painted Admin Tools buttons green/red if
        # a tool's NAME happened to match 'exchange' or 'sharepoint' -- harmless
        # with today's names, a mystery bug the first time one is renamed.
        $connPanel = $script:SectionPanels['Connect Scripts']
        foreach ($ctrl in @(if ($connPanel) { $connPanel.Controls } else { @() })) {
            if (-not $ctrl.Tag) { continue }
            $svc = Get-ConnectButtonService -Tool $ctrl.Tag
            if (-not $svc) {
                if ($script:ConnClrNone) { $ctrl.BackColor = $script:ConnClrNone }
                continue
            }
            $state = switch ($svc) { 'ExO' { $exo } 'SPO' { $spo } }
            if (-not $state) { continue }

            $ctrl.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $ctrl.BackColor = if ($state.Available) { $script:ConnClrOk } else { $script:ConnClrBad }
            # Pale status colours need dark text. Without this the button keeps
            # the theme's white ForeColor and reads as disabled.
            $ctrl.ForeColor = $script:ConnTextClr
            $toolTip.SetToolTip($ctrl, "$svc -- $($state.Detail)")
        }

        $green = [System.Drawing.Color]::FromArgb(0, 128, 0)
        $red   = [System.Drawing.Color]::FromArgb(200, 0, 0)

        if ($exo.Available) {
            $lblExo.Text = 'ExO: Connected';     $lblExo.ForeColor = $green
        } else {
            $lblExo.Text = 'ExO: Not Connected'; $lblExo.ForeColor = $red
        }
        $toolTip.SetToolTip($lblExo, "$($exo.Detail)")

        if (-not $spo) {
            $lblSpo.Text = 'SPO: not checked'
            $lblSpo.ForeColor = [System.Drawing.Color]::DimGray
            $toolTip.SetToolTip($lblSpo, 'Not checked yet -- runs after any tool closes.')
        } elseif ($spo.Available) {
            $lblSpo.Text = 'SPO: Connected';     $lblSpo.ForeColor = $green
            $toolTip.SetToolTip($lblSpo, "$($spo.Detail)")
        } else {
            $lblSpo.Text = 'SPO: Not Connected'; $lblSpo.ForeColor = $red
            $toolTip.SetToolTip($lblSpo, "$($spo.Detail)")
        }
        $lblWarn.Visible = -not ($exo.Available -and $spo -and $spo.Available)

        # Same probe, same answer: repaint the tool buttons from the map that
        # just produced these labels.
        Update-CapabilityState
    } catch {
        $lblExo.Text = 'ExO: unavailable'
        $lblExo.ForeColor = [System.Drawing.Color]::DimGray
        $lblSpo.Text = 'SPO: unavailable'
        $lblSpo.ForeColor = [System.Drawing.Color]::DimGray
        # Probe failed -- show the prompt, since we cannot confirm either one.
        $lblWarn.Visible = $true
    } finally {
        $form.Cursor = $prevCursor
    }
}
#endregion

#region NEW 2026-07-29 -- grey out tools this window cannot run ---------------
<#
    Two independent reasons a button can be disabled, and they must be decided
    together in ONE place.

    The object-type grey-out below used to end with an unconditional
    $ctrl.Enabled = $true in its else branch. Adding capability greying as a
    separate second pass would mean whichever ran last won, and a Look Up would
    silently re-enable a button whose Exchange session does not exist. So both
    reasons are evaluated here and capability is reported first, because it is
    the more fundamental block: the wrong object type means "not for this
    target", a missing connection means "not at all".
#>
$script:ToolAvail = @{}

function Update-CapabilityState {
    <#
        Recomputes per-tool availability from the CURRENT capability map and
        repaints the buttons. Does not probe -- Update-ConnectionButtonColors
        owns the probing, so the labels and the grey-out can never be computed
        from two different answers.
    #>
    try {
        $map = Get-CapabilityStateMap
        $script:ToolAvail = @{}
        foreach ($tool in @($script:AllTools)) {
            $path = Join-Path $PSScriptRoot $tool.Script
            $script:ToolAvail["$($tool.Name)"] = Get-ToolAvailability -ScriptPath $path -StateMap $map
        }
    } catch {
        # A failed classification must not disable anything. An empty map means
        # every tool falls back to enabled -- the same fail-open rule that keeps
        # an unreadable script from being blocked.
        $script:ToolAvail = @{}
    }
    Update-ToolButtonState -DetectedType $script:LastType
}

function Get-ToolCapabilityBlock {
    param($Tool)
    $a = $script:ToolAvail["$($Tool.Name)"]
    if (-not $a) { return $null }        # never classified: fail open
    if ($a.Enabled) { return $null }
    return $a
}
#endregion

function Update-ToolButtonState {
    param([string]$DetectedType)
    $showDimming = ($DetectedType -and ($script:KnownObjectTypes -contains $DetectedType))
    $allCtrls = @()
    foreach ($n in $script:SectionNames) {
        $p = $script:SectionPanels[$n]
        if ($p) { $allCtrls += @($p.Controls) }
    }
    foreach ($ctrl in $allCtrls) {
        if (-not $ctrl.Tag) { continue }
        $tool = $ctrl.Tag

        $capBlock = Get-ToolCapabilityBlock -Tool $tool
        $typeOk   = -not ($showDimming -and -not (@($tool.ObjectTypes) -contains $DetectedType))

        if ($capBlock) {
            $ctrl.Enabled = $false
            $tip = "'$($tool.Name)' cannot run in this window." + "`r`n`r`n" + $capBlock.Reason
            if (-not $typeOk) {
                $tip += "`r`n`r`nIt also applies only to: $(@($tool.ObjectTypes) -join ', ') -- the last lookup resolved to '$DetectedType'."
            }
            $toolTip.SetToolTip($ctrl, $tip)
        } elseif (-not $typeOk) {
            $ctrl.Enabled = $false
            $toolTip.SetToolTip($ctrl, "'$($tool.Name)' applies to: $(@($tool.ObjectTypes) -join ', '). The last lookup resolved to '$DetectedType', so this is disabled for it. Look up a matching object to enable it -- or clear the input box and re-run Look Up with nothing entered to release the grey-out on every tool.")
        } else {
            $ctrl.Enabled = $true
            $a = $script:ToolAvail["$($tool.Name)"]
            $toolTip.SetToolTip($ctrl, $(if ($a) { $a.Reason } else { '' }))
        }
    }
}

$btnRun.Add_Click({
    $btnRun.Enabled = $false
    $txtOut.Text = ''
    $lblResult.Text = 'Looking up...'
    try {
        $smtp = Get-TargetSmtp -Text $txtInput.Text
        if (-not $smtp) {
            $script:LastSmtp = $null
            $script:LastADServer = $null
            $script:LastType = $null
            $lblResult.Text = 'Enter an email address first. (Tools grey-out has been cleared.)'
            Update-ToolButtonState -DetectedType $null
            return
        }
        $lblResult.Text = 'Looking up in the directory...'
        $hit = Find-ADObjectBySmtp -Smtp $smtp
        if (-not $hit.Found) {
            $script:LastSmtp = $null
            $script:LastADServer = $null
            $script:LastType = $null
            # A directory that refused to answer and an address that does not
            # exist are different problems. Reporting the wrong one sends you to
            # check your typing when the network is at fault.
            $lblResult.Text = if ($hit.Errors.Count -gt 0) {
                "Directory lookup failed -- $smtp"
            } else {
                "Not found in the directory -- $smtp"
            }
            if ($hit.Errors.Count -gt 0) {
                $txtOut.Text = ($hit.Errors -join "`r`n`r`n")
            }
            Update-ToolButtonState -DetectedType $null
            return
        }
        $script:LastADServer = Get-DomainController
        $target = Resolve-Target -Smtp $smtp -ADServer $script:LastADServer
        $script:LastSmtp = $smtp
        $script:LastType = $target.Type
        switch ($target.Type) {
            'Mailbox'      { $lblResult.Text = "Found: Mailbox -- $smtp" }
            'UnifiedGroup' { $lblResult.Text = "Found: Microsoft 365 Group -- $smtp" }
            'OnpremGroup'  { $lblResult.Text = "Found: On-premises Group -- $smtp" }
            default {
                $lblResult.Text = "Directory object found but not confirmed as a Mailbox, Microsoft 365 Group, or on-premises group -- $smtp"
                if ($target.ResolutionErrors.Count -gt 0) {
                    $txtOut.Text = ($target.ResolutionErrors -join "`r`n`r`n")
                }
            }
        }
        Update-ToolButtonState -DetectedType $target.Type
    } catch {
        $lblResult.Text = 'Lookup failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Admin Console lookup'
        Update-ToolButtonState -DetectedType $null
    } finally {
        $btnRun.Enabled = $true
    }
})
$btnClear.Add_Click({
    $txtInput.Text = ''
    $lblResult.Text = 'Not looked up yet.'
    $txtOut.Text = ''
    $script:LastSmtp = $null
    $script:LastADServer = $null
        $script:LastType = $null
    Update-ToolButtonState -DetectedType $null
})
$btnClose.Add_Click({ $form.Close() })

Build-ToolButtons
Update-ToolButtonState -DetectedType $null

$script:Theme = Get-ConsoleTheme
if ($script:Theme) {
    Set-ConsoleTheme -Form $form -Theme $script:Theme
    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'AdminConsole' | Out-Null
    # lblWarn is NoTheme-tagged, so it is coloured explicitly here: the theme's
    # own text colour on a themed console, grey on the default one.
    $wc = ConvertTo-ThemeColor $script:Theme.FormText
    if ($wc) { $lblWarn.ForeColor = $wc }
}
Add-FormLogo -Form $form -Corner 'TopRight' -MaxHeight 32 -Margin 6 | Out-Null

Update-ConnectionButtonColors

# Surfaced here, not at migration time -- the move happens before $txtOut
# exists, and a file move the operator never sees is the kind of thing that
# gets blamed for something else three weeks later.
if ($script:ConfigMigrationNotes -and $script:ConfigMigrationNotes.Count -gt 0) {
    $form.Add_Shown({
        $txtOut.Text = ("Config folder created -- existing settings moved in:" +
            [Environment]::NewLine +
            (($script:ConfigMigrationNotes | ForEach-Object { "  $_" }) -join [Environment]::NewLine) +
            [Environment]::NewLine +
            "Update rule from here: replace the .ps1 files, leave Config\ alone.")
    })
}

Set-FormScreenCorner -Form $form -Corner 'TopLeft'

Update-PimLabel
if ($script:PimTimerOn) {
    # One tick a minute doing arithmetic on two timestamps. Not created at all
    # when the countdown is off -- an idle timer firing into a hidden label is
    # work nobody asked for.
    $pimTimer = New-Object System.Windows.Forms.Timer
    $pimTimer.Interval = 60000
    $pimTimer.Add_Tick({ Update-PimLabel })
    $pimTimer.Start()
    $form.Add_FormClosed({ $pimTimer.Stop(); $pimTimer.Dispose() })
}

[void]$form.ShowDialog()
#endregion
