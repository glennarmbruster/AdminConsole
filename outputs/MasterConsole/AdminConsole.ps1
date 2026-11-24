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
# BEFORE ANYTHING ELSE: activate your access (PIM)
# ---------------------------------------------------------------------------
#
# Your admin rights are eligible, not standing. Until you activate them you are
# signed in but have no admin access, and tools will fail in ways that look
# like faults rather than permissions.
#
# Use the ACTIVATE button in the Connect row. It opens Privileged Identity
# Management in your admin browser profile, with two choices:
#
#     Groups      -- the normal route; Exchange and most day-to-day work
#     Entra roles -- use this for SharePoint (see below)
#
# Activation happens in the portal. Doing it from PowerShell needs Graph
# permissions this tenant does not consent to.
#
# ORDER MATTERS, and this is the mistake worth avoiding:
#
#     1. Activate FIRST
#     2. Connect SECOND  (ExO / SPO / OnPrem)
#
# A connection made before activating holds the token it was issued with, and
# that token does not contain the access you activated afterwards. Everything
# will then fail as though the activation never happened. If you have already
# connected, activate and then reconnect.
#
# Some assignments need approval. Those come back as PENDING APPROVAL rather
# than active -- a request is not access, and you have to wait for it.
#
# ---------------------------------------------------------------------------
# SHAREPOINT: use the Entra role, not the group
# ---------------------------------------------------------------------------
#
# SharePoint access can be activated through the group assignment, and it does
# work -- but it is noticeably slower to take effect, because the membership
# has to propagate before SharePoint honours it.
#
# Activating the SINGULAR ENTRA ROLE for SharePoint instead is faster and is
# the recommended route when SPO is what you need. That is the "Entra roles"
# button on the Activate window. The group route is fine for everything else.
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
#   * NOT CONNECTED. A tool that needs Exchange Online, SharePoint Online,
#     on-premises Exchange or Active Directory is greyed until that connection
#     exists in this window. The status row at the top says which you have.
#     Connect what is missing -- and if you have not activated yet, see above,
#     because connecting without access is the usual cause.
#
# Hover a greyed button and the tooltip says which of the two it is.
#
# ---------------------------------------------------------------------------
#
# Every tool has its own ? button in its top-right corner explaining what it
# does and which directories it touches.
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
$lblExo.Size = New-Object System.Drawing.Size(156, 20)
$lblExo.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$lblExo.Text = "ExO: ..."
$form.Controls.Add($lblExo)

$lblSpo = New-Object System.Windows.Forms.Label
$lblSpo.Tag = 'NoTheme'
$lblSpo.Location = New-Object System.Drawing.Point(170, 8)
$lblSpo.Size = New-Object System.Drawing.Size(156, 20)
$lblSpo.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$lblSpo.Text = "SPO: ..."
$form.Controls.Add($lblSpo)

$lblTenant = New-Object System.Windows.Forms.Label
$lblTenant.Tag = 'NoTheme'
$lblTenant.Location = New-Object System.Drawing.Point(330, 8)
$lblTenant.Size = New-Object System.Drawing.Size(330, 20)
$lblTenant.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblTenant.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblTenant)

# On-premises indicators share the second row with the hint text. The
# capability map has always probed EMS and AD -- that is what greys the on-prem
# tools out -- but nothing displayed the answer, so a missing on-prem session
# showed up as a mysteriously dead button rather than as a status you could read.
$lblEms = New-Object System.Windows.Forms.Label
$lblEms.Tag = 'NoTheme'
$lblEms.Location = New-Object System.Drawing.Point(12, 30)
$lblEms.Size = New-Object System.Drawing.Size(156, 20)
$lblEms.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$lblEms.Text = "EMS: ..."
$form.Controls.Add($lblEms)

$lblAd = New-Object System.Windows.Forms.Label
$lblAd.Tag = 'NoTheme'
$lblAd.Location = New-Object System.Drawing.Point(170, 30)
$lblAd.Size = New-Object System.Drawing.Size(156, 20)
$lblAd.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$lblAd.Text = "AD: ..."
$form.Controls.Add($lblAd)

$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Tag = 'NoTheme'
# Short, and stopping at x=640, because everything from there to the right edge
# is reserved for the logo and the '?'.
#
# It was 482px wide reaching x=812, which is why the logo had nowhere to grow
# and the '?' kept being pushed into it. The four status labels above already
# name exactly what is not connected, so this only has to point at the buttons.
$lblWarn.Text = "Use Connect below."
$lblWarn.Location = New-Object System.Drawing.Point(508, 31)
$lblWarn.Size = New-Object System.Drawing.Size(132, 18)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Italic)
$lblWarn.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblWarn)

# PIM countdown. Row 2, DIRECTLY BELOW the tenant label -- both start at x=330.
#
# It was in the top-right corner, where the '?' sat on top of it. The help
# button positions itself at run time and treats a Label as an obstacle only as
# far as its TEXT reaches -- and at the moment the button is placed this label
# is still empty and hidden, so it measured as zero-width and was not avoided.
# Content that appears later cannot be dodged by a layout decision made
# earlier; the fix is to put it somewhere the corner never reaches.
#
# Hidden until something has actually been recorded, so a console that has
# never used the Activate button shows nothing rather than a blank promise.
$lblPim           = New-Object System.Windows.Forms.Label
$lblPim.Tag       = 'NoTheme'
$lblPim.Location  = New-Object System.Drawing.Point(330, 31)
$lblPim.Size      = New-Object System.Drawing.Size(170, 18)
$lblPim.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblPim.Visible   = $false
$form.Controls.Add($lblPim)

function Update-PimLabel {
    try {
        $line = Get-PimStatusLine
        if (-not $line) { $lblPim.Visible = $false; return }
        $lblPim.Text      = $line
        $lblPim.ForeColor = Get-PimStatusColour
        $lblPim.Visible   = $true
        $toolTip.SetToolTip($lblPim, 'Recorded by the Activate button, not read from Entra. Clear it there if it looks wrong.')
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
        $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(80, 220, 130)
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

$script:DefaultDC = 'fran-vm-dc01.franchises.expresspersonnel.com'

$lblInput = New-Object System.Windows.Forms.Label
$lblInput.Text = "Enter mailbox or group email address ($($script:MailDomain) optional):"
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
            [PSCustomObject]@{ Name = 'Connect Scripts'; Rows = 1 }
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
#
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
        $lblVersion.Text      = 'Express Admin Toolkit -- unversioned copy (no Version.json)'
        $lblVersion.ForeColor = [System.Drawing.Color]::DimGray
    } elseif ($vv.Modified.Count -eq 0 -and $vv.Missing.Count -eq 0 -and $vv.Extra.Count -eq 0) {
        $lblVersion.Text      = "Express Admin Toolkit v$($vv.Version)  --  $($vv.Matched) files verified  --  released $($vv.Released)"
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

        $lblVersion.Text      = "Express Admin Toolkit v$($vv.Version)  --  $detail  --  run Tester for the list"
        $lblVersion.ForeColor = [System.Drawing.Color]::FromArgb(230, 170, 60)
    }
} catch {
    $lblVersion.Text      = 'Express Admin Toolkit -- version unavailable'
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
$script:LastEnvironment = $null
$script:AllTools = @()

try {
    $script:AllTools = Get-ToolDefinitions -Path $ToolsConfigPath
} catch {
    $script:AllTools = @()
    $txtOut.Text = $_.Exception.Message
}

# Find-TargetEnvironment comes from Common.ps1.
#
# A local copy lived here and shadowed it. It had gone stale: it still
# pinned $script:DefaultDC where Common uses $script:FranchiseAutoDetectDC
# -- the pinning removed after the write-scope bug (14f). The console was
# quietly running the pre-fix version of a function it dot-sources.
# See ARCHITECTURE 14n.
function Get-ToolParameterSet {
    <#
        Which of the console's context parameters a tool actually declares.

        Passing a parameter a script does not declare is a BINDING ERROR, and
        the failure mode is asymmetric in a way that hides it: a script with no
        param() block at all absorbs anything into $args and works, so most of
        the toolkit tolerated being handed all three. A script that declares
        SOME of them rejects the rest outright -- which is how a newly added
        tool can fail on its very first click while every older one is fine.

        Same approach Invoke-CasQuery uses against cmdlet surfaces: ask what is
        declared, pass only that.
    #>
    param([Parameter(Mandatory)][string]$ScriptPath)

    $ctx = [ordered]@{
        SmtpAddress      = $script:LastSmtp
        Environment      = $script:LastEnvironment
        DomainController = $script:LastADServer
    }
    try {
        $cmd = Get-Command -Name $ScriptPath -CommandType ExternalScript -ErrorAction Stop
        # No param() block means no declared parameters -- and no rejection
        # either, so pass the lot and let $args have them.
        if ($cmd.Parameters.Count -eq 0) { return $ctx }

        $out = @{}
        foreach ($k in $ctx.Keys) {
            if ($cmd.Parameters.ContainsKey($k)) { $out[$k] = $ctx[$k] }
        }
        return $out
    } catch {
        # Unreadable script: fall back to the old behaviour rather than
        # refusing to launch. A tool that might work beats one that cannot.
        return $ctx
    }
}

function Test-ToolDeclaresOwnProcess {
    <#
        Does the script itself insist on running in its own process?

        Tools.json has a NewProcess checkbox, and on 2026-08-06 an on-premises
        tool ran with it unticked. It loaded into this console's process, where
        Exchange Online was already connected. The two share cmdlet names, so
        every write went to the cloud and failed with a message that read like
        a permissions problem.

        The operator did not do anything wrong. A checkbox that must be
        remembered for the tool to be CORRECT -- as opposed to merely
        convenient -- is a trap, and the same trap as the hidden DC selector
        that went on applying its default after it stopped being visible.

        So the script declares the requirement and the console obeys it. The
        checkbox still works, and still means "I would like this in its own
        window", but it can no longer be the only thing standing between an
        on-premises tool and the wrong directory.

        Declared by putting this on any line in the first 40 lines:

            # TOOLKIT-REQUIRES: OwnProcess

        Read from the file rather than from Tools.json on purpose: the file is
        the thing that knows, and it stays true when the tool is copied into
        another tenant's toolkit with a fresh Tools.json.
    #>
    param([Parameter(Mandatory)][string]$ScriptPath)

    if ($null -eq $script:OwnProcessCache) { $script:OwnProcessCache = @{} }
    if ($script:OwnProcessCache.ContainsKey($ScriptPath)) { return $script:OwnProcessCache[$ScriptPath] }

    $result = $false
    try {
        # Header only. This runs on every click, and reading whole files to
        # answer a header question is how the cold-launch scan got slow (14c).
        $head = Get-Content -LiteralPath $ScriptPath -TotalCount 40 -ErrorAction Stop
        $result = [bool]($head -match 'TOOLKIT-REQUIRES:\s*OwnProcess')
    } catch { }

    $script:OwnProcessCache[$ScriptPath] = $result
    return $result
}

function Start-ToolScriptInNewProcess {
    <#
        Launches a tool in its OWN PowerShell process, connecting to
        on-premises Exchange first.

        Why this is safe here when it was not for the connect buttons: a
        connect script has to leave its session in the console's process, so a
        child process was useless. A TOOL only needs the session itself -- it
        connects, does the work, and exits. Nothing has to cross the boundary.

        Two things fall out for free. The child has its own thread and message
        loop, so sign-in and credential dialogs behave normally (section 14d).
        And on-premises cmdlets are loaded only there, so Get-Mailbox in the
        child unambiguously means on-prem while the console's stays cloud.

        The cost is real and paid every launch: a fresh Import-PSSession is
        tens of seconds before the tool's window appears.
    #>
    param([Parameter(Mandatory)]$Tool, [Parameter(Mandatory)][string]$ScriptPath)

    # Launch the same PowerShell that is running us, not whatever
    # 'powershell.exe' resolves to on PATH.
    $psExe = $null
    try { $psExe = (Get-Process -Id $PID).Path } catch { }
    if (-not $psExe) { $psExe = 'powershell.exe' }

    # Every path and value is single-quoted with internal quotes doubled, so
    # the whole -Command payload can be wrapped in double quotes without
    # anything inside fighting it. Start-Process does NOT quote for you --
    # the same trap that once sent Edge to a profile called 'Profile'.
    $esc = { param([string]$s) "'" + ("$s" -replace "'", "''") + "'" }

    $parts = New-Object System.Collections.ArrayList
    $connect = Join-Path $PSScriptRoot 'Connect-ExSE.ps1'
    if (Test-Path -LiteralPath $connect) {
        [void]$parts.Add("& $(& $esc $connect)")
    }

    $call = "& $(& $esc $ScriptPath)"
    foreach ($kv in (Get-ToolParameterSet -ScriptPath $ScriptPath).GetEnumerator()) {
        if ("$($kv.Value)".Trim()) { $call += " -$($kv.Key) $(& $esc $kv.Value)" }
    }
    [void]$parts.Add($call)

    # -NoExit so the window survives the tool closing: the output is the record
    # of what happened, and a window that vanishes takes it with it.
    $argLine = '-NoExit -ExecutionPolicy Bypass -Command "' + ($parts -join '; ') + '"'
    Start-Process -FilePath $psExe -ArgumentList $argLine
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
    # The script's own declaration wins over the checkbox. See
    # Test-ToolDeclaresOwnProcess for why this is not merely a convenience.
    $needsOwnProcess = [bool]$Tool.NewProcess -or (Test-ToolDeclaresOwnProcess -ScriptPath $scriptPath)

    try {
        if ($needsOwnProcess) {
            Start-ToolScriptInNewProcess -Tool $Tool -ScriptPath $scriptPath
            $txtOut.Text = "'$($Tool.Name)' opened in its own PowerShell window." +
                [Environment]::NewLine +
                'It connects to on-premises Exchange first, so the tool takes a while to appear.' +
                [Environment]::NewLine +
                'That window keeps its own connections -- this console is unaffected.'
            return
        }
        # @toolParams, not @(...). Splatting needs a VARIABLE holding the
        # dictionary; @(...) is an array subexpression and would hand the whole
        # thing over as one positional argument.
        $toolParams = Get-ToolParameterSet -ScriptPath $scriptPath
        & $scriptPath @toolParams
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
            $tool = $this.Tag

            # A connect script CANNOT run from here.
            #
            # ShowDialog() holds this thread in a modal message loop, and a
            # modal form disables every other window on its thread. The
            # Exchange Online sign-in prompt is created on this thread, so it
            # draws and then never accepts input -- confirmed 2026-08-03 by
            # running the same script from a bare prompt, where it works.
            # -DisableWAM was not enough: it moves sign-in off the WAM broker
            # onto an EMBEDDED browser, which is still a window on this thread.
            #
            # So the button records what to run and closes the form. The work
            # happens after ShowDialog() returns, where there is no modal loop
            # -- and in this SAME process, which is the whole point: a
            # connection made anywhere else would not be ours.
            if (Get-ConnectButtonService -Tool $tool) {
                $script:PendingConnect = $tool
                $form.Close()
                return
            }

            Start-ToolScript -Tool $tool
            # Tools run in THIS process via '&', so a tool that connected to
            # something has changed what every other tool can do. The moment a
            # tool returns is the reliable re-probe point. Full probe including
            # SharePoint -- a tool finishing is infrequent and user-initiated,
            # so the cost is fine here in a way it would not be on every Look Up.
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
# ORDER MATTERS: first match wins, and the on-premises pattern MUST be tested
# before the cloud one. 'Connect Exchange SE' contains the word 'exchange', so
# with ExO first the on-prem button would be painted with Exchange ONLINE's
# status -- green while the on-prem session was down, and no error anywhere.
$script:ConnectButtonServices = @(
    [PSCustomObject]@{ Service = 'EMS'; Pattern = 'exse|\bse\b|on-?prem|\bems\b' }
    [PSCustomObject]@{ Service = 'ExO'; Pattern = 'exo|exchange' }
    [PSCustomObject]@{ Service = 'SPO'; Pattern = 'spo|sharepoint' }
)

$script:ConnClrOk   = [System.Drawing.Color]::FromArgb(198, 239, 206)
$script:ConnClrBad  = [System.Drawing.Color]::FromArgb(255, 199, 206)
$script:ConnClrNone = $null   # $null = leave the themed colour alone
$script:ConnTextClr = [System.Drawing.Color]::FromArgb(20, 20, 20)

function Get-ConnectButtonService {
    <#
        Which service a CONNECT button is for, or $null if it is an ordinary
        tool.

        The name must contain the word 'connect' before any service pattern is
        considered. Without that gate the service patterns match far too much:
        'on-prem' alone caught 1-Onprem-Export and 2-Onprem-Decomm, which are
        ordinary DL Conversions tools -- and since this function now also
        decides which buttons SKIP capability greying, they would have quietly
        stopped greying out when the on-premises session was missing.

        'Connect' is the honest test anyway. This asks whether a button's job
        is to establish a connection, and that is what the word means.
    #>
    param($Tool)

    # Identified by SCRIPT FILE first, name second.
    #
    # The filename is a convention this toolkit controls -- every connect
    # script is Connect-<service>.ps1 -- whereas the button NAME is free text
    # the operator types into ToolManager. Keying only on the name made the
    # exemption depend on somebody happening to include the word 'connect',
    # which is a rule nobody was told and which fails silently: the button
    # greys out and nothing explains that its label was the problem.
    $leaf = ''
    try { $leaf = Split-Path -Leaf "$($Tool.Script)" } catch { }
    $name = "$($Tool.Name)"

    $isConnect = ($leaf -imatch '^connect[-_]') -or ($name -imatch '\bconnect\b')
    if (-not $isConnect) { return $null }

    # Service still comes from whichever of the two says something. Filename
    # first for the same reason.
    foreach ($svc in $script:ConnectButtonServices) {
        if ($leaf -imatch $svc.Pattern) { return $svc.Service }
    }
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
        $ems = $map['EMS']
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
            $state = switch ($svc) { 'ExO' { $exo } 'EMS' { $ems } 'SPO' { $spo } }
            if (-not $state) { continue }

            $ctrl.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $ctrl.BackColor = if ($state.Available) { $script:ConnClrOk } else { $script:ConnClrBad }
            # Pale status colours need dark text. Without this the button keeps
            # the theme's white ForeColor and reads as disabled.
            $ctrl.ForeColor = $script:ConnTextClr
            $toolTip.SetToolTip($ctrl, "$svc -- $($state.Detail)")
        }

        # Contrast, not decoration. These sit on the theme's dark navy, where
        # #008000 green and #C80000 red are both dark-on-dark: legible if you
        # already know what they say, which is not what a status light is for.
        # Lifted to a mid-bright green and a warm red that hold up on dark
        # without glowing on the light themes.
        $green = [System.Drawing.Color]::FromArgb(80, 220, 130)
        $red   = [System.Drawing.Color]::FromArgb(255, 105, 97)

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
        # EMS and AD come from the same map, so these labels cannot disagree
        # with the grey-out (which is the whole reason the map exists).
        $ems = $map['EMS']
        $ad  = $map['AD']
        if ($ems.Available) { $lblEms.Text = 'EMS: Connected';     $lblEms.ForeColor = $green }
        else                { $lblEms.Text = 'EMS: Not Connected'; $lblEms.ForeColor = $red }
        $toolTip.SetToolTip($lblEms, "$($ems.Detail)")

        if ($ad.Available) { $lblAd.Text = 'AD: Connected';     $lblAd.ForeColor = $green }
        else               { $lblAd.Text = 'AD: Not Connected'; $lblAd.ForeColor = $red }
        $toolTip.SetToolTip($lblAd, "$($ad.Detail)")

        $lblWarn.Visible = -not ($exo.Available -and $spo -and $spo.Available -and
                                 $ems.Available -and $ad.Available)

        # Same probe, same answer: repaint the tool buttons from the map that
        # just produced these labels.
        Update-CapabilityState
    } catch {
        $lblExo.Text = 'ExO: unavailable'
        $lblExo.ForeColor = [System.Drawing.Color]::DimGray
        $lblSpo.Text = 'SPO: unavailable'
        $lblSpo.ForeColor = [System.Drawing.Color]::DimGray
        $lblEms.Text = 'EMS: unavailable'
        $lblEms.ForeColor = [System.Drawing.Color]::DimGray
        $lblAd.Text = 'AD: unavailable'
        $lblAd.ForeColor = [System.Drawing.Color]::DimGray
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

    # A tool whose job is to CREATE a connection must never be gated on that
    # connection. Connect-ExO.ps1 calls Connect-ExchangeOnline, so the
    # capability scan correctly reports it as needing Exchange Online -- and
    # acting on that would grey the button out at exactly the moment it is the
    # only useful control on the form. The greying is right for every other
    # tool and exactly inverted for these.
    #
    # Keyed on Get-ConnectButtonService rather than the section name, because
    # sections are user-editable in Sections.json and renaming one must not
    # quietly re-arm this.
    if (Get-ConnectButtonService -Tool $Tool) { return $null }

    # A NewProcess tool provisions its OWN session in its OWN process. This
    # console's probes describe this console's process, which is not the one
    # the tool will use -- so greying it here would block it for a condition
    # that has nothing to do with whether it can run. Same inverted logic as
    # the connect buttons, arrived at from the other direction.
    if ($Tool.NewProcess) { return $null }
    $tp = Join-Path $PSScriptRoot $Tool.Script
    if ((Test-Path -LiteralPath $tp) -and (Test-ToolDeclaresOwnProcess -ScriptPath $tp)) { return $null }

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
            if ($tool.NewProcess) {
                $toolTip.SetToolTip($ctrl,
                    "Opens in its own PowerShell window and connects to on-premises Exchange itself." + "`r`n" +
                    "This console's connections do not apply to it, so it is never greyed out." + "`r`n" +
                    "Expect a delay while that window imports the Exchange session.")
            } else {
                $a = $script:ToolAvail["$($tool.Name)"]
                $toolTip.SetToolTip($ctrl, $(if ($a) { $a.Reason } else { '' }))
            }
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
            $script:LastEnvironment = $null
            $script:LastType = $null
            $lblResult.Text = 'Enter an email address first. (Tools grey-out has been cleared.)'
            Update-ToolButtonState -DetectedType $null
            return
        }
        $lblResult.Text = 'Looking up (HQ first, then Franchise)...'
        $envResult = Find-TargetEnvironment -Smtp $smtp
        if (-not $envResult.Environment) {
            $script:LastSmtp = $null
            $script:LastADServer = $null
            $script:LastEnvironment = $null
            $script:LastType = $null
            $lblResult.Text = "Not found in HQ or Franchise -- $smtp"
            if ($envResult.Errors.Count -gt 0) {
                $txtOut.Text = ($envResult.Errors -join "`r`n`r`n")
            }
            Update-ToolButtonState -DetectedType $null
            return
        }
        $target = Resolve-Target -Smtp $smtp -ADServer $envResult.ADServer
        $script:LastSmtp = $smtp
        $script:LastADServer = $envResult.ADServer
        $script:LastEnvironment = $envResult.Environment
        $script:LastType = $target.Type
        switch ($target.Type) {
            'Mailbox'      { $lblResult.Text = "Found: Mailbox ($($envResult.Environment)) -- $smtp" }
            'UnifiedGroup' { $lblResult.Text = "Found: Microsoft 365 Group ($($envResult.Environment)) -- $smtp" }
            'OnpremGroup'  { $lblResult.Text = "Found: On-Prem Group ($($envResult.Environment)) -- $smtp" }
            default {
                $lblResult.Text = "AD object found in $($envResult.Environment) but not confirmed as a Mailbox, Microsoft 365 Group, or on-prem AD group -- $smtp"
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
    $script:LastEnvironment = $null
    $script:LastType = $null
    Update-ToolButtonState -DetectedType $null
})
$btnClose.Add_Click({ $form.Close() })

Build-ToolButtons
Update-ToolButtonState -DetectedType $null

$script:Theme = Get-ConsoleTheme
if ($script:Theme) {
    Set-ConsoleTheme -Form $form -Theme $script:Theme
    # lblWarn is NoTheme-tagged, so it is coloured explicitly here: the theme's
    # own text colour on a themed console, grey on the default one.
    $wc = ConvertTo-ThemeColor $script:Theme.FormText
    if ($wc) { $lblWarn.ForeColor = $wc }
}
# Margin 10, not 6. At 6 the logo sat hard against the window edge and looked
# clipped; the corner now has room to breathe and the '?' tucks in beside it.
Add-FormLogo -Form $form -Corner 'TopRight' -MaxHeight 64 -Margin 10 | Out-Null

Update-ConnectionButtonColors

# One write, after every button has been classified, rather than one per tool.
# The scan itself is what used to make a cold launch slow; persisting it means
# only the FIRST launch on this machine pays, and only again for scripts that
# have actually changed since. Failing to save costs a rescan, never an error.
Save-ScriptCapabilityCache

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

$script:PendingConnect = $null
# The console's own '?', in the same corner as every tool's. x = ClientSize-40
# lands at 784..812, which is clear -- lblTenant ends at 660.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Admin Console' | Out-Null

Update-PimLabel

# One tick a minute, doing arithmetic on two timestamps and setting a label.
# Deliberately not shorter, and deliberately no DoEvents anywhere near it --
# see 14c on why extra message-loop activity in this console is treated with
# suspicion.
$pimTimer = New-Object System.Windows.Forms.Timer
$pimTimer.Interval = 60000
$pimTimer.Add_Tick({ Update-PimLabel })
$pimTimer.Start()
$form.Add_FormClosed({ $pimTimer.Stop(); $pimTimer.Dispose() })

[void]$form.ShowDialog()

# Past this line the modal loop is gone, so a sign-in window can take input.
if ($script:PendingConnect) {
    $pending = $script:PendingConnect
    $script:PendingConnect = $null

    $path = Join-Path $PSScriptRoot $pending.Script
    if (Test-Path -LiteralPath $path) {
        Write-Host ''
        Write-Host "Closing the console to sign in -- it will reopen when you are done." -ForegroundColor Cyan
        try {
            & $path
        } catch {
            Write-Host (Format-ErrorDetail -ErrorRecord $_ -Context "Running $($pending.Name)") -ForegroundColor Red
        }
    } else {
        Write-Host "'$($pending.Name)' points to '$($pending.Script)', which does not exist." -ForegroundColor Red
    }

    # Relaunch in the SAME process, so the session just established is still
    # ours. A new PowerShell would authenticate perfectly and leave this
    # console exactly as disconnected as before.
    #
    # Each connect nests one invocation deeper. Two or three per session is
    # normal and harmless; the guard exists so that a stuck loop ends with a
    # readable message instead of a stack overflow.
    $global:AdminConsoleDepth = [int]$global:AdminConsoleDepth + 1
    if ($global:AdminConsoleDepth -gt 8) {
        Write-Host ''
        Write-Host 'Console reopened 8 times in this session; not reopening again.' -ForegroundColor Yellow
        Write-Host "Run .\$(Split-Path -Leaf $PSCommandPath) yourself when ready." -ForegroundColor Yellow
        $global:AdminConsoleDepth = 0
    } else {
        & $PSCommandPath
        $global:AdminConsoleDepth = [int]$global:AdminConsoleDepth - 1
    }
}
#endregion
