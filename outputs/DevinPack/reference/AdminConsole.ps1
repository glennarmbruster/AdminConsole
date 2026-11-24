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
# WHAT THIS TENANT HAS
# ---------------------------------------------------------------------------
#
# The status row shows only what this tenant actually has. A tenant with no
# on-premises Exchange has no EMS indicator at all -- not a grey one. Config\
# Tenant.json decides that, and TenantManager edits it.
#
# (This used to say ToolManager > Settings, which cannot edit Tenant.json and
# never could. That dialog writes Settings.json -- engineer and machine
# preferences. Tenant.json is tenant FACTS and had no editor at all until
# TenantManager was written.)
#
# So the row describes THIS deployment. If something you expect is missing from
# it, the capability is switched off in Tenant.json rather than broken.
#
# ---------------------------------------------------------------------------
# CONNECT FIRST
# ---------------------------------------------------------------------------
#
# Tools stay greyed until this window holds the connection they need. The
# connection belongs to THIS window, not to your account.
#
# If your admin rights are activated on demand rather than standing, activate
# BEFORE connecting. A session made first holds the token it was issued with,
# and that token does not contain access granted afterwards -- everything then
# fails as though the activation never happened. Activate, then reconnect.
#
# The PIM row and the Activate button appear only where Tenant.json says this
# tenant uses on-demand rights.
#
# ---------------------------------------------------------------------------
# CHECK THE TENANT BEFORE YOU WRITE
# ---------------------------------------------------------------------------
#
# One toolkit serves several tenants and looks the same in all of them. The
# tenant name is in the top row, and a colour can be bound per tenant in
# ThemeManager -- colour tells you where you are faster than reading does.
#
# A (?) beside the tenant name means the connected domain does not match what
# Tenant.json expects. That is a question, not a fault -- a tenant with several
# accepted domains can disagree legitimately -- but answer it before a write.
#
# ---------------------------------------------------------------------------
# WHY A BUTTON IS GREYED
# ---------------------------------------------------------------------------
#
# Three reasons, and they want three different responses:
#
#   * WRONG OBJECT TYPE. The tool cannot act on what you just looked up. Look
#     up something else and it comes back.
#
#   * NOT CONNECTED. It needs a service you have not connected in this window.
#     Connect it -- and if rights are activated on demand, activate first.
#
#   * NOT AVAILABLE IN THIS TENANT. The tool needs something this tenant does
#     not have. Connecting will not help. That tool can be removed from
#     Tools.json for this deployment, and the tooltip says so.
#
# Hover a greyed button and the tooltip says which of the three it is.
#
# ---------------------------------------------------------------------------
#
# The line at the bottom shows the toolkit version and whether the files still
# match it. Amber means this copy differs from the one it was stamped from --
# run Tester for the list.
#
# Every tool has its own ? button explaining what it does and what it touches.
# TOOLKIT-HELP-END

param(
    [string]$ToolsConfigPath = ''
)

#region Relaunch trampoline ----------------------------------------------------
<#
    THE CONSOLE RE-OPENS ITSELF IN A LOOP, NOT BY RECURSION.

    Why it re-opens at all: a Connect script has to run in THIS process, or the
    session it establishes belongs to a process that is about to exit and this
    console stays as disconnected as it was. So the window closes, the connect
    runs on the same thread, and the console comes back.

    HOW IT USED TO COME BACK, AND WHY THAT WAS WRONG. The bottom of this file
    called '& $PSCommandPath' -- the script invoking itself. That works, but
    every connect left the process one stack frame deeper, and the frames only
    unwound when the console was finally closed. A depth counter capped it at 8
    and printed "not reopening again", which is a guard doing its job and an
    operator being told to go away.

    Nobody hit it while PowerShell was restarted daily; the counter reset with
    the process. It surfaced the day the console first stayed open overnight --
    ExchangeOnlineManagement v3 keeps a REST-backed session alive, so the
    process now outlives the working day and the connects accumulate.

    The fix is a trampoline. The FIRST invocation becomes a loop and does no
    work of its own; it calls the script again and keeps calling it for as long
    as the body asks to be re-opened. Every real invocation therefore runs at
    exactly one level deep, forever, and the depth guard is gone because there
    is no depth to guard.

    $global: rather than $script:, because the two invocations are separate
    script scopes in one process and a $script: variable would not be visible
    across them. Cleared in a finally so an unhandled error inside the console
    cannot leave the flag set and strand the next launch inside a loop.
#>
if (-not $global:ToolkitConsoleTrampoline) {
    $global:ToolkitConsoleTrampoline = $true
    try {
        do {
            # Reset BEFORE the call, not after. The body sets this when it
            # wants to come back; reading a stale value would re-open a console
            # nobody asked for.
            $global:ToolkitConsoleRelaunch = $false
            & $PSCommandPath @PSBoundParameters
        } while ($global:ToolkitConsoleRelaunch)
    } finally {
        $global:ToolkitConsoleTrampoline = $null
        $global:ToolkitConsoleRelaunch   = $null
    }
    return
}
#endregion

#region Startup breadcrumbs ----------------------------------------------------
<#
    A trail on disk, written before anything else can fail.

    WHY THIS EXISTS. The console occasionally appears and vanishes, and the
    second launch is always fine (first seen 2026-08-13). Nothing was recorded,
    so there was nothing to work from.

    The first attempt at this wrote the reason into the console's own output
    box, which was useless: if the window dies, the message dies with it. That
    is the whole failure mode. A diagnostic has to outlive the thing it is
    diagnosing, so this writes to a FILE.

    IT LOGS PROGRESS, NOT JUST ERRORS. An exception that never surfaces still
    leaves the last breadcrumb behind, so the file answers WHERE it stopped even
    when it cannot say why. "Reached: dot-sourcing Common" followed by nothing
    is a complete answer.

    DELIBERATELY SELF-CONTAINED. No Common.ps1, no Get-LogPath, no toolkit
    helpers -- those are the very things that might not have loaded. Only
    $PSScriptRoot and built-ins, and every write is wrapped so that a failure to
    log can never itself take the console down.

    The file is truncated at 200 lines on each launch, so it stays readable and
    cannot grow without bound.
#>
$script:StartupLogPath = Join-Path $PSScriptRoot 'Logs\ConsoleStartup.log'

function Write-StartupBreadcrumb {
    param([string]$Text)
    try {
        $dir = Split-Path -Parent $script:StartupLogPath
        if (-not (Test-Path -LiteralPath $dir)) {
            New-Item -Path $dir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        $line = '{0}  pid {1}  {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $PID, $Text
        Add-Content -LiteralPath $script:StartupLogPath -Value $line -Encoding UTF8 -ErrorAction Stop
    } catch { }
}

function Write-StartupFailure {
    param([string]$Stage, $ErrorRecord)
    try {
        Write-StartupBreadcrumb -Text "FAILED at: $Stage"
        if ($ErrorRecord) {
            Write-StartupBreadcrumb -Text "    $($ErrorRecord.Exception.GetType().FullName): $($ErrorRecord.Exception.Message)"
            # The stack is what turns "it broke" into a line number, and it is
            # the first thing lost when an exception is caught and summarised.
            $st = "$($ErrorRecord.ScriptStackTrace)"
            foreach ($l in ($st -split "`r?`n")) { if ("$l".Trim()) { Write-StartupBreadcrumb -Text "    $l" } }
        }
    } catch { }
}

# Keep the file short. Reading a startup log means reading the LAST run, and a
# thousand old lines is how the relevant twelve get missed.
try {
    if (Test-Path -LiteralPath $script:StartupLogPath) {
        $existing = @(Get-Content -LiteralPath $script:StartupLogPath -ErrorAction Stop)
        if ($existing.Count -gt 200) {
            Set-Content -LiteralPath $script:StartupLogPath -Value $existing[-200..-1] -Encoding UTF8 -ErrorAction Stop
        }
    }
} catch { }

Write-StartupBreadcrumb -Text '--- launch ---'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Write-StartupBreadcrumb -Text 'WinForms assemblies loaded'

try {
    . (Join-Path $PSScriptRoot 'Common.ps1')
    Write-StartupBreadcrumb -Text 'Common.ps1 dot-sourced'
} catch {
    Write-StartupFailure -Stage 'dot-sourcing Common.ps1' -ErrorRecord $_
    throw
}

# THE MARKER THAT SAYS "the console started this".
#
# Set once, here, before any tool can be launched -- and before the form is
# built, so there is no path where a button exists and the marker does not.
#
# A GUID rather than a fixed word, so it cannot be guessed from reading one
# tool. That is not a security property -- anyone can read Common.ps1 and see
# that any value passes -- it just means the value in a process listing is not
# a reusable password.
#
# Environment variables are inherited by child processes, which is why this
# covers both launch paths: the call operator (same process) and Start-Process
# (new process, inherits). Nothing needs passing through the command line.
#
# It dies with this process. A prompt opened afterwards does not have it.
$env:ToolkitConsoleSession = [guid]::NewGuid().ToString()
Write-StartupBreadcrumb -Text 'console session marker set'
#endregion

# Config now lives in Config\. Resolved after Common.ps1 loads, because the
# resolver is defined there -- a param default cannot call it.
if (-not $ToolsConfigPath) { $ToolsConfigPath = Get-ConfigFilePath -Name 'Tools.json' }

# One-time tidy-up for installs that predate Config\. Runs before anything reads
# a config file, and reports what it did rather than moving files silently.
$script:ConfigMigrationNotes = @(Invoke-ConfigMigration)

# No tenant configured -- say so, loudly, before drawing anything.
#
# Test-TenantCapability returns $false when Tenant.json is absent, and the
# capability gate HIDES tools that are off. So without this the console opens
# looking finished, with most of its buttons simply not there, and nothing
# anywhere saying why. That is indistinguishable from a broken install, and it
# is the first thing a new tenant would ever see.
#
# The toolkit deliberately ships WITHOUT a Tenant.json -- only
# Config\Tenant.TEMPLATE.json -- because shipping one populated with somebody
# else's tenant is worse: it looks configured and is wrong.
if (-not (Test-Path -LiteralPath (Get-ConfigFilePath -Name 'Tenant.json'))) {
    $tm = Join-Path $PSScriptRoot 'TenantManager.ps1'
    $msg = "No Config\Tenant.json, so this deployment does not know which tenant it is." +
           [Environment]::NewLine + [Environment]::NewLine +
           "Until it exists, every capability reads as OFF and the tools that need one are HIDDEN -- the console will look half empty." +
           [Environment]::NewLine + [Environment]::NewLine
    if (Test-Path -LiteralPath $tm) {
        $msg += "Open Tenant Manager now to set it up?"
        $ans = [System.Windows.Forms.MessageBox]::Show($msg, 'Tenant not configured',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Its own process: this console has not drawn its form yet, and a
            # modal dialog on this thread would block the setup window it just
            # offered to open. See 14d.
            Start-Process -FilePath (Get-Command powershell.exe).Source `
                -ArgumentList "-ExecutionPolicy Bypass -File `"$tm`""
            return
        }
    } else {
        $msg += "Run TenantManager.ps1 to set it up."
        [void][System.Windows.Forms.MessageBox]::Show($msg, 'Tenant not configured',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
}

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

# AUTOSIZE, not a fixed 170px.
#
# The line holds Groups AND Roles once both timers are running, and 170px fits
# only the first. Roles was not hidden -- it was off the right-hand edge of a
# label too small for its own text, which looks identical to "not set".
#
# Update-StatusRowLayout positions the labels after it by adding $lbl.Width, so
# a fixed width that disagrees with the content also mis-places everything to
# its right. AutoSize keeps Width honest and the row flows correctly.
$lblPim.AutoSize  = $true
$lblPim.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblPim.Visible   = $false
$form.Controls.Add($lblPim)

function Update-PimLabel {
    <#
        TAKES THE CONTROLS AS PARAMETERS. It used to read $lblPim and $toolTip
        from the enclosing scope, and that was a real fault rather than a
        stylistic one.

        This function is called from a 60-second Timer tick. A Timer tick
        scriptblock is not a closure, so PowerShell resolves its variables
        against whatever scope is CURRENT when the tick fires -- and the
        console launches tools with '& $scriptPath', which gives the tool a
        CHILD scope of this one. While a tool's ShowDialog() pumps messages on
        this same thread, the timer keeps ticking, and any name the tool has
        assigned to shadows the console's.

        Dashboard.ps1 opens with

            $lblPim = $null
            if ($script:HasPim) { $lblPim = Add-DashLine ... }

        so in a tenant WITHOUT PIM the timer found the dashboard's $null and
        threw "The property 'Visible' cannot be found on this object" -- from
        inside this function's own catch, which then threw as well, so nothing
        contained it and the operator got a JIT-debugging dialog captioned
        'Admin Dashboard'.

        IN A TENANT WITH PIM IT DID NOT THROW, WHICH IS WORSE. The dashboard's
        $lblPim was a real Label, so the console's countdown was written into
        the DASHBOARD's label instead of its own -- silently, for as long as
        both existed.

        Passing the controls in removes the name lookup entirely. The timer
        carries them on its own .Tag; see the wiring at the bottom of the file.
    #>
    param(
        [System.Windows.Forms.Label]$Lbl,
        $Tip
    )
    if (-not $Lbl) { return }
    try {
        $line = Get-PimStatusLine
        if (-not $line) { $Lbl.Visible = $false; return }
        $Lbl.Text      = $line
        $Lbl.ForeColor = Get-PimStatusColour
        $Lbl.Visible   = $true
        if ($Tip) {
            $Tip.SetToolTip($Lbl, 'Recorded by the Activate button, not read from Entra. Clear it there if it looks wrong.')
        }
    } catch { $Lbl.Visible = $false }
}

# Lay the status row out from whatever is actually showing.
#
# A capability this tenant does not have is HIDDEN, not greyed. Grey still
# occupies space and still invites a glance -- and "EMS: n/a" in a cloud-only
# tenant is a label that will never say anything else. The row should describe
# this tenant, not the union of every tenant.
#
# Hiding leaves a hole, so the survivors are repositioned rather than left with
# a gap where a control used to be. Fixed coordinates were the reason the old
# row could not lose a member.
#
# Tool BUTTONS keep their grey. That is different: a greyed button says "this
# tool is here and cannot run", which is exactly the prompt to take it out of
# Tools.json. A status label has no such job.
function Update-StatusRowLayout {
    param([System.Windows.Forms.Label[]]$Row, [int]$Y, [int]$Gap = 8)
    $x = 12
    foreach ($lbl in $Row) {
        if (-not $lbl -or -not $lbl.Visible) { continue }
        $lbl.Location = New-Object System.Drawing.Point($x, $Y)
        $x += $lbl.Width + $Gap
    }
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

$script:DefaultDC = (Get-TenantDirectoryServer)

$lblInput = New-Object System.Windows.Forms.Label
$lblInput.Text = "Enter mailbox or group email address ($($(Get-MailDomain)) optional):"
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

# Find box, on the SAME ROW as Look Up / Clear / Close.
#
# Placed here rather than on a row of its own because the console is already
# tall enough to be capped against the working area on a 1080p screen, and a
# search box is not worth 30px of everything else. The buttons end at x=318 and
# the panel is nearly 800 wide, so the space to the right was doing nothing.
#
# Width is set later, with the other PanelW-derived controls -- PanelW is not
# known until Sections.json has been read, and referencing it here would give a
# $null that WinForms turns into a width of zero.
$lblFind = New-Object System.Windows.Forms.Label
$lblFind.Text = 'Find tool:'
$lblFind.Location = New-Object System.Drawing.Point(340, 103)
$lblFind.Size = New-Object System.Drawing.Size(62, 18)
$form.Controls.Add($lblFind)

$txtFind = New-Object System.Windows.Forms.TextBox
$txtFind.Location = New-Object System.Drawing.Point(404, 99)
$txtFind.Size = New-Object System.Drawing.Size(200, 24)
$form.Controls.Add($txtFind)

$btnFindClear = New-Object System.Windows.Forms.Button
$btnFindClear.Text = 'Show all'
$btnFindClear.Location = New-Object System.Drawing.Point(612, 98)
$btnFindClear.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnFindClear)

$lblFindCount = New-Object System.Windows.Forms.Label
$lblFindCount.Location = New-Object System.Drawing.Point(710, 103)
$lblFindCount.Size = New-Object System.Drawing.Size(120, 18)
$form.Controls.Add($lblFindCount)

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
$script:BtnH = $script:ConsoleBtnHeight; $script:BtnPad = $script:ConsoleBtnPad

# The tool list is loaded HERE, several hundred lines above where the buttons
# are built, because button width is measured from the captions and the
# measurement has to happen before any Size() is computed.
#
# This is as high as it can go. The failure path reports into $txtOut, so
# $txtOut has to exist first -- moving this any earlier would turn "Tools.json
# is missing" into a null reference, which is a worse message for the same
# problem.
$script:AllTools = @()
try {
    $script:AllTools = Get-ToolDefinitions -Path $ToolsConfigPath
} catch {
    $script:AllTools = @()
    $txtOut.Text = $_.Exception.Message
}

# PanelW is DERIVED, not chosen, and it allows for a scrollbar that may not be
# there.
#
# It was a flat 800. A button occupies BtnW + Margin on both sides = 158, so a
# row of five needs 790 and the panel's client area was 798 -- five fitted, with
# eight pixels to spare. Add a vertical scrollbar and the client area drops by
# SystemInformation.VerticalScrollBarWidth (17 at 100% DPI) to 781, nine pixels
# short, so the fifth button wrapped to the next row.
#
# That is why ONE section looked different: Admin Tools holds 26 buttons in a
# 4-row panel and is the only section over its declared capacity, so it is the
# only one with a scrollbar. The setting said five across and four of the five
# sections agreed, which made it look like Admin Tools was special-cased
# somewhere. It was not -- it was the only one wide enough to notice.
#
# Reserving the scrollbar width unconditionally costs 17px of window and makes
# the row count independent of how full a section happens to be. A layout that
# changes shape when content grows is a layout that will surprise someone
# mid-demo.
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

$script:SectionCfg    = Get-SectionDefinitions
$script:SectionPanels = @{}

# Labels and declared row counts are kept alongside the panels so the search
# filter can collapse a whole section when nothing in it matches. Without the
# label reference a filtered console shows a heading over an empty box, which
# reads as "these tools disappeared" rather than "nothing here matched".
$script:SectionLabels = @{}
$script:SectionRows   = @{}
$script:SectionStartY = 308
$script:SectionNames  = @(@($script:SectionCfg.Sections) | ForEach-Object { "$($_.Name)" })

$script:PerRow  = [int]$script:SectionCfg.ButtonsPerRow
if ($script:PerRow -lt 1) { $script:PerRow = 5 }

# Button width is MEASURED, not chosen.
#
# It was a flat 150 -- a number that fitted the captions of the day, in the
# font of the machine this was written on, at 100% scaling. None of those three
# things is a constant, so 150 was wrong in both directions at once: wider than
# short captions need, and too narrow for any long one added later.
#
# Get-ConsoleButtonWidth (Common.ps1) returns the narrowest width at which
# every caption still fits INSIDE THE BOX, wrapping allowed. Wrapping has to be
# allowed: 'Mailbox / Onedrive Perms' has always been two lines, and forcing it
# onto one would set the width for all forty-odd buttons and make the window
# considerably wider.
#
# Every tool in Tools.json is measured, including ones a capability check will
# later hide. Which tools are hidden is not known yet -- the probe has not run
# -- and measuring a button that turns out to be invisible costs a few pixels,
# where skipping it would clip a caption the first time a capability was
# switched on.
#
# ButtonWidth in Sections.json overrides the measurement, and is deliberately
# NOT in the Sections dialog. It is there for the case where the fit is correct
# but not to taste; a layout setting that is easy to reach is one that gets
# changed by accident and then reported as a bug.
$cfgW = 0
if ($script:SectionCfg.PSObject.Properties['ButtonWidth']) {
    [int]::TryParse("$($script:SectionCfg.ButtonWidth)", [ref]$cfgW) | Out-Null
}
if ($cfgW -ge 60) {
    $script:BtnW = $cfgW
} else {
    $script:BtnW = Get-ConsoleButtonWidth `
        -Captions @(@($script:AllTools) | ForEach-Object { "$($_.Name)" }) `
        -Font $form.Font `
        -ButtonHeight $script:BtnH
}

$script:BtnCell = $script:BtnW + ($script:BtnPad * 2)

# Start NARROW. Update-SectionPanelWidth widens later if a scrollbar actually
# turns up.
#
# Scrollbar width used to be reserved unconditionally, which cost ~27px down the
# right-hand edge of every panel permanently -- a FlowLayoutPanel is
# left-aligned, so width nothing uses collects on the right and reads as the
# layout being lopsided.
#
# Unconditional was itself a fix, and the bug behind it is still real: a panel
# that grows a scrollbar loses that width from its client area and wraps its
# last button, so the fullest section looked special-cased when it was only the
# fullest. Predicting which sections will overflow gets that right most of the
# time and wrong exactly where it matters -- capability-hidden tools are not
# known at this point, so a count taken here can be higher than what is drawn.
#
# So this no longer predicts. It lays out narrow, asks the panels whether they
# grew scrollbars, and widens every panel together if any of them did. See
# Update-SectionPanelWidth.
$script:PanelW = Get-ConsolePanelWidth -ButtonWidth   $script:BtnW `
                                       -ButtonPad     $script:BtnPad `
                                       -ButtonsPerRow $script:PerRow

# Widened HERE rather than where they were created, because PanelW is not known
# until Sections.json has been read and Get-SectionDefinitions is defined below
# the controls. Referencing it at creation time gave a $null that WinForms
# silently converts to a width of zero -- the input and output boxes would have
# vanished rather than erred.
$txtInput.Width  = $script:PanelW
$lblResult.Width = $script:PanelW
$txtOut.Width    = $script:PanelW

# The find row is laid out from the RIGHT edge of the panel so it stays put as
# ButtonsPerRow changes the console width -- fixed left offsets would strand it
# mid-form at 7 columns and run it over the Close button at 4.
#
# TWO LAYOUTS, because at four or five columns there is genuinely not enough
# room on the button row. The first version of this "handled" that by falling
# back to a fixed 150px box, which at 4 columns put the textbox at 404..554
# inside a panel ending at 540 -- off the edge, on top of the count label. A
# fallback that produces an overlapping layout is not a fallback, it is the
# same bug with an if in front of it.
#
# So: share the button row when it fits, take a row of its own when it does
# not, and move everything below down by exactly that much.
$findRight = 12 + $script:PanelW
$findLeft  = 336
$boxW      = $findRight - 214 - 8 - $findLeft - 64

if ($boxW -ge 90) {
    $lblFind.Location      = New-Object System.Drawing.Point($findLeft, 103)
    $txtFind.Location      = New-Object System.Drawing.Point(($findLeft + 64), 99)
    $txtFind.Size          = New-Object System.Drawing.Size($boxW, 24)
    $btnFindClear.Location = New-Object System.Drawing.Point(($findRight - 214), 98)
    $lblFindCount.Location = New-Object System.Drawing.Point(($findRight - 118), 103)
    $lblFindCount.Size     = New-Object System.Drawing.Size(118, 18)
} else {
    $rowY = 132
    $lblFind.Location      = New-Object System.Drawing.Point(12, ($rowY + 5))
    $txtFind.Location      = New-Object System.Drawing.Point(76, ($rowY + 1))
    $txtFind.Size          = New-Object System.Drawing.Size([Math]::Max(120, ($script:PanelW - 300)), 24)
    $btnFindClear.Location = New-Object System.Drawing.Point(($findRight - 214), $rowY)
    $lblFindCount.Location = New-Object System.Drawing.Point(($findRight - 118), ($rowY + 5))
    $lblFindCount.Size     = New-Object System.Drawing.Size(118, 18)

    # Everything below the find row moves down together. SectionStartY is read
    # by Update-ConsoleSearch when it reflows, so shifting it here is enough --
    # the filter does not need to know which layout is in use.
    $shift = 34
    $lblResult.Location    = New-Object System.Drawing.Point(12, (136 + $shift))
    $txtOut.Location       = New-Object System.Drawing.Point(12, (160 + $shift))
    $script:SectionStartY += $shift
}

$y = $script:SectionStartY
foreach ($sec in @($script:SectionCfg.Sections)) {
    $rows = [int]$sec.Rows
    if ($rows -lt 1) { $rows = 1 }
    $script:SectionRows["$($sec.Name)"] = $rows
    $panelH = ($rows * ($script:BtnH + ($script:BtnPad * 2))) + 12

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "$($sec.Name)"
    $lbl.Location = New-Object System.Drawing.Point(12, $y)
    $lbl.Size = New-Object System.Drawing.Size($script:PanelW, 20)
    $form.Controls.Add($lbl)

    $flow = New-Object System.Windows.Forms.FlowLayoutPanel
    $flow.Location = New-Object System.Drawing.Point(12, ($y + 22))
    $flow.Size = New-Object System.Drawing.Size($script:PanelW, $panelH)

    # Half the slack goes on the left so the spare width reads as a margin on
    # both edges rather than a gap on one. A FlowLayoutPanel lays out from its
    # left padding edge, so this shifts the grid right by half and leaves the
    # other half past the last column.
    $flow.Padding = New-Object System.Windows.Forms.Padding(
                        [int]($script:ConsolePanelSlack / 2), 0, 0, 0)

    $flow.AutoScroll = $true
    $flow.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $form.Controls.Add($flow)

    $script:SectionPanels["$($sec.Name)"] = $flow
    $script:SectionLabels["$($sec.Name)"] = $lbl
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
$lblVersion.Size         = New-Object System.Drawing.Size($script:PanelW, 18)
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

# Width follows PanelW plus the 12px margin on each side, rather than the 824
# it was hardcoded to. PanelW is now derived from ButtonsPerRow, so changing
# that setting to 4 or 6 resizes the window instead of clipping the panels.
$form.ClientSize = New-Object System.Drawing.Size(($script:PanelW + 24), ($y + 8))

# Never taller than the screen it opens on.
#
# The height is the sum of whatever Sections.json declares, so it has always
# been unbounded -- add a section or a row and the window grows, with nothing
# stopping it growing past the bottom of the display. Raising the button height
# to 38 for two-line captions added six pixels per row, eleven rows, sixty-six
# pixels, which is enough to put the version footer and the last section under
# the taskbar on a 1080p screen.
#
# AutoScroll rather than a smaller layout: the buttons stay legible and the
# window scrolls. WorkingArea already excludes the taskbar.
$form.AutoScroll = $true
$work = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea
if ($form.Height -gt $work.Height) {
    $form.Height = $work.Height
    $form.Top    = $work.Top
}
if ($form.Width -gt $work.Width) { $form.Width = $work.Width }
#endregion

#region State + lookup ----------------------------------------------------------
$script:LastSmtp = $null
$script:LastType = $null
$script:LastADServer = $null
$script:LastEnvironment = $null
# $script:AllTools is loaded ABOVE, in the section geometry region, and is
# deliberately NOT re-read here.
#
# It had to move: button width is measured from the tool captions now, so the
# list has to exist before the first Size() is computed. Leaving a second load
# here would re-read the file on every launch and, on the failure path, blank a
# list the layout had already been built from.

# Find-TargetEnvironment comes from Common.ps1.
#
# A local copy lived here and shadowed it. It had gone stale: it still
# pinned $script:DefaultDC where Common uses (Get-TenantDirectoryServer)
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

    # EMPTY IS NOT A VALUE. Context the console does not have is left OUT of the
    # splat rather than passed as ''.
    #
    # Passing it defeats the script's own default: an explicitly supplied empty
    # string binds, and binding happens BEFORE the default can apply. A
    # parameter with [ValidateSet] then fails outright -- 2-Export-Usage died
    # with 'the argument "" does not belong to the set' -- naming the
    # directories it did allow -- when the
    # operator simply had not looked anything up yet.
    #
    # "I do not know which directory" and "the directory is the empty string"
    # are different statements, and only one of them is true here.
    $ctx = [ordered]@{}
    if ("$($script:LastSmtp)".Trim())        { $ctx['SmtpAddress']      = $script:LastSmtp }
    if ("$($script:LastEnvironment)".Trim()) { $ctx['Environment']      = $script:LastEnvironment }
    if ("$($script:LastADServer)".Trim())    { $ctx['DomainController'] = $script:LastADServer }
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

    # Shrink the console window before anything prints.
    #
    # It only ever shows connect output and whatever the tool writes, and at the
    # default size it covers most of the screen -- including the tool's own
    # window, which is the thing you actually wanted.
    #
    # Added FIRST so it runs before the connect banner prints, not after.
    [void]$parts.Add((Get-ConsoleResizeCommand))

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

# Resolve-PairedToolScript and the whole ScriptOnPrem "AD twin" feature were
# REMOVED on 2026-08-13. See ARCHITECTURE section 30.
#
# The short version: a tool could name an on-premises twin and become one
# button that asked which directory to act on. It worked, but ToolManager's
# Move Up / Move Down swapped a hand-maintained list of columns that did not
# include ScriptOnPrem -- so reordering left the twin behind on the old row
# position and it silently reattached to whatever tool landed there. Message
# Trace, DL Send To/As and Connect-ExSE all acquired twins they never had, and
# Save wrote the corruption back.
#
# The pairing was also invisible: nothing on the console said a button had two
# scripts behind it, so the wrong one running looked like the tool
# misbehaving rather than like config.
#
# Grouping (the Group column) does the same job out loud -- several tools
# behind one button, with a chooser that names them -- so pairs become groups.

function Start-ToolScript {
    param($Tool)

    $scriptPath = Join-Path $PSScriptRoot $Tool.Script
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "'$($Tool.Name)' points to '$(Split-Path -Leaf $scriptPath)', which does not exist under this console's folder.`n`nThis is expected if this is a reduced-tier distribution that doesn't include every tool -- otherwise, check Tools.json (ToolManager.ps1) for a stale/typo'd path.",
            'Tool Script Not Found', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    # The script's own declaration wins over the checkbox. See
    # Test-ToolDeclaresOwnProcess for why this is not merely a convenience.
    # THE SCRIPT DECIDES. There is no longer a Tools.json checkbox in the or.
    #
    # NewProcess started as the only control and was demoted to a hint when an
    # unticked on-premises tool wrote to the cloud (see Test-ToolDeclaresOwnProcess).
    # From then on every tool that needed isolation declared it, and the
    # checkbox never decided anything again -- all 55 rows sat unticked for
    # months with no symptom, because the declarations were doing the work.
    #
    # Leaving it meant one requirement expressed in two places, with the
    # weaker copy editable and visible. That is a question waiting to be asked
    # by whoever opens Tool Manager next, and the honest answer was "ignore
    # that column".
    $needsOwnProcess = Test-ToolDeclaresOwnProcess -ScriptPath $scriptPath

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

function Get-ToolGroupName {
    param($Tool)
    if (-not $Tool) { return '' }
    if (-not $Tool.PSObject.Properties['Group']) { return '' }
    return "$($Tool.Group)".Trim()
}

function Show-ToolGroupChooser {
    <#
        The tools behind a group button.

        A group exists to buy back console real estate: five room tools are
        five buttons' worth of grid for something touched twice a month. One
        button, one extra click, and the front page stays readable.

        The chooser repeats each tool's availability rather than just listing
        names. A tool that cannot run is disabled HERE too, with the same
        reason -- otherwise the group button becomes a place where greying goes
        to die, and the operator learns the reason one dialog too late.
    #>
    param([Parameter(Mandatory)][string]$GroupName)

    $tools = @($script:ToolGroups[$GroupName])
    if ($tools.Count -eq 0) { return }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $GroupName
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false
    $dlg.StartPosition = 'CenterParent'
    $dlg.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $pad = 12
    $bw  = 260
    $bh  = 34
    $y   = $pad

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "$($tools.Count) tools"
    $lbl.Location = New-Object System.Drawing.Point($pad, $y)
    $lbl.Size = New-Object System.Drawing.Size($bw, 18)
    $dlg.Controls.Add($lbl)
    $y += 24

    $script:GroupPick = $null
    foreach ($t in $tools) {
        $b = New-Object System.Windows.Forms.Button
        $b.Text = "$($t.Name)"
        $b.Location = New-Object System.Drawing.Point($pad, $y)
        $b.Size = New-Object System.Drawing.Size($bw, $bh)
        $b.Tag = $t

        $block = Get-ToolCapabilityBlock -Tool $t
        if ($block) {
            $b.Enabled = $false
            $tip = New-Object System.Windows.Forms.ToolTip
            $tip.SetToolTip($b, $block.Reason)
        }
        $b.Add_Click({
            $script:GroupPick = $this.Tag
            $this.FindForm().DialogResult = [System.Windows.Forms.DialogResult]::OK
            $this.FindForm().Close()
        })
        $dlg.Controls.Add($b)
        $y += $bh + 6
    }

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'
    $cancel.Location = New-Object System.Drawing.Point($pad, ($y + 6))
    $cancel.Size = New-Object System.Drawing.Size($bw, 28)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($cancel); $dlg.CancelButton = $cancel

    $dlg.ClientSize = New-Object System.Drawing.Size(($bw + ($pad * 2)), ($y + 46))
    Set-ConsoleTheme -Form $dlg
    [void]$dlg.ShowDialog()

    if ($script:GroupPick) {
        $picked = $script:GroupPick
        $script:GroupPick = $null
        Start-ToolScript -Tool $picked
        Update-ConnectionButtonColors
    }
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

    # Group buttons are inserted AT THE POSITION OF THE FIRST MEMBER, and the
    # members are still created -- just hidden. Two reasons for building them
    # rather than skipping them: the search filter reveals a matching member
    # directly, which is how you reach a grouped tool by name without going
    # through the group; and the grey-out pass can compute the group's own
    # state from the members it already has in front of it.
    $seen = @{}
    foreach ($tool in $Tools) {
        $grp = Get-ToolGroupName -Tool $tool
        if ($grp -and -not $seen.ContainsKey($grp)) {
            $seen[$grp] = $true
            $gb = New-Object System.Windows.Forms.Button
            $gb.Text   = $grp
            $gb.Size   = New-Object System.Drawing.Size($script:BtnW, $script:BtnH)
            $gb.Margin = New-Object System.Windows.Forms.Padding($script:BtnPad)
            $gb.Tag    = [PSCustomObject]@{ IsToolGroup = $true; GroupName = $grp }
            $gb.Add_Click({ Show-ToolGroupChooser -GroupName $this.Tag.GroupName })
            $Panel.Controls.Add($gb)
        }

        $btn = New-Object System.Windows.Forms.Button

        # Sized from the script constants, not from literals.
        #
        # This said Size(150, 32) and Padding(4) while $script:BtnW/BtnH/BtnPad
        # sat three hundred lines above holding the same numbers -- so the
        # constants were decorative, the panel height was computed from one set
        # of values and the buttons drawn from another, and raising BtnH changed
        # the panels without changing a single button.
        #
        # The height went 32 -> 38 because a two-line caption did not fit.
        # 'Mailbox / Onedrive Perms' wraps, and at 32px the first line was
        # clipped along its top edge.
        $btn.Text = $tool.Name
        $btn.Size = New-Object System.Drawing.Size($script:BtnW, $script:BtnH)
        $btn.Margin = New-Object System.Windows.Forms.Padding($script:BtnPad)
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
        # Hidden, not skipped -- see the note above the group button.
        if ($grp) { $btn.Visible = $false }
        $Panel.Controls.Add($btn)
    }
}

function Build-ToolButtons {
    # Membership is worked out ONCE, from the whole tool list, before any panel
    # is built. Deriving it per panel would scatter a group across sections
    # silently if two members disagreed about which section they belong to.
    $script:ToolGroups = @{}
    foreach ($t in @($script:AllTools)) {
        $g = Get-ToolGroupName -Tool $t
        if (-not $g) { continue }
        if (-not $script:ToolGroups.ContainsKey($g)) { $script:ToolGroups[$g] = @() }
        $script:ToolGroups[$g] += $t
    }

    foreach ($name in $script:SectionNames) {
        $panel = $script:SectionPanels[$name]
        if (-not $panel) { continue }
        $panel.Controls.Clear()
        $tools = @($script:AllTools | Where-Object { (Get-ToolSection $_) -eq $name })
        Add-ToolButtonsToPanel -Panel $panel -Tools $tools `
            -EmptyText "No tools in '$name'. Use ToolManager.ps1 to add some."
    }
}

function Update-ConsoleSearch {
    <#
        Filter the tool buttons by caption.

        HIDES rather than greys. Greying is already taken -- it means "this tool
        cannot run right now", with a tooltip explaining which of three reasons.
        A second meaning on the same signal would make both useless, and a grid
        of fifty dimmed buttons is not a search result. A tool that does not
        match simply is not there until the box is cleared.

        Visible is safe to own here because nothing else touches it: capability
        blocking and object-type mismatch both work through Enabled. That
        separation is why search and grey-out can coexist without a combined
        state machine -- checked before this was written, not assumed.

        SECTIONS WITH NO MATCH COLLAPSE ENTIRELY, label included, and everything
        below moves up. A heading above an empty box reads as "the tools in here
        vanished"; no heading reads as "nothing in here matched", which is what
        actually happened.

        Matching is a plain case-insensitive substring on the caption, not a
        regex. An operator typing 'perms' is searching, not writing a pattern,
        and a stray bracket should not throw.
    #>
    param([string]$Text)

    $q = "$Text".Trim()
    $filtering = [bool]$q

    $shown = 0
    $total = 0
    $y = $script:SectionStartY

    foreach ($name in $script:SectionNames) {
        $panel = $script:SectionPanels[$name]
        $lbl   = $script:SectionLabels[$name]
        if (-not $panel) { continue }

        $secVisible = 0
        foreach ($ctrl in @($panel.Controls)) {
            # Labels inside a panel are the "no tools in this section" notice,
            # not buttons. They follow the section, not the filter.
            if (-not $ctrl.Tag) { continue }

            # GROUPS INVERT WHILE FILTERING.
            #
            # Unfiltered, the group button is shown and its members are hidden
            # -- that is the whole point of grouping. Filtered, the group button
            # goes and the matching members appear in its place, so a search for
            # 'roomplace' reaches the tool directly instead of stopping at a
            # button called 'Rooms' that gives no hint of what is inside.
            $isGroup = [bool]$ctrl.Tag.PSObject.Properties['IsToolGroup']
            $grouped = (-not $isGroup) -and [bool](Get-ToolGroupName -Tool $ctrl.Tag)

            if (-not $filtering) {
                $ctrl.Visible = -not $grouped
                if ($ctrl.Visible) { $secVisible++ }
                if (-not $isGroup) { $total++ }
                if ($ctrl.Visible -and -not $isGroup) { $shown++ }
                continue
            }

            if ($isGroup) { $ctrl.Visible = $false; continue }
            $total++

            # Script name as well as caption, because the two are often the only
            # thing an operator remembers -- 'not' finds Returning User through
            # NotDeadYet.ps1, which no amount of searching the caption would.
            $hay = "$($ctrl.Text)"
            if ($ctrl.Tag.PSObject.Properties['Script']) { $hay += " $($ctrl.Tag.Script)" }
            $match = ($hay.IndexOf($q, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
            $ctrl.Visible = $match
            if ($match) { $secVisible++; $shown++ }
        }

        $rows = [int]$script:SectionRows[$name]
        if ($rows -lt 1) { $rows = 1 }

        # While filtering, a section only needs as many rows as it is showing.
        # Keeping the declared height would leave a filtered console mostly
        # empty space and push the last section off the bottom.
        if ($filtering) {
            $need = [int][Math]::Ceiling($secVisible / [double][Math]::Max(1, $script:PerRow))
            if ($need -lt 1) { $need = 1 }
            if ($need -lt $rows) { $rows = $need }
        }
        $panelH = ($rows * ($script:BtnH + ($script:BtnPad * 2))) + 12

        $empty = ($secVisible -eq 0)
        if ($lbl) { $lbl.Visible = -not $empty }
        $panel.Visible = -not $empty
        if ($empty) { continue }

        if ($lbl) { $lbl.Location = New-Object System.Drawing.Point(12, $y) }
        $panel.Location = New-Object System.Drawing.Point(12, ($y + 22))
        $panel.Height   = $panelH
        $y += 22 + $panelH + 8
    }

    if ($lblVersion) {
        $lblVersion.Location = New-Object System.Drawing.Point(12, ($y + 4))
        $y += 26
    }

    $form.ClientSize = New-Object System.Drawing.Size(($script:PanelW + 24), ($y + 8))
    $work = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea
    if ($form.Height -gt $work.Height) { $form.Height = $work.Height }
    if ($form.Width  -gt $work.Width)  { $form.Width  = $work.Width }

    if ($filtering) {
        $lblFindCount.Text = "$shown of $total"
    } else {
        $lblFindCount.Text = ''
    }
}

function Update-SectionPanelWidth {
    <#
        Widen every panel if any of them actually grew a scrollbar.

        This ASKS rather than predicts, which is the whole point. Counting tools
        against declared capacity is a good guess and wrong in the one case that
        matters: capability-hidden tools are removed after the count would have
        been taken, so a predicted overflow can be an overflow that never
        happens -- and the console would carry a scrollbar's width of dead space
        for a scrollbar that was never drawn.

        Widening is GLOBAL. One panel a scrollbar's width narrower than its
        neighbours is the original layout bug: it wraps its last button and
        looks special-cased when it is only the fullest.

        NO OSCILLATION. The scrollbar exists because there are more ROWS than
        the section declared, and adding width does not remove a row: the extra
        is one scrollbar wide, always far less than the width of a button, so it
        can never admit another column and never reduce the row count. The
        widened layout therefore still needs the scrollbar, which is what makes
        one pass enough.

        Called twice on purpose -- see the call sites.
    #>
    if (@($script:SectionPanels.Values).Count -eq 0) { return }

    $need = $false
    foreach ($flow in $script:SectionPanels.Values) {
        if (-not $flow) { continue }
        try {
            # Force the handle and a layout pass. Scrollbar visibility is
            # decided during layout, and before the form is shown neither has
            # necessarily happened -- reading the property first would report
            # the state of a panel that has not been arranged yet.
            $null = $flow.Handle
            $flow.PerformLayout()

            # Two ways of asking, because AutoScroll reports through both and
            # they do not always agree before the form is visible.
            if ($flow.VerticalScroll.Visible) { $need = $true }
            elseif ($flow.DisplayRectangle.Height -gt $flow.ClientSize.Height) { $need = $true }
        } catch { }
    }

    $want = Get-ConsolePanelWidth -ButtonWidth      $script:BtnW `
                                  -ButtonPad        $script:BtnPad `
                                  -ButtonsPerRow    $script:PerRow `
                                  -ReserveScrollBar:$need
    if ($want -eq $script:PanelW) { return }

    $script:PanelW = $want
    $form.SuspendLayout()
    try {
        foreach ($flow in $script:SectionPanels.Values) {
            if ($flow) { $flow.Width = $script:PanelW }
        }
        # Everything else that was sized from PanelW has to follow, or the input
        # and output boxes stop lining up with the button grid -- which is the
        # same lopsided look this was written to remove, just moved.
        $txtInput.Width   = $script:PanelW
        $lblResult.Width  = $script:PanelW
        $txtOut.Width     = $script:PanelW
        $lblVersion.Width = $script:PanelW

        $form.ClientSize = New-Object System.Drawing.Size(
                                ($script:PanelW + 24), $form.ClientSize.Height)

        # Re-apply the screen clamp. Growing the form here can push it past the
        # working area on a small display, and the cap set at build time was
        # calculated against the narrower width.
        $work = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea
        if ($form.Width -gt $work.Width) { $form.Width = $work.Width }
    } finally {
        $form.ResumeLayout()
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

        # One label writer, so the four status labels cannot disagree about what
        # a state means. Three states, not two:
        #
        #   Connected      green
        #   Not Connected  red    -- present in this tenant, unreachable now
        #   n/a            grey   -- this tenant does not have it at all
        #
        # The third used to render as red "Not Connected", which in a cloud-only
        # tenant is a permanent alarm about a server that was never supposed to
        # exist. Red is for things worth fixing.
        $setState = {
            param($Lbl, $State, $Prefix)
            if (-not $State) {
                $Lbl.Text = "${Prefix}: not checked"
                $Lbl.ForeColor = [System.Drawing.Color]::DimGray
                return
            }
            if ($State.PSObject.Properties['NotConfigured'] -and $State.NotConfigured) {
                # Not greyed -- gone. See Update-StatusRowLayout.
                $Lbl.Visible = $false
                return
            }
            $Lbl.Visible = $true
            if ($State.Available) {
                $Lbl.Text = "${Prefix}: Connected"
                $Lbl.ForeColor = $green
            } else {
                $Lbl.Text = "${Prefix}: Not Connected"
                $Lbl.ForeColor = $red
            }
        }

        & $setState $lblExo $exo 'ExO'
        $toolTip.SetToolTip($lblExo, "$($exo.Detail)")

        if (-not $spo) {
            $lblSpo.Text = 'SPO: not checked'
            $lblSpo.ForeColor = [System.Drawing.Color]::DimGray
            $toolTip.SetToolTip($lblSpo, 'Not checked yet -- runs after any tool closes.')
        } else {
            & $setState $lblSpo $spo 'SPO'
            $toolTip.SetToolTip($lblSpo, "$($spo.Detail)")
        }
        # EMS and AD come from the same map, so these labels cannot disagree
        # with the grey-out (which is the whole reason the map exists).
        $ems = $map['EMS']
        $ad  = $map['AD']
        & $setState $lblEms $ems 'EMS'
        $toolTip.SetToolTip($lblEms, "$($ems.Detail)")

        & $setState $lblAd $ad 'AD'
        $toolTip.SetToolTip($lblAd, "$($ad.Detail)")

        # Only counts things this tenant actually has. Otherwise the hint stays
        # on for ever in a cloud-only tenant, telling you to connect to a
        # server that does not exist -- advice that cannot be followed.
        Update-StatusRowLayout -Row @($lblExo, $lblSpo, $lblTenant) -Y 8
        Update-StatusRowLayout -Row @($lblEms, $lblAd, $lblPim, $lblWarn) -Y 30

        $relevant = @($exo, $spo, $ems, $ad) | Where-Object {
            $_ -and -not ($_.PSObject.Properties['NotConfigured'] -and $_.NotConfigured)
        }
        $lblWarn.Visible = @($relevant | Where-Object { -not $_.Available }).Count -gt 0

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

        # A group button is not a tool and must be handled before anything asks
        # it for a Name or ObjectTypes.
        #
        # Its state is the OR of its members: enabled while at least one thing
        # behind it can run. Disabling it whenever any member is blocked would
        # hide four working tools behind one that is not, and enabling it
        # unconditionally would send the operator through a click to reach a
        # dialog of dead buttons.
        if ($tool.PSObject.Properties['IsToolGroup']) {
            $members = @($script:ToolGroups["$($tool.GroupName)"])
            $usable  = 0
            foreach ($m in $members) {
                if (Get-ToolCapabilityBlock -Tool $m) { continue }
                $md = @($m.ObjectTypes | Where-Object { "$_".Trim() })
                if ($showDimming -and $md.Count -gt 0 -and ($md -notcontains $DetectedType)) { continue }
                $usable++
            }
            $ctrl.Enabled = ($usable -gt 0)
            if ($usable -eq 0) {
                $toolTip.SetToolTip($ctrl,
                    "None of the $($members.Count) tools in '$($tool.GroupName)' can run right now. Open it to see why each one is blocked.")
            } else {
                $toolTip.SetToolTip($ctrl,
                    "$($members.Count) tools. $usable available for the current lookup.")
            }
            continue
        }

        $capBlock = Get-ToolCapabilityBlock -Tool $tool

        # An EMPTY ObjectTypes means "applies to anything", not "applies to
        # nothing".
        #
        # The old expression read @() as a list that contains no match, so any
        # tool with no declared types went grey the moment a lookup succeeded --
        # Mailbox / OneDrive Perms, ResourceDuplicator and MB w No activity were
        # permanently disabled after any Look Up, with a tooltip claiming they
        # "apply to: " and then nothing.
        #
        # Declaring no types is how you say a tool is not object-scoped, which
        # is the opposite of what it was being read as. Fail open, same rule the
        # capability classifier follows: an unclassified tool stays enabled.
        $declared = @($tool.ObjectTypes | Where-Object { "$_".Trim() })
        $typeOk   = $true
        if ($showDimming -and $declared.Count -gt 0) {
            $typeOk = ($declared -contains $DetectedType)
        }

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
            # Asks the script, like everything else now. This branch used to
            # read the checkbox, so a tool that DECLARED its own process got the
            # generic tooltip instead of this one -- the message describing the
            # behaviour was shown only to tools that were not getting it.
            $tp = Join-Path $PSScriptRoot $tool.Script
            if ((Test-Path -LiteralPath $tp) -and (Test-ToolDeclaresOwnProcess -ScriptPath $tp)) {
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
        $lblResult.Text = 'Looking up...'

        # THE DIRECTORY LOOKUP IS OPTIONAL. THE OBJECT IS NOT.
        #
        # This required an on-premises hit before it would look at anything, so
        # on a cloud-only tenant every address came back "not found" and every
        # tool stayed greyed -- for a mailbox that exists and that Exchange
        # Online can see perfectly well. The address was never the problem; the
        # question was.
        #
        # Resolve-Target already tries Exchange Online FIRST and only falls back
        # to AD, so it needs no directory at all to find a cloud mailbox. All
        # that was missing was letting it run.
        $hasAd = $true
        try { $hasAd = Test-TenantCapability -Name 'OnPremAD' } catch { $hasAd = $true }

        $envName = $null
        $adServer = $null
        if ($hasAd) {
            $envResult = Find-TargetEnvironment -Smtp $smtp
            if (-not $envResult.Environment) {
                $script:LastSmtp = $null
                $script:LastADServer = $null
                $script:LastEnvironment = $null
                $script:LastType = $null
                $lblResult.Text = "Not found in the on-premises directory -- $smtp"
                if ($envResult.Errors.Count -gt 0) {
                    $txtOut.Text = ($envResult.Errors -join "`r`n`r`n")
                }
                Update-ToolButtonState -DetectedType $null
                return
            }
            $envName  = $envResult.Environment
            $adServer = $envResult.ADServer
        }

        $target = Resolve-Target -Smtp $smtp -ADServer $adServer
        if ($target.Type -eq 'NotFound' -and -not $hasAd) {
            $script:LastSmtp = $null
            $script:LastADServer = $null
            $script:LastEnvironment = $null
            $script:LastType = $null
            $lblResult.Text = "Not found in Exchange Online -- $smtp"
            if ($target.ResolutionErrors.Count -gt 0) {
                $txtOut.Text = ($target.ResolutionErrors -join "`r`n`r`n")
            }
            Update-ToolButtonState -DetectedType $null
            return
        }
        $script:LastSmtp = $smtp
        $script:LastADServer = $adServer
        $script:LastEnvironment = $envName
        $script:LastType = $target.Type

        # Name the directory only when there IS one. On a cloud-only tenant
        # "Found: Mailbox () -- name" reads as a bug in the console rather
        # than as an absent directory.
        $where = if ($envName) { " ($envName)" } else { '' }

        switch ($target.Type) {
            'Mailbox'      { $lblResult.Text = "Found: Mailbox$where -- $smtp" }
            'UnifiedGroup' { $lblResult.Text = "Found: Microsoft 365 Group$where -- $smtp" }
            'OnpremGroup'  { $lblResult.Text = "Found: On-Prem Group$where -- $smtp" }
            default {
                $lblResult.Text = "Object found$where but not confirmed as a Mailbox, Microsoft 365 Group, or on-prem AD group -- $smtp"
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

# Filter as you type. SuspendLayout around it because the handler moves every
# panel and the version label on each keystroke, and WinForms would otherwise
# repaint the whole form per character.
$txtFind.Add_TextChanged({
    $form.SuspendLayout()
    try { Update-ConsoleSearch -Text $txtFind.Text }
    finally { $form.ResumeLayout() }
})

# Escape clears the filter rather than closing the console.
#
# $form.CancelButton is Close, so without this Escape would shut the window --
# which is exactly what a person does when they want to abandon a search, and
# losing the console for it would be infuriating. Only swallowed when there is
# something to clear, so Escape still closes an unfiltered console.
$txtFind.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape -and $txtFind.Text) {
        $_.Handled = $true
        $_.SuppressKeyPress = $true
        $txtFind.Text = ''
    }
})

$btnFindClear.Add_Click({
    $txtFind.Text = ''
    $txtFind.Focus()
})

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

# Panel width is settled HERE, not after Build-ToolButtons.
#
# Update-ConnectionButtonColors ends in Update-CapabilityState, which is what
# hides tools this tenant does not have. Measuring before that runs would ask
# the panels about a button set that is about to change -- and the answer would
# be a scrollbar's width too wide in exactly the tenants that hide the most.
Update-SectionPanelWidth

# Second pass once the window is really on screen.
#
# Scrollbar visibility is decided during layout, and a forced PerformLayout on
# an unshown form is close to the truth but not guaranteed to BE it -- an
# off-by-one in the scrollbar's own width, or a theme that changed a control's
# size after the first pass, and the pre-show answer is stale. This one costs
# nothing when the answer is unchanged: the function returns early when the
# width it computes matches the width already applied.
# GUARDED, because an exception in a Shown handler kills the window.
#
# Shown runs inside the modal loop. A terminating error here does not print and
# does not stop the script -- WinForms tears the form down and ShowDialog
# returns as though the operator had closed it. The console appears for a
# fraction of a second and vanishes, with nothing written anywhere.
#
# That matches a rare, unreproducible "it closed immediately, second launch was
# fine" (2026-08-13) -- unproven as the cause, but this is measurement code
# running at first paint, so it depends on scrollbar state, DPI and theme
# timing, which is exactly the shape of a fault that only sometimes happens.
#
# The catch does not hide anything: the width stays at its pre-show value,
# which is the value that was already good enough to draw with, and the reason
# lands in $txtOut where it can be read and sent on.
$form.Add_Shown({
    Write-StartupBreadcrumb -Text 'form Shown -- laying out panels'
    try {
        Update-SectionPanelWidth
        Write-StartupBreadcrumb -Text 'panels laid out'
    }
    catch {
        # BOTH, and the file first. The output box is nicer to read but only
        # exists while the window does, and a window that dies takes the
        # message with it -- which is exactly how the first version of this
        # guard turned out to be worthless.
        Write-StartupFailure -Stage 'Update-SectionPanelWidth (form Shown)' -ErrorRecord $_
        try {
            $txtOut.Text = ("The console opened, but laying out the button panels failed." +
                [Environment]::NewLine + [Environment]::NewLine +
                "$($_.Exception.Message)" + [Environment]::NewLine + [Environment]::NewLine +
                "Everything still works -- the panels may be slightly the wrong width. " +
                "This is also in Logs\ConsoleStartup.log.")
        } catch { }
    }
})

# One write, after every button has been classified, rather than one per tool.
# The scan itself is what used to make a cold launch slow; persisting it means
# only the FIRST launch on this machine pays, and only again for scripts that
# have actually changed since. Failing to save costs a rescan, never an error.
Save-ScriptCapabilityCache

# Surfaced here, not at migration time -- the move happens before $txtOut
# exists, and a file move the operator never sees is the kind of thing that
# gets blamed for something else three weeks later.
if ($script:ConfigMigrationNotes -and $script:ConfigMigrationNotes.Count -gt 0) {
    # Guarded for the same reason as the handler above: a Shown handler that
    # throws takes the window with it, and this one runs only on the FIRST
    # launch after a config migration -- the launch least likely to be
    # reproducible and most likely to be blamed on something else.
    $form.Add_Shown({
        try {
            $txtOut.Text = ("Config folder created -- existing settings moved in:" +
                [Environment]::NewLine +
                (($script:ConfigMigrationNotes | ForEach-Object { "  $_" }) -join [Environment]::NewLine) +
                [Environment]::NewLine +
                "Update rule from here: replace the .ps1 files, leave Config\ alone.")
        } catch { }
    })
}

Set-FormScreenCorner -Form $form -Corner 'TopLeft'

$script:PendingConnect = $null
# The console's own '?', in the same corner as every tool's. x = ClientSize-40
# lands at 784..812, which is clear -- lblTenant ends at 660.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Admin Console' | Out-Null

Update-PimLabel -Lbl $lblPim -Tip $toolTip

# One tick a minute, doing arithmetic on two timestamps and setting a label.
# Deliberately not shorter, and deliberately no DoEvents anywhere near it --
# see 14c on why extra message-loop activity in this console is treated with
# suspicion.
#
# THE CONTROLS RIDE ON THE TIMER'S .Tag, and the handler must not name them.
#
# A tick scriptblock is not a closure: PowerShell resolves its variables
# against whatever scope is current WHEN THE TICK FIRES. The console launches
# tools with '& $scriptPath', so while a tool's ShowDialog() pumps messages on
# this thread the current scope is the TOOL's -- and a tool that happens to use
# the name $lblPim shadows the console's label for the duration.
#
# That is not hypothetical: it is why this console threw
# "The property 'Visible' cannot be found on this object" once a minute while
# Dashboard was open in a tenant without PIM. See Update-PimLabel.
#
# Tag is the same fix Add-ReportFileControls and Add-ToolHelpButton use, and
# for the same underlying reason: the scope a scriptblock RUNS in is not the
# scope it was WRITTEN in. GetNewClosure is not the answer here -- it would
# bind this scriptblock to a dynamic module and Update-PimLabel, which is
# dot-sourced into script scope, would stop resolving.
$pimTimer = New-Object System.Windows.Forms.Timer
$pimTimer.Interval = 60000
$pimTimer.Tag      = [PSCustomObject]@{ Lbl = $lblPim; Tip = $toolTip }
$pimTimer.Add_Tick({
    $c = $this.Tag
    if (-not $c) { return }
    Update-PimLabel -Lbl $c.Lbl -Tip $c.Tip
})
$pimTimer.Start()
$form.Add_FormClosed({ $pimTimer.Stop(); $pimTimer.Dispose() })

# The two breadcrumbs that localise the fault.
#
# "about to ShowDialog" with nothing after it = the window never opened, and
# everything above this line is suspect. "ShowDialog returned" a fraction of a
# second later = the window DID open and something inside the message loop tore
# it down. Those are different investigations, and without these two lines
# there is no way to tell them apart after the fact.
Write-StartupBreadcrumb -Text 'about to ShowDialog'
[void]$form.ShowDialog()
Write-StartupBreadcrumb -Text 'ShowDialog returned (window closed)'

# Past this line the modal loop is gone, so a sign-in window can take input.
if ($script:PendingConnect) {
    $pending = $script:PendingConnect
    $script:PendingConnect = $null

    $path = Join-Path $PSScriptRoot $pending.Script
    if (Test-Path -LiteralPath $path) {
        Write-Host ''
        Write-Host "Closing the console to sign in -- it will reopen when you are done." -ForegroundColor Cyan

        # ShowDialog installs a WindowsFormsSynchronizationContext on this
        # thread and leaves it there after the form is gone. An interactive
        # sign-in draws its own window, and MSAL posts that window onto whatever
        # context it finds -- so it lands on a pump that closed with the form
        # and fails with "the current thread is not in a single-threaded
        # apartment".
        #
        # Cleared HERE as well as inside the connect script. The script defends
        # itself because it can be run directly; this line means the console
        # never hands it the problem in the first place. Nothing below needs the
        # context -- the console is rebuilt from scratch when it reopens.
        [System.Threading.SynchronizationContext]::SetSynchronizationContext($null)

        try {
            & $path
        } catch {
            Write-Host (Format-ErrorDetail -ErrorRecord $_ -Context "Running $($pending.Name)") -ForegroundColor Red
        }
    } else {
        Write-Host "'$($pending.Name)' points to '$($pending.Script)', which does not exist." -ForegroundColor Red
    }

    # ASK TO BE RE-OPENED. Do not re-open here.
    #
    # This used to be '& $PSCommandPath' -- the script calling itself, which
    # left every connect one stack frame deeper and needed a depth guard that
    # eventually told the operator to relaunch by hand. The trampoline at the
    # top of this file owns the loop now, so setting a flag and falling off the
    # end of the script is the whole mechanism.
    #
    # Falling through rather than returning MATTERS: the session-marker cleanup
    # below still has to run on the way out. The trampoline sets it again on
    # the next pass, and nothing executes in the gap between.
    $global:ToolkitConsoleRelaunch = $true
}

# THE MARKER DIES WITH THE CONSOLE.
#
# Set at the top, cleared here, and this line is the whole reason the guard is
# worth anything on a shared machine.
#
# The console runs INSIDE the shell that started it -- there is no separate
# console process -- so an environment variable it sets survives it closing.
# Without this line, launching the console once left that window permanently
# marked: every script typed at that prompt afterwards passed the guard, for as
# long as the window stayed open. Confirmed live on 2026-08-13, which is how it
# was found.
#
# CLEARED HERE, NOT IMMEDIATELY AFTER ShowDialog, because the connect path runs
# in between: the console closes, a guarded Connect-*.ps1 runs, and then the
# console relaunches itself. Clearing any earlier blocks the connect script and
# breaks sign-in.
#
# The relaunch loop is safe. Each pass sets the marker on entry and clears it
# here on the way out, and the trampoline starts the next pass immediately --
# so the marker is only empty in the gap between two passes, where no script
# runs. The last pass leaves it empty, which is the state that matters.
#
# HONEST LIMIT: if the console dies on an unhandled error this line never runs
# and the marker persists until that window closes. That is a speed bump losing
# its bump in a rare case, not a hole in something load-bearing -- see
# Test-ConsoleLaunch on what this is and is not.
$env:ToolkitConsoleSession = $null
Write-StartupBreadcrumb -Text 'console session marker cleared'
#endregion
