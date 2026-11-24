#Requires -Version 5.1
<#
    AdminConsole.ps1
    Layer 1 launcher -- Mailbox / Group lookup + dynamic tool menu.

    Author   : Glenn Armbruster
    Revised  : 2026-09-13
#>
# TOOLKIT-HELP-RUNS: whatever you connect it to. The status row at the top of this window shows what it currently has.
# TOOLKIT-HELP-BEGIN
# The launcher. Look up a mailbox or group, then use the tools below.
#
# TWO BOXES AT THE TOP, AND THEY DO DIFFERENT THINGS. Each is labelled at the
# start of its own row:
#
#   Unified lookup   a name, alias or address -- Look Up opens Unified Lookup
#                    with it. A full address also points the tools below at
#                    that one object, which is what greys the ones that do
#                    not apply.
#   Find Tool        filters the buttons on this window. It does not search
#                    the tenant. Show all clears it, and so does Escape.
#
# Ctrl+F jumps to Find Tool from anywhere. The console closes from its own
# title bar -- there is no Close button, and Escape no longer closes it.
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
    [string]$ToolsConfigPath = '',

    # SAFE MODE: open with the shipped look and the shipped geometry, ignoring
    # Themes.json and Sections.json for this run only. Nothing is written and
    # nothing is lost -- it exists so a console made unreadable by a theme can
    # still be opened and repaired, which a Reset button buried inside that
    # same unreadable console cannot help with.
    #
    #     .\AdminConsole.ps1 -SafeMode
    #
    # The trampoline re-invokes with @PSBoundParameters, so this survives the
    # relaunch a theme change triggers without any extra plumbing.
    [switch]$SafeMode,

    # ZOOM DIAGNOSTIC: write the zoom factor and the real measured geometry of
    # every header control to Logs\ConsoleStartup.log, once, after the window
    # has settled.
    #
    #     .\AdminConsole.ps1 -ZoomDiag
    #
    # WHY IT EXISTS. On 2026-09-14 the same console was crisp for one engineer
    # and crushed for another on the SAME server, and three plausible causes
    # were argued from screenshots before any of them was measured: display
    # scaling, the Windows text-size slider, a per-app DPI override. All three
    # were wrong. What settled it was the factor and the pixel heights, and
    # neither is visible from outside the process.
    #
    # Off by default and writes nothing when absent. It is a diagnostic, not
    # telemetry.
    [switch]$ZoomDiag
)

# THE SWITCH DECIDES ON EVERY LAUNCH -- INCLUDING WHEN IT IS ABSENT.
#
# $env: variables belong to the POWERSHELL WINDOW, not to this script. The
# first version only ever SET the flag, so safe mode was sticky: run the
# console once with -SafeMode, close it, run it again without the switch from
# the same prompt, and it came back in safe mode with nothing on screen or on
# disk explaining why. Only opening a new PowerShell window cleared it.
#
# So the off path has to be as real as the on path. This is the same family as
# the theme cache and the settings cache -- state that outlives the script
# inside one process -- and it is the third time today that shape has bitten.
#
# REMOVED rather than blanked, so a tool launched from here does not inherit
# an empty variable and raise the same question one level down.
#
# Both branches run BEFORE Common.ps1 is dot-sourced, because the flag has to
# be settled before anything can ask for a theme.
if ($SafeMode) {
    $env:ToolkitSafeMode = '1'
} else {
    try { Remove-Item -Path 'Env:\ToolkitSafeMode' -ErrorAction Stop } catch { $env:ToolkitSafeMode = '' }
}
# The resolved theme is cached in a global that ALSO outlives the script, so
# flipping the mode has to drop it -- otherwise the first lookup after the
# switch answers from the wrong side of it.
$global:AdminConsoleTheme = $null

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

function Write-ConsoleZoomDiagnostic {
    <#
        Dump the zoom factor and the header's REAL geometry to the startup log.

        WHAT TO READ IN IT. For each control: the box it ended up with, the
        pixels its own text needs in the font it is actually wearing, and the
        percentage. Anything at or near 100 % is clipped or about to be -- that
        is the whole question, and it cannot be answered from a screenshot or
        from outside the process.

        MEASURES, NEVER ASSUMES. Every number here is read off the live control
        at the moment it is written: $c.Width, $c.Font, TextRenderer against
        that font. No design literals, because the design literals are exactly
        what turned out not to be true.

        Nothing here may throw. A diagnostic that takes the console down with it
        is worse than no diagnostic, and this one runs on the operator's machine
        precisely when something is already wrong.
    #>
    if (-not $ZoomDiag) { return }
    try {
        $z = $script:Zoom
        $work = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea
        Write-StartupBreadcrumb -Text '--- ZOOM DIAGNOSTIC ---------------------------------------'
        Write-StartupBreadcrumb -Text ("screen working area : {0}x{1}" -f $work.Width, $work.Height)
        Write-StartupBreadcrumb -Text ("form ClientSize     : {0}x{1}" -f $form.ClientSize.Width, $form.ClientSize.Height)
        if ($z) {
            Write-StartupBreadcrumb -Text ("design WxH          : {0}x{1}   StartY0={2}" -f $z.DesignW, $z.DesignH, $z.StartY0)
            Write-StartupBreadcrumb -Text ("zoom factor applied : {0}   (header factor = max(f,1.0))" -f $z.Applied)
            Write-StartupBreadcrumb -Text ("grid BtnW/BtnH/Pad  : {0}/{1}/{2}   PerRow={3}" -f $script:BtnW, $script:BtnH, $script:BtnPad, $script:PerRow)
        } else {
            Write-StartupBreadcrumb -Text 'zoom not initialised -- no factor to report'
        }

        # The header controls, by the names used in this file, plus whatever
        # caption each is carrying at the time.
        $named = @(
            @{ N = 'lblTitle';      C = $lblTitle }
            @{ N = 'lblUnified';    C = $lblUnified }
            @{ N = 'txtInput';      C = $txtInput }
            @{ N = 'btnRun';        C = $btnRun }
            @{ N = 'lblFind';       C = $lblFind }
            @{ N = 'txtFind';       C = $txtFind }
            @{ N = 'btnFindClear';  C = $btnFindClear }
            @{ N = 'lblFindCount';  C = $lblFindCount }
            @{ N = 'btnView';       C = $btnView }
            @{ N = 'chkArrange';    C = $chkArrange }
            @{ N = 'lblResult';     C = $lblResult }
        )
        Write-StartupBreadcrumb -Text 'control          box WxH     font        text WxH    w%   h%   verdict'
        foreach ($n in $named) {
            $c = $n.C
            if (-not $c -or $c.IsDisposed) {
                Write-StartupBreadcrumb -Text ("  {0,-14} (absent)" -f $n.N)
                continue
            }
            $sz = [System.Windows.Forms.TextRenderer]::MeasureText("$($c.Text)", $c.Font)
            $wp = if ($c.Width  -gt 0) { [int]($sz.Width  / $c.Width  * 100) } else { 0 }
            $hp = if ($c.Height -gt 0) { [int]($sz.Height / $c.Height * 100) } else { 0 }
            $verdict = if ($wp -gt 100 -or $hp -gt 100) { 'CLIPPED' } elseif ($wp -gt 90 -or $hp -gt 90) { 'tight' } else { 'ok' }
            Write-StartupBreadcrumb -Text ("  {0,-14} {1,4}x{2,-4}  {3,-6}{4,5}pt {5,4}x{6,-4} {7,4} {8,4}   {9}" -f `
                $n.N, $c.Width, $c.Height, $c.Font.Name.Substring(0, [Math]::Min(6, $c.Font.Name.Length)), `
                [Math]::Round($c.Font.SizeInPoints, 2), $sz.Width, $sz.Height, "$wp%", "$hp%", $verdict)
        }
        Write-StartupBreadcrumb -Text '--- END ZOOM DIAGNOSTIC -----------------------------------'
    } catch {
        try { Write-StartupBreadcrumb -Text "zoom diagnostic failed: $($_.Exception.Message)" } catch { }
    }
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

# Timed, because Common.ps1 is dot-sourced again by every tool the console
# launches in-process; the number shown in the status row at startup is what
# each tool launch pays before it draws. (See ARCHITECTURE 43 for the
# skip-if-loaded idea and why it was not done blind.)
$script:CommonLoadWatch = [System.Diagnostics.Stopwatch]::StartNew()
try {
    . (Join-Path $PSScriptRoot 'Common.ps1')
    $script:CommonLoadWatch.Stop()
    Write-StartupBreadcrumb -Text ('Common.ps1 dot-sourced in {0} ms' -f $script:CommonLoadWatch.ElapsedMilliseconds)
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
        $ans = Show-ToolDialog($msg, 'Tenant not configured',
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
        $null = Show-ToolDialog($msg, 'Tenant not configured',
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

# DPI SCALING, DECLARED HERE AND NOT LEFT TO Set-ConsoleTheme. Reported
# 2026-09-14: on a 175 % display the console's header was unreadable -- "Look Up"
# rendered as "Look", "Arrange" as "Arra", and the result line was sliced through
# the middle of the glyphs -- while the button grid and every TOOL window on the
# same machine were perfect.
#
# THAT SPLIT IS THE WHOLE DIAGNOSIS. Common.ps1 makes the process per-monitor DPI
# aware at dot-source time, so text renders at the real dpi; the LAYOUT only
# follows if the form also declares it was designed at 96. Set-ConsoleTheme does
# declare it -- but only `if (-not $Form.IsHandleCreated)`, because both
# properties are ignored once the window exists. A tool calls the theme
# immediately before ShowDialog and gets it. THIS form is built, measured,
# populated and laid out across thousands of lines first, so by the time the
# theme runs the guard can already be false -- and then the console alone is
# 96 dpi geometry wrapped around 175 % text.
#
# Set at creation, the guard cannot have gone stale: nothing has touched this
# form yet. It is also the only version that survives SAFE MODE, which skips
# Set-ConsoleTheme entirely -- so the mode that exists to rescue an unreadable
# console was the least readable of all on that machine.
#
# AT 96 DPI THIS IS EXACTLY A NO-OP: the factor is 96/96, WinForms scales by
# 1.0, and nothing on a 100 % display moves by a pixel.
#
# This is NOT the same thing as Enable-FormAutoScale, which this form opts out
# of further down with NoAutoScale. That one scales a whole window by one factor
# to fit its screen; this is WinForms' own dpi correction. The console wants the
# second and not the first, because it reflows instead of scaling uniformly.
try {
    $form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
} catch { }

$form.Size = New-Object System.Drawing.Size(840, 808)
$form.StartPosition = "Manual"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblExo = New-Object System.Windows.Forms.Label
$lblExo.Tag = 'NoTheme'
$lblExo.Location = New-Object System.Drawing.Point(12, 8)
$lblExo.Size = New-Object System.Drawing.Size(156, 20)
$lblExo.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$lblExo.Text = "ExO"
$form.Controls.Add($lblExo)

$lblSpo = New-Object System.Windows.Forms.Label
$lblSpo.Tag = 'NoTheme'
$lblSpo.Location = New-Object System.Drawing.Point(170, 8)
$lblSpo.Size = New-Object System.Drawing.Size(156, 20)
$lblSpo.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$lblSpo.Text = "SPO"
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
$lblEms.Text = "EMS"
$form.Controls.Add($lblEms)

$lblAd = New-Object System.Windows.Forms.Label
$lblAd.Tag = 'NoTheme'
$lblAd.Location = New-Object System.Drawing.Point(170, 30)
$lblAd.Size = New-Object System.Drawing.Size(156, 20)
$lblAd.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$lblAd.Text = "AD"
$form.Controls.Add($lblAd)

# "Use Connect below." REMOVED 2026-09-13, at the operator's request, as
# redundant -- and it was. By the time it was written the status strip at the
# bottom already carried a live pill per service (ExO, SPO, EMS, AD, Tenant),
# each of which names what is not connected and does it in colour. A second
# italic sentence in the header repeated that in words, and the Connect section
# is the first thing on the grid, so it was also pointing at something nobody
# was struggling to find.
#
# Its Visible flag was driven from the capability probe in two places; both are
# gone with it. Nothing else referenced it.

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
# The console's name, where the connection labels used to be; those now
# live on the status strip at the bottom (see New-ConsolePill).
# THE NAME COMES FROM CONFIG, NOT FROM THIS LINE.
#
# Settings.ConsoleTitle carries it, because the toolkit's standing rule is
# that no .ps1 names a company or a tenant -- Test-Toolkit flags it, and the
# same file ships to more than one place. Blank falls back to 'Admin Console',
# which is what every installation that has not set one still sees.
#
# The face comes from Settings.ConsoleTitleFont the same way, and is used only
# if it is really installed; see Get-ConsoleTitleFont for why that check is
# not optional.
$lblTitle          = New-Object System.Windows.Forms.Label
$lblTitle.Tag      = 'NoTheme'
# SAFE MODE SAYS SO IN THE TITLE, and that is not decoration. A console that
# is quietly ignoring Themes.json and Sections.json looks exactly like a
# console whose theme has stopped working -- so without this, the tool built
# to end one debugging session starts another.
$lblTitle.Text     = $(if (Test-ToolkitSafeMode) { "$(Get-ConsoleTitleText)  --  SAFE MODE" } else { Get-ConsoleTitleText })
$lblTitle.Location = New-Object System.Drawing.Point(12, 12)
$lblTitle.AutoSize = $true
$lblTitle.Font     = (Get-ConsoleTitleFont -Size 15)
$form.Controls.Add($lblTitle)

# "signed in as ..." BESIDE THE TITLE IS GONE, removed 2026-09-13 at the
# operator's request. Get-OperatorDisplayName went with it -- it had one caller,
# this label -- as did Set-ConsoleHeaderOperatorLayout and the
# Settings.ConsoleHeaderOperator switch that gated it.
#
# The operator is still named, bottom right on the status strip, from
# Get-LogOperator -- the same identity the audit log writes. That one stays.
#
# Settings.ConsoleHeaderOperator is now inert. It can come out of Tool Manager's
# Settings dialog whenever that file is next touched; leaving a key that nothing
# reads is untidy but harmless, and editing a second script for it was not what
# was asked for.

$lblPim           = New-Object System.Windows.Forms.Label
$lblPim.Tag       = 'NoTheme'
$lblPim.Location  = New-Object System.Drawing.Point(330, 31)

# AUTOSIZE, not a fixed 170px.
#
# The line holds Groups AND Roles once both timers are running, and 170px fits
# only the first. Roles was not hidden -- it was off the right-hand edge of a
# label too small for its own text, which looks identical to "not set".
#
# Set-ConsoleStatusLayout measures each pill's caption at layout time, so
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
        if (-not $line) {
            # Nothing recorded yet. On a tenant that HAS PIM, say so as a
            # zero rather than vanishing -- a pill that comes and goes moves
            # everything beside it. A tenant without PIM keeps it hidden.
            $hasPim = $false
            try { $hasPim = [bool](Test-TenantCapability -Name 'Pim') } catch { }
            if (-not $hasPim) { $Lbl.Visible = $false; return }
            $Lbl.Text      = 'PIM: 0:00'
            $Lbl.ForeColor = [System.Drawing.Color]::DimGray
            $Lbl.Visible   = $true
            if ($Tip) { $Tip.SetToolTip($Lbl, 'No activation recorded. The Activate button starts the clock.') }
            return
        }
        $Lbl.Text      = $line
        $Lbl.ForeColor = Get-PimStatusColour
        $Lbl.Visible   = $true
        if ($Tip) {
            $Tip.SetToolTip($Lbl, 'Recorded by the Activate button, not read from Entra. Clear it there if it looks wrong.')
        }
    } catch { $Lbl.Visible = $false }
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

# THE HINT LINE ABOVE THE ROWS IS GONE, removed 2026-09-13. It read "Find a
# name, alias or address -- Look Up opens Unified Lookup (a full address also
# targets the tools below)" and it was written when neither box had a label. The
# rows now say "Unified lookup" and "Find Tool" in front of themselves, and the
# result line under them repeats the full-address part after a search, so the
# sentence was explaining something the window already says twice.
#
# The row's y positions are unchanged -- 68 and 98, as they were with the hint
# above them -- so nothing below moves. The space it leaves is deliberate.
# TWO LABELLED ROWS, CHANGED 2026-09-13. The header had two input boxes and
# nothing said which was which: a full-width box on one row, then Look Up /
# Clear / Close with "Find tool:" stranded beside them on the next. Reported as
# "it's hard to tell what's what", which is the only evidence that matters.
#
# Now each row reads left to right as one sentence:
#
#     Unified lookup  [................................]  [Look Up ]
#     Find Tool       [................................]  [Show all]
#
# The two boxes share a left edge and a width and the two buttons share a left
# edge, so the rows line up as columns. None of those three numbers is a
# literal -- Set-ConsoleFindRowLayout derives all of them, because the theme
# and the DPI scaling both resize these controls after this block runs.
$lblUnified = New-Object System.Windows.Forms.Label
$lblUnified.Text = 'Unified lookup'
$lblUnified.Location = New-Object System.Drawing.Point(12, 73)
$lblUnified.Size = New-Object System.Drawing.Size(104, 18)
$form.Controls.Add($lblUnified)

$txtInput = New-Object System.Windows.Forms.TextBox
$txtInput.Location = New-Object System.Drawing.Point(124, 69)
# 24 high to match the find box, not 20. They sit one above the other now, and
# two boxes of different heights beside labels of the same size reads as one of
# them being broken.
$txtInput.Size = New-Object System.Drawing.Size(400, 24)
$txtInput.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtInput)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Look Up"
# 90x28 to match Show all below it, and placed at the END of the input box.
$btnRun.Location = New-Object System.Drawing.Point(532, 68)
$btnRun.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnRun)

# CLEAR AND CLOSE REMOVED 2026-09-13 at the operator's request.
#
# WHAT WENT WITH CLEAR, because it was not only a box-emptier: its handler also
# reset the lookup target -- LastSmtp, LastType, LastADServer, LastEnvironment
# -- and re-ran Update-ToolButtonState, which is what un-greys the tools after
# they have been pointed at one object. There is now no button that does that;
# looking something else up replaces the target, which is how it is normally
# cleared anyway. Say the word if emptying the box should also un-target.
#
# CLOSE was also $form.CancelButton, so Escape closed the console when the find
# box was empty. That is gone too -- Escape now only ever clears the filter,
# and the window closes from its own title bar. See the KeyDown handler.

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
# 'Find Tool', not 'Find tool:' -- it is a row label beside a box now, matching
# 'Unified lookup' above it, rather than a sentence fragment wedged between two
# buttons. The colon went with the change of role.
$lblFind.Text = 'Find Tool'
$lblFind.Location = New-Object System.Drawing.Point(12, 103)
$lblFind.Size = New-Object System.Drawing.Size(104, 18)
$form.Controls.Add($lblFind)

$txtFind = New-Object System.Windows.Forms.TextBox
$txtFind.Location = New-Object System.Drawing.Point(124, 99)
$txtFind.Size = New-Object System.Drawing.Size(400, 24)
$form.Controls.Add($txtFind)

$btnFindClear = New-Object System.Windows.Forms.Button
$btnFindClear.Text = 'Show all'
$btnFindClear.Location = New-Object System.Drawing.Point(532, 98)
$btnFindClear.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnFindClear)

$lblFindCount = New-Object System.Windows.Forms.Label
$lblFindCount.Location = New-Object System.Drawing.Point(710, 103)
$lblFindCount.Size = New-Object System.Drawing.Size(120, 18)
$lblFindCount.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblFindCount.Tag = 'RightAnchored'   # keeps its distance from Arrange under zoom (see Initialize-ConsoleZoom)
$form.Controls.Add($lblFindCount)

# ARRANGE MODE. While ticked, tool buttons drag instead of launch: drop one
# on another to reorder, drop it on a different section to move it. The
# order is written straight back to Tools.json, so ToolManager is no longer
# the only way to lay the console out. A toggle rather than always-on,
# because a click that becomes a drag by accident is worse than a trip to
# ToolManager. See ARCHITECTURE 42.
# A plain Button that toggles, not a CheckBox drawn as one. A CheckBox with
# Appearance=Button keeps painting a system-coloured face under every theme
# and under FlatStyle.Flat drew a stray box over its own caption; a Button
# is themed like every other button on the form. State lives in
# $script:Arrange.On; Set-ArrangeMode below is the one place it changes.
$chkArrange = New-Object System.Windows.Forms.Button
$chkArrange.Text       = 'Arrange'
$chkArrange.Location   = New-Object System.Drawing.Point(766, 98)
$chkArrange.Size       = New-Object System.Drawing.Size(74, 28)
$chkArrange.Tag        = 'Plain'
$form.Controls.Add($chkArrange)
$script:Arrange = @{ On = $false; Drag = $null }

# VIEW. The same tools, drawn a different way. See the views region further
# down for what a view is and how to add one -- the point of the button is that
# adding the third and fourth costs one entry in a list each.
$btnView = New-Object System.Windows.Forms.Button
$btnView.Text     = 'View'
$btnView.Location = New-Object System.Drawing.Point(688, 98)
$btnView.Size     = New-Object System.Drawing.Size(74, 28)
# 'Plain' keeps it off the accent styling; 'RightAnchored' keeps it a fixed
# distance from Arrange when the zoom walk widens the form, the same as the
# count label beside it. Without that it drifts left as the window grows and
# ends up marooned in the middle of the find row.
$btnView.Tag      = 'Plain RightAnchored'
$form.Controls.Add($btnView)

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
# THE OUTPUT BOX IS HIDDEN. Since 2026-09-12 Look Up opens Unified Lookup,
# which is a far better place to read about an object than 140 pixels of
# read-only text under the input row. The control stays -- eleven places
# write to it, the zoom baseline measures it -- but it is not shown, and
# anything written to it is raised as a dialog instead (see the handler
# below the grid build), so an error cannot vanish into an invisible box.
$txtOut.Visible = $false
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
    # SAFE MODE IGNORES THE FILE. Sections.json drives every geometry number in
    # this console -- panel heights, the form's size, the fit at startup -- so
    # a nonsense value here is one of the two ways to end up with a window you
    # cannot work in. Safe mode is how you get back to a console that can edit
    # it.
    if ((Test-ToolkitSafeMode) -or -not (Test-Path -LiteralPath $script:SectionsConfigPath)) { return $default }
    try {
        $cfg = Get-Content -LiteralPath $script:SectionsConfigPath -Raw | ConvertFrom-Json
        if (-not $cfg.Sections -or @($cfg.Sections).Count -eq 0) { return $default }

        # CLAMPED, NOT TRUSTED.
        #
        # These numbers are multiplied into every size on the form, and the
        # file is hand-editable. ButtonsPerRow of 0 divides the layout by
        # nothing; 200 builds a form wider than any monitor and then clamps it
        # to the work area with the buttons off the edge; Rows of 500 makes a
        # panel taller than the screen. None of those is a value anyone MEANT,
        # so honouring them faithfully is not respect for the operator's
        # intent -- it is a transcription error rendered at full size.
        #
        # 1..12 columns, 1..10 rows. The column cap was 10 and Tool Manager's
        # spinner stopped at 8 -- two different limits for one number, and the
        # lower one was the one an operator could see. Raised to 12 on
        # 2026-09-12 because a wide screen genuinely fits it and the Connect
        # section wanted a single row of ten. Past 12 the grid stops being
        # easier to read than the list view, and the window stops fitting a
        # normal monitor, so the cap stays.
        $per = 5
        try { $per = [int]$cfg.ButtonsPerRow } catch { }
        if ($per -lt 1)  { $per = 1 }
        if ($per -gt 12) { $per = 12 }
        Add-Member -InputObject $cfg -NotePropertyName ButtonsPerRow -NotePropertyValue $per -Force

        # A SECTION WITH NO NAME, OR A DUPLICATE NAME, IS DROPPED.
        #
        # Not tidiness: the panels are held in a hashtable keyed on the name
        # ($script:SectionPanels), so two sections called the same thing means
        # the second silently overwrites the first and one panel is orphaned --
        # built, filled with buttons, and never laid out or shown.
        $seen = @{}
        $keep = @()
        foreach ($sec in @($cfg.Sections)) {
            $nm = "$($sec.Name)".Trim()
            if (-not $nm) { continue }
            if ($seen.ContainsKey($nm.ToLower())) { continue }
            $seen[$nm.ToLower()] = $true
            $r = 1
            try { $r = [int]$sec.Rows } catch { }
            if ($r -lt 1)  { $r = 1 }
            if ($r -gt 10) { $r = 10 }
            # OPTIONAL ACCENT, a hex colour per section. Validated here so a
            # typo in the file becomes "no accent" rather than an exception in
            # the middle of building the window. Anything that is not #RGB or
            # #RRGGBB is dropped silently, which is the same clamp-and-carry-on
            # rule the row count above follows.
            $acc = ''
            try {
                $raw = "$($sec.Accent)".Trim()
                if ($raw -match '^#([0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})$') { $acc = $raw }
            } catch { }
            $keep += [PSCustomObject]@{ Name = $nm; Rows = $r; Accent = $acc }
        }
        if ($keep.Count -eq 0) { return $default }
        Add-Member -InputObject $cfg -NotePropertyName Sections -NotePropertyValue $keep -Force
        return $cfg
    } catch { return $default }
}

$script:SectionCfg    = Get-SectionDefinitions
$script:SectionPanels = @{}
# Set by the Arrange drag handler when a drop grew a section; Save Arrangement
# writes the new row count to Sections.json and clears it.
$script:SectionRowsDirty = $false
# Section name -> $true while folded. Seeded from the operator's state file
# once, below, after the sections exist; read by both reflow paths.
$script:CollapsedSections = @{}

# Labels and declared row counts are kept alongside the panels so the search
# filter can collapse a whole section when nothing in it matches. Without the
# label reference a filtered console shows a heading over an empty box, which
# reads as "these tools disappeared" rather than "nothing here matched".
$script:SectionLabels = @{}
$script:SectionRows   = @{}
$script:SectionAccents = @{}
# 164, not 308: the 140px output box below lblResult is hidden (see the
# control), so the sections start directly under the result line.
$script:SectionStartY = 164
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
        -Font (Get-ThemeUiFont -Size $form.Font.Size) `
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
#
# $txtInput IS NO LONGER ONE OF THEM. From 2026-09-13 it is the same width as
# the find box, not the width of the button grid, so it is owned by the block
# below and by Set-ConsoleFindRowLayout. Setting it to PanelW here would make
# it full width for the instant before that block narrowed it again.
$lblResult.Width = $script:PanelW
$txtOut.Width    = $script:PanelW

# Both input rows are laid out from the RIGHT edge of the panel so they stay
# put as ButtonsPerRow changes the console width -- fixed left offsets would
# strand them mid-form at 7 columns and overrun at 4.
#
# THE TWO-LAYOUT BRANCH IS GONE, AND THAT IS A SIMPLIFICATION, NOT A LOSS.
# It existed because the find box used to share the button row with Look Up /
# Clear / Close, and at four or five columns there was no room to share; the
# fallback dropped the find box onto a row of its own and pushed everything
# below it down by 34px. As of 2026-09-13 the find box HAS a row of its own at
# every width -- that is the point of the change -- so there is nothing to fall
# back to and nothing to shift. The box simply gets narrower, with a floor.
$findRight = 12 + $script:PanelW

# THE RIGHT END OF THE ROW IS NOW LAID OUT BY SUBTRACTION, RIGHT TO LEFT,
# AND THAT FIXES A REAL OVERLAP RATHER THAN JUST MAKING ROOM FOR View.
#
# The old line put the count label at $findRight - 118 with a width of 50, so
# its right edge landed at $findRight - 68 -- six pixels INSIDE the Arrange
# button, which starts at $findRight - 74. That is the "Arrange is partially
# covered" report, and it was never going to be found by looking at Arrange:
# the arithmetic that covers it is in the label's line.
#
# Every offset below is derived from the control to its right plus a gap, so
# adding another button means adding one line, not re-deriving four constants.
$gap       = 8
$arrangeX  = $findRight - 74
$viewX     = $arrangeX - $gap - 74
$countW    = 50
$countX    = $viewX - $gap - $countW
$clearX    = $countX - $gap - 90

# THE LABEL COLUMN IS ONE WIDTH FOR BOTH ROWS, so the boxes share a left edge.
# 104 is a starting value only: Set-ConsoleFindRowLayout re-measures both
# captions once the theme has settled the typeface and re-derives all of this.
$labelW  = 104
$boxLeft = 12 + $labelW + $gap

# 70% OF WHAT WOULD FIT, asked for on 2026-09-13 as "30% less wide". Neither box
# ever needed the full span -- an address is 30-odd characters and a tool name is
# fewer -- and the row was reading as one long empty rule across the window.
#
# The trailing button follows the box rather than staying pinned to the right,
# because "at the end of the input box" is where it was asked to be. The count,
# View and Arrange are still right-anchored, so the slack lands between the
# button and them.
$script:ConsoleBoxWidthShare = 0.7
$boxW = [int](((($clearX - $gap) - $boxLeft)) * $script:ConsoleBoxWidthShare)
if ($boxW -lt 120) { $boxW = 120 }

# Row 1 -- Unified lookup.
$lblUnified.Location   = New-Object System.Drawing.Point(12, 73)
$lblUnified.Size       = New-Object System.Drawing.Size($labelW, 18)
$txtInput.Location     = New-Object System.Drawing.Point($boxLeft, 69)
$txtInput.Size         = New-Object System.Drawing.Size($boxW, 24)
$btnRun.Location       = New-Object System.Drawing.Point(($boxLeft + $boxW + $gap), 68)

# Row 2 -- Find Tool, and the three right-anchored controls that share it.
$lblFind.Location      = New-Object System.Drawing.Point(12, 103)
$lblFind.Size          = New-Object System.Drawing.Size($labelW, 18)
$txtFind.Location      = New-Object System.Drawing.Point($boxLeft, 99)
$txtFind.Size          = New-Object System.Drawing.Size($boxW, 24)
$btnFindClear.Location = New-Object System.Drawing.Point(($boxLeft + $boxW + $gap), 98)
$lblFindCount.Location = New-Object System.Drawing.Point($countX, 103)
$lblFindCount.Size     = New-Object System.Drawing.Size($countW, 18)
$btnView.Location      = New-Object System.Drawing.Point($viewX, 98)
$chkArrange.Location   = New-Object System.Drawing.Point($arrangeX, 98)

function Set-SectionPanelGrid {
    <#
        Make a section panel hold EXACTLY ButtonsPerRow buttons per row.

        THE COLUMN COUNT USED TO BE EMERGENT AND THAT WAS THE BUG. A
        FlowLayoutPanel wraps when the next control will not fit, so the
        number of columns was whatever the panel's width happened to allow --
        Sections.json said 10 and the console drew 9 or 12 depending on the
        window, a vertical scrollbar stealing 17px, and how the slack divided.
        Reported 2026-09-12 as "I have it set to 10 columns but it shows 9, or
        12 if I resize".

        ButtonsPerRow is now a HARD limit. The usable width inside the panel
        is set to exactly PerRow cells, so the PerRow+1'th button cannot fit
        and wraps, and the PerRow'th always can. Whatever slack the panel has
        beyond that becomes padding, half each side, so the grid still sits
        centred in the panel.

        AUTOSCROLL IS OFF for the same reason. A vertical scrollbar takes its
        width out of the client area, which silently removed a column from the
        fullest section -- the one section where the operator would notice. The
        row count is computed from the tool count, so there is nothing to
        scroll to; and when a drag overfills a section, Arrange grows a row
        rather than hiding the overflow behind a bar.
    #>
    param([Parameter(Mandatory)]$Panel)
    if (-not $Panel -or $Panel.IsDisposed) { return }
    $per = [int]$script:PerRow
    if ($per -lt 1) { $per = 1 }
    $cell = [int]$script:BtnW + ([int]$script:BtnPad * 2)
    if ($cell -lt 1) { return }

    $Panel.AutoScroll = $false

    # ClientSize, not Width: the border is already out of it.
    $inner = [int]$Panel.ClientSize.Width
    $want  = $cell * $per
    $slack = $inner - $want
    if ($slack -lt 0) { $slack = 0 }
    $left  = [int]($slack / 2)
    $right = $slack - $left
    $Panel.Padding = New-Object System.Windows.Forms.Padding($left, 0, $right, 0)
}

# ---------------------------------------------------------------------------
# SECTION HEADING GEOMETRY -- ONE DEFINITION, FOUR CONSUMERS.
#
# THE HEADING USED TO BE A 20px BOX WITH THE PANEL HARDCODED 22px BELOW IT, and
# that is the bug behind "text / headings are buried under buttons". Nothing
# ever set the label's HEIGHT from the text it had to hold:
#
#   * the theme re-faces every label after the form is built, and the section
#     headings are then made BOLD by the accent pass;
#   * Update-ConsoleZoom scales the heading FONT by the zoom factor and scales
#     the 22px offset with it, but never touches the label's own Height;
#   * so on a high-dpi or zoomed session the font grew to 13pt bold -- 47px of
#     text -- and stayed in a 20px box.
#
# A Label clips its own text at its bounds, so what was actually on screen was
# the top third of every heading, which reads exactly like the panel below has
# been drawn over it. Measured on a 192dpi session: Height=20, needed=47.
#
# THE HEIGHT IS NOW MEASURED FROM THE FONT THE LABEL ACTUALLY HAS, at the
# moment it is placed. That is deliberately not the same as computing it once:
# DPI auto-scaling and the zoom both change the font after the build loop has
# run, and each re-stack is the only point where the current answer is knowable.
$script:SecHeadPad = 6     # leading added under the heading text
$script:SecHeadGap = 6     # heading box -> panel top
$script:SecAfter   = 14    # panel bottom -> next heading (was 8; too tight)
# What Get-ConsoleZoomFit reserves per heading in DESIGN space, as a
# heading-top -> panel-top offset.
#
# MEASURED, NOT GUESSED, and the guess is what put a scrollbar back. This
# started as a flat 30, which was fine for the old 20px box and far too small
# for a measured one: on a 192dpi session the real design offset is ~44 + the
# gap, so the fit reserved too little per section, chose a factor a step too
# large, and the finished layout then overflowed the window it had just been
# told would hold it. The build loop overwrites this with the real figure the
# first time it sizes a heading; the literal is only the value used before
# that has happened.
$script:SecHeadRes = 30
# Per-section slack for the zoom fit only, absorbing the font/bounds rounding
# described where it is used.
$script:SecHeadSlack = 6

function Set-SectionHeadBox {
    <#
        Size a section heading to its own text and return that height.

        MEASURED WITH 'Xg', NOT WITH THE SECTION NAME. Every heading then gets
        the same box whether or not its name happens to have a descender, so
        the headings and the cards below them line up across sections. 'Xg'
        spans cap-height to descender, which is the tallest a single line of
        this face can be.

        MEASURED IN THE FACE THE THEME WILL APPLY, NOT THE ONE THE LABEL HAS.
        That is what lets this be a single pass. The build loop runs before
        Set-ConsoleTheme has re-faced anything and before the accent loop has
        made the headings bold, so measuring the label's current font would
        size every box for a face it is about to stop using -- and bold Segoe
        is taller than the default. Asking the theme what it is going to use
        gets the final answer up front.

        Re-stacking after the theme instead was tried and is the wrong fix: see
        the note where the accent loop ends.

        The factor only scales the padding. The font SIZE is already whatever
        the zoom left it, so the measurement needs no adjusting -- scaling it
        again is what produced a 20px box for 47px of text.
    #>
    param($Label, [double]$Factor = 1.0)
    $h = [int][Math]::Round(20 * $Factor)
    if ($Label) {
        try {
            # Get-ThemeUiFont is cached by Get-ThemeFont, so this does not leak
            # a GDI handle per call however often a re-stack runs.
            $mf = $Label.Font
            try { $mf = Get-ThemeUiFont -Size $Label.Font.Size -Style ([System.Drawing.FontStyle]::Bold) } catch { }
            if (-not $mf) { $mf = $Label.Font }
            $m = [System.Windows.Forms.TextRenderer]::MeasureText('Xg', $mf).Height
            if ($m -gt 0) { $h = $m + [int][Math]::Round($script:SecHeadPad * $Factor) }
        } catch { }
        # AutoSize stays OFF: the label spans the full panel width because the
        # whole heading strip is the click target that folds the section.
        try { $Label.Height = $h } catch { }
    }
    return $h
}

function Get-SectionHeadOffset {
    <# Heading top -> panel top, i.e. the box plus the gap under it. #>
    param($Label, [double]$Factor = 1.0)
    return ((Set-SectionHeadBox -Label $Label -Factor $Factor) +
            [int][Math]::Round($script:SecHeadGap * $Factor))
}

function Expand-LabelToFont {
    <#
        Grow a fixed-size label so its text fits. Never shrinks it.

        THE VERSION FOOTER IS THE SAME BUG AS THE SECTION HEADINGS, in a
        different control. It is built as an 18px box because AutoSize is not an
        option -- it needs a FIXED width so AutoEllipsis can truncate a long
        manifest line -- and 18px is right for the 8pt face it is created with.
        The theme then re-faces it (Set-ThemeFontOn reaches even NoTheme
        controls, by design) and the zoom scales its font by $fFontTop while
        scaling its bounds by $fTop. Those two factors are not equal, so the
        text outgrew the box: measured at 10.5pt needing 28px in a 23px box,
        which is the sliced version line along the bottom of the window.

        GROW-ONLY IS THE WHOLE SAFETY ARGUMENT. Every caller here is a placement
        pass that runs repeatedly, and a height that could also shrink would
        fight the zoom's own bounds scaling on alternate passes. Widening is
        never attempted for the same reason the label is fixed-width to start
        with.
    #>
    param($Label)
    if (-not $Label) { return }
    try {
        # MeasureText already includes ascent, descent and leading, so the
        # measured height needs no padding added to it.
        $m = [System.Windows.Forms.TextRenderer]::MeasureText('Xg', $Label.Font).Height
        if ($m -gt 0 -and $Label.Height -lt $m) { $Label.Height = $m }
    } catch { }
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
    $lbl.Cursor = [System.Windows.Forms.Cursors]::Hand
    $form.Controls.Add($lbl)

    # CLICK THE HEADING TO FOLD THE SECTION. The name rides on the label
    # itself because the handler runs from the message loop, where the loop
    # variable is long gone. Folding is a per-operator preference, kept in
    # the same file as pinned and recent -- nothing shipped in Config changes.
    $lbl | Add-Member -NotePropertyName 'SectionName' -NotePropertyValue "$($sec.Name)" -Force
    $lbl.Add_Click({
        param($s, $e)
        try {
            if ($script:Arrange.On) { return }
            $n = "$($s.SectionName)"
            $now = -not [bool]$script:CollapsedSections[$n]
            $script:CollapsedSections[$n] = $now
            Set-ConsoleSectionCollapsed -Name $n -Collapsed $now
            Update-ConsoleSearch -Text $txtFind.Text
        } catch { }
    })

    # Measured rather than the old literal 22. Set-SectionHeadBox measures the
    # face the THEME will use, so this is already the final answer even though
    # the theme has not run yet.
    $headOff = Get-SectionHeadOffset -Label $lbl
    # This loop runs before any zoom, so what it measures IS the design-space
    # offset -- which is exactly what Get-ConsoleZoomFit needs to reserve per
    # section. Publishing it here keeps the fit and the stacking in step
    # automatically instead of by keeping two literals equal by hand.
    $script:SecHeadRes = $headOff

    $flow = New-Object System.Windows.Forms.FlowLayoutPanel
    $flow.Location = New-Object System.Drawing.Point(12, ($y + $headOff))
    $flow.Size = New-Object System.Drawing.Size($script:PanelW, $panelH)

    # Half the slack goes on the left so the spare width reads as a margin on
    # both edges rather than a gap on one. A FlowLayoutPanel lays out from its
    # left padding edge, so this shifts the grid right by half and leaves the
    # other half past the last column.
    # Padding and the column limit are both set by Set-SectionPanelGrid, from
    # the panel's real client width -- see it for why the count has to be
    # forced rather than left to the flow to work out.
    $flow.AutoScroll = $false
    $flow.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $flow.Tag = "$($sec.Name)"
    $form.Controls.Add($flow)

    # Drop target for Arrange mode. The button is moved LIVE during DragOver
    # so the grid shows where it will land; DragDrop only persists.
    $flow.AllowDrop = $true
    $flow.Add_DragEnter({
        param($s, $e)
        if ($script:Arrange.On -and $e.Data.GetDataPresent([System.Windows.Forms.Button])) { $e.Effect = [System.Windows.Forms.DragDropEffects]::Move }
        else { $e.Effect = [System.Windows.Forms.DragDropEffects]::None }
    })
    $flow.Add_DragOver({
        param($s, $e)
        if (-not $script:Arrange.On) { $e.Effect = [System.Windows.Forms.DragDropEffects]::None; return }
        $btn = $e.Data.GetData([System.Windows.Forms.Button])
        if (-not $btn) { $e.Effect = [System.Windows.Forms.DragDropEffects]::None; return }
        if ($btn.Parent -ne $s) {
            # Crossing into a FULL section GROWS it by a row. This used to
            # refuse the drop, silently -- the button simply would not go, and
            # the reason (Rows in Sections.json) was nowhere on screen. Reported
            # 2026-09-12 as "unable to move icons to Mailbox and Access". The
            # extra row is written to Sections.json by Save Arrangement, the
            # same save that records where the button went, so the two files
            # cannot disagree at the next launch. Never shrunk here: rows the
            # operator declared are theirs.
            $cap = [int]$script:SectionRows["$($s.Tag)"] * $script:PerRow
            $n = @($s.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] -and $_.Visible }).Count
            if ($n -ge $cap) {
                $script:SectionRows["$($s.Tag)"] = [int]$script:SectionRows["$($s.Tag)"] + 1
                $script:SectionRowsDirty = $true
            }
            $btn.Parent.Controls.Remove($btn)
            $s.Controls.Add($btn)
        }
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Move
        $pt = $s.PointToClient((New-Object System.Drawing.Point($e.X, $e.Y)))
        $over = $s.GetChildAtPoint($pt)
        if ($over -and $over -ne $btn -and $over -is [System.Windows.Forms.Button]) {
            $s.Controls.SetChildIndex($btn, $s.Controls.GetChildIndex($over))
        }
    })
    $flow.Add_DragDrop({
        param($s, $e)
        if (-not $script:Arrange.On) { return }
        Save-ToolArrangement
    })

    Set-SectionPanelGrid -Panel $flow
    $script:SectionPanels["$($sec.Name)"] = $flow
    $script:SectionLabels["$($sec.Name)"] = $lbl
    $script:SectionAccents["$($sec.Name)"] = "$($sec.Accent)"
    $y += $headOff + $panelH + $script:SecAfter
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

# NO OPERATOR NAME ON THE VERSION LINE.
#
# It was added here (backlog #187) before the bottom status strip existed. The
# strip now ends with DOMAIN\user permanently, so printing it again three
# lines above was the same fact twice -- and the second copy is the one that
# pushed the version line's useful part toward the ellipsis.
#
# Kept as an empty string rather than deleted outright: three format strings
# below append it, and an empty append is a smaller, safer edit than three
# more string changes for a line that may want a suffix again one day.
$script:OperatorTag = ''

try {
    $vv = Test-ToolkitIntegrity
    if (-not $vv.HasManifest) {
        $lblVersion.Text      = 'Admin Toolkit -- unversioned copy (no Version.json)' + $script:OperatorTag
        $lblVersion.ForeColor = [System.Drawing.Color]::DimGray
    } elseif ($vv.Modified.Count -eq 0 -and $vv.Missing.Count -eq 0 -and $vv.Extra.Count -eq 0) {
        $lblVersion.Text      = "Admin Toolkit v$($vv.Version)  --  $($vv.Matched) files verified  --  released $($vv.Released)$script:OperatorTag"
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

        $lblVersion.Text      = "Admin Toolkit v$($vv.Version)  --  $detail  --  run Tester for the list$script:OperatorTag"
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
    # with 'the argument "" does not belong to the set "DirA,DirB"' when the
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

function Test-ConsoleThemeChanged {
    <#
        After a tool returns: did the active theme change while it ran
        (Theme Manager is a tool)? The console cannot re-skin itself in place
        -- Set-ConsoleTheme is a one-shot pass that installs paint handlers
        and scaling, and running it twice stacks them -- so the honest offer
        is a restart in a fresh process, which is the only thing that
        guarantees every cache is gone.
    #>
    try {
        $now = Get-ConsoleTheme -Refresh
        $was = $(if ($script:Theme) { "$($script:Theme.Name)" } else { '' })
        $new = $(if ($now) { "$($now.Name)" } else { '' })
        if ($new -eq $was) { return }
        $a = Show-ToolDialog(
            "The active theme is now '$(if ($new) { $new } else { 'Default' })' (it was '$(if ($was) { $was } else { 'Default' })').`r`n`r`nThe console applies its theme when it starts. Restart it now?",
            'Theme changed', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($a -ne [System.Windows.Forms.DialogResult]::Yes) {
            $lblResult.Text = "Theme '$new' applies on the next start."
            return
        }
        # Same mechanism the Connect scripts use: close, and the trampoline
        # at the top of this file re-runs the script in THIS process, so the
        # Exchange sessions survive. The theme cache was already refreshed
        # above, and the next pass resolves it fresh anyway.
        $global:ToolkitConsoleRelaunch = $true
        $form.Close()
    } catch { }
}

function Start-ToolScript {
    param($Tool, [string]$SmtpOverride)

    $scriptPath = Join-Path $PSScriptRoot $Tool.Script
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        $null = Show-ToolDialog(
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

    # RECORDED HERE BECAUSE THIS IS THE ONE PLACE EVERY LAUNCH GOES THROUGH.
    # The button grid, the list view, the palette, the catalogue and the group
    # chooser all end up in this function. Recording at each call site instead
    # would mean the next view someone adds silently stops counting, and the
    # Recent view would quietly start lying.
    #
    # Recorded BEFORE the launch, not after: a tool that throws is still a tool
    # you reached for, and after a crash "what was I just running" is exactly
    # the question being asked.
    try { Add-ConsoleToolRun -Name "$($Tool.Name)" } catch { }

    try {
        if ($needsOwnProcess) {
            Start-ToolScriptInNewProcess -Tool $Tool -ScriptPath $scriptPath
            Set-ToolStatus -Form $form -Done -Text "'$($Tool.Name)' opened in its own PowerShell window -- it connects to on-premises Exchange first, so it takes a moment to appear."
            return
        }
        # @toolParams, not @(...). Splatting needs a VARIABLE holding the
        # dictionary; @(...) is an array subexpression and would hand the whole
        # thing over as one positional argument.
        $toolParams = Get-ToolParameterSet -ScriptPath $scriptPath
        # Look Up hands Unified Lookup whatever was TYPED, resolved or not --
        # that tool searches by fragment and does its own resolving. Only a
        # script that declares SmtpAddress may be called this way.
        if ($SmtpOverride) { $toolParams['SmtpAddress'] = $SmtpOverride }
        # The tool's window is modal, so this stays "Running ..." until it
        # closes -- which is the truth, and the elapsed time on return is how
        # long the operator had it open.
        $before = $lblResult.Text
        Set-ToolStatus -Form $form -Busy -Text "Running $($Tool.Name) ..."
        try { & $scriptPath @toolParams }
        finally { Set-ToolStatus -Form $form -Done -Text $before }
        Test-ConsoleThemeChanged
        # A tool may have started or cleared the PIM countdown. The header
        # timer would catch up within a minute; a click deserves better.
        try { Update-PimLabel -Lbl $lblPim -Tip $toolTip } catch { }
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

function Add-ArrangeDragHandlers {
    <#
        MouseDown remembers, MouseMove past the system drag threshold starts
        the drag. The threshold is what keeps a click a click: a button that
        started DoDragDrop on MouseDown would never raise Click again.
    #>
    param([Parameter(Mandatory)]$Button)
    $Button.Add_MouseDown({
        param($s, $e)
        if ($script:Arrange.On -and $e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
            $script:Arrange.Drag = @{ Btn = $s; X = $e.X; Y = $e.Y }
        }
    })
    $Button.Add_MouseMove({
        param($s, $e)
        $d = $script:Arrange.Drag
        if (-not $d -or $d.Btn -ne $s -or $e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
        $th = [System.Windows.Forms.SystemInformation]::DragSize
        if ([Math]::Abs($e.X - $d.X) -ge $th.Width -or [Math]::Abs($e.Y - $d.Y) -ge $th.Height) {
            $script:Arrange.Drag = $null
            [void]$s.DoDragDrop($s, [System.Windows.Forms.DragDropEffects]::Move)
        }
    })
    $Button.Add_MouseUp({ $script:Arrange.Drag = $null })
}

function Save-ToolArrangement {
    <#
        Panel order -> Order and Section on every tool -> Tools.json.

        A group button stands in for its members, which are hidden buttons
        that stay wherever they were created; so the walk emits the group's
        members at the group button's position and skips them when it meets
        them on their own. Order is one running number across all sections,
        which is what Get-ToolDefinitions sorts by.
    #>
    $n = 0
    foreach ($name in $script:SectionNames) {
        $panel = $script:SectionPanels[$name]
        if (-not $panel) { continue }
        foreach ($c in $panel.Controls) {
            if (-not ($c -is [System.Windows.Forms.Button])) { continue }
            $t = $c.Tag
            if ($null -eq $t) { continue }
            if ($t.PSObject.Properties['IsToolGroup'] -and $t.IsToolGroup) {
                foreach ($m in @($script:ToolGroups["$($t.GroupName)"])) {
                    Add-Member -InputObject $m -NotePropertyName 'Order' -NotePropertyValue $n -Force
                    Add-Member -InputObject $m -NotePropertyName 'Section' -NotePropertyValue $name -Force
                    $n++
                }
                continue
            }
            if (Get-ToolGroupName -Tool $t) { continue }   # emitted with its group
            Add-Member -InputObject $t -NotePropertyName 'Order' -NotePropertyValue $n -Force
            Add-Member -InputObject $t -NotePropertyName 'Section' -NotePropertyValue $name -Force
            $n++
        }
    }
    try {
        $tools = @($script:AllTools | Sort-Object { [int]$_.Order })
        $json = $tools | ConvertTo-Json -Depth 4
        if ($tools.Count -le 1) { $json = "[`r`n$json`r`n]" }
        Set-Content -LiteralPath $ToolsConfigPath -Value $json -Encoding UTF8
        $script:AllTools = $tools
        $lblResult.Text = "Layout saved to Tools.json ($($tools.Count) tools)."
        Write-CompletedLog -Action 'ArrangeConsole' -Target 'Tools.json' -Result 'Success' -Detail "$($tools.Count) tools re-ordered by drag"
    } catch {
        $lblResult.Text = "Layout NOT saved: $($_.Exception.Message)"
    }
    # A drag grew a section (see the DragOver handler): persist the new row
    # count and reflow so the extra row is visible now, not at next launch.
    # Sections.json is read and written whole so ButtonsPerRow, Accent and
    # anything else in it survive -- only Rows changes, only where it grew.
    if ($script:SectionRowsDirty) {
        $script:SectionRowsDirty = $false
        try {
            $raw = Get-Content -LiteralPath $script:SectionsConfigPath -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $changed = $false
            foreach ($sec in @($raw.Sections)) {
                $nm = "$($sec.Name)"
                if (-not $script:SectionRows.ContainsKey($nm)) { continue }
                $want = [int]$script:SectionRows[$nm]
                if ($want -gt [int]$sec.Rows) { $sec.Rows = $want; $changed = $true }
            }
            if ($changed) {
                $raw | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $script:SectionsConfigPath -Encoding UTF8 -ErrorAction Stop
                $lblResult.Text += ' Sections.json: a section grew by a row.'
            }
        } catch {
            $lblResult.Text += " (Sections.json not updated: $($_.Exception.Message))"
        }
        try { Update-ConsoleSearch -Text $txtFind.Text } catch { }
    }
    if ($script:Zoom -and $script:Zoom.Applied -gt 1.0) { Update-ConsoleZoom }
}

function Add-ToolButtonsToPanel {
    param($Panel, $Tools, [string]$EmptyText)
    # THE CHROME IS ALSO APPLIED BY Set-ConsoleTheme, which reaches every form
    # in the toolkit. It is repeated here because these buttons are REBUILT
    # after theming -- Build-ToolButtons runs again when the arrangement is
    # saved and when the tool list is reloaded -- and a button created after
    # the theme pass would otherwise be the only flat one on the form.
    # Add-RaisedButtonChrome is idempotent, so the overlap costs nothing.
    $raised = Test-RaisedButtonStyle
    # -Fallback IS THE DIFFERENCE BETWEEN HERE AND THE THEME PASS. Set-ConsoleTheme
    # only gives a button an icon when its caption confidently says what it does,
    # which is right for dialogs -- 'Browse...' should stay clean. A grid is the
    # opposite: one tile without a picture among thirty with them reads as a tool
    # that failed to load, so every tile gets one and the unmatched ones get a
    # neutral glyph rather than a hole.
    $icons  = Test-ConsoleIconStyle
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
            $gb.Add_Click({ if ($script:Arrange.On) { return }; Show-ToolGroupChooser -GroupName $this.Tag.GroupName })
            if ($raised) { Add-RaisedButtonChrome -Button $gb }
            Add-ArrangeDragHandlers -Button $gb
            $Panel.Controls.Add($gb)
            # AFTER the button joins the panel, and that is load-bearing. A
            # control inherits its Font from its PARENT, so until this line
            # $gb.Font is Control.DefaultFont -- and the glyph sizes itself and
            # measures the caption from that font. Adding the icon first would
            # price it in the wrong face on every button on the console.
            if ($icons) { Add-ButtonGlyph -Button $gb -Fallback }
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
        if ($raised) { Add-RaisedButtonChrome -Button $btn }
        Add-ArrangeDragHandlers -Button $btn
        $btn.Add_Click({
            if ($script:Arrange.On) { return }
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
        # AFTER the panel adopts it, for the font (see the group button above),
        # and after .Tag is set, because Tools.json can name an icon outright
        # and Add-ButtonGlyph reads it off the tool object sitting in there.
        if ($icons) { Add-ButtonGlyph -Button $btn -Fallback }
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

        # A FOLDED SECTION IS A HEADING AND NOTHING ELSE. Searching unfolds it
        # for the duration of the search: a match hidden behind a fold would
        # read as "not found", and the operator typed the name of the thing.
        $folded = (-not $filtering) -and [bool]$script:CollapsedSections[$name]
        # Sized from the font the label has RIGHT NOW. This is the pass that
        # runs after the theme and after the accent pass made the headings
        # bold, so it is the first point at which the real text height is
        # knowable -- see Set-SectionHeadBox for what the old 20px box did.
        $headOff = Get-SectionHeadOffset -Label $lbl
        if ($lbl) {
            $lbl.Text = $(if ($folded) { "$name  [+]" } else { "$name" })
            $lbl.Location = New-Object System.Drawing.Point(12, $y)
        }
        if ($folded) {
            $panel.Visible = $false
            $y += $headOff + 6
            continue
        }
        $panel.Location = New-Object System.Drawing.Point(12, ($y + $headOff))
        $panel.Height   = $panelH
        $y += $headOff + $panelH + $script:SecAfter
    }

    if ($lblVersion) {
        $lblVersion.Location = New-Object System.Drawing.Point(12, ($y + 4))
        Expand-LabelToFont -Label $lblVersion
        $y += 26
    }

    # ONCE THE OPERATOR HAS RESIZED, THE WINDOW SIZE IS THEIRS AND NOT OURS.
    #
    # This used to set ClientSize unconditionally, which is right while the
    # console is still at its design size -- filtering to four tools should
    # shrink the window rather than leave a field of grey. It is wrong the
    # moment the window has been zoomed: typing in the Find box would snap the
    # window back toward design size and undo the zoom, which reads as the
    # search box randomly resizing the console.
    #
    # So when zoomed, re-run the zoom instead. It reflows the now-smaller set of
    # buttons into the window the operator chose and leaves that window alone.
    if ($script:Zoom -and $script:Zoom.Applied -gt 1.0) {
        Update-ConsoleZoom
    } else {
        $form.ClientSize = New-Object System.Drawing.Size(($script:PanelW + 24), ($y + 8))
        $work = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea
        if ($form.Height -gt $work.Height) { $form.Height = $work.Height }
        if ($form.Width  -gt $work.Width)  { $form.Width  = $work.Width }
    }

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
            if ($flow) { $flow.Width = $script:PanelW; Set-SectionPanelGrid -Panel $flow }
        }
        # Everything else that was sized from PanelW has to follow, or the
        # output box stops lining up with the button grid -- which is the same
        # lopsided look this was written to remove, just moved.
        #
        # The lookup box is NOT in this list any more: it tracks the find box,
        # and Set-ConsoleFindRowLayout below re-derives both from the panel's
        # new right edge.
        $lblResult.Width  = $script:PanelW
        if (Get-Command Set-ConsoleStatusLayout -ErrorAction SilentlyContinue) { Set-ConsoleStatusLayout }
        $txtOut.Width     = $script:PanelW
        $lblVersion.Width = $script:PanelW
        if (Get-Command Set-ConsoleFindRowLayout -ErrorAction SilentlyContinue) { Set-ConsoleFindRowLayout }

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
                $Lbl.Text = $Prefix
                $Lbl.ForeColor = [System.Drawing.Color]::DimGray
                return
            }
            if ($State.PSObject.Properties['NotConfigured'] -and $State.NotConfigured) {
                # Not greyed -- gone. A hidden label gets a zero-width pill
                # (Set-ConsoleStatusLayout), so the row closes up rather than
                # leaving a gap where a capability this tenant lacks used to be.
                $Lbl.Visible = $false
                return
            }
            $Lbl.Visible = $true
            if ($State.Available) {
                $Lbl.Text = $Prefix
                $Lbl.ForeColor = $green
            } else {
                $Lbl.Text = $Prefix
                $Lbl.ForeColor = $red
            }
        }

        & $setState $lblExo $exo 'ExO'
        $toolTip.SetToolTip($lblExo, "$($exo.Detail)")

        if (-not $spo) {
            $lblSpo.Text = 'SPO'
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
        # THIS IS WHERE THE STATUS STRIP KEPT BREAKING, AND IT IS NOT THE
        # STRIP'S FAULT.
        #
        # These two lines used to re-run the HEADER row layout over lblExo,
        # lblSpo, lblTenant, lblEms, lblAd and lblPim -- the six labels that
        # New-ConsolePill RE-PARENTED into the pills at the bottom of the
        # window. That function placed controls from x = 12 at the Y
        # it is given, which inside an 18px-tall pill panel puts the caption
        # off the visible area entirely. The dot stayed, the words went, and
        # the pills read as a row of coloured squares.
        #
        # WHY IT LOOKED LIKE AN INTERMITTENT BUG: this function runs after
        # every tool returns and on every connection probe, never at startup.
        # So a clean launch was always right and the first piece of real work
        # always broke it -- and any later event that happened to re-run
        # Set-ConsoleStatusLayout put it back, which is why it also appeared to
        # fix itself.
        #
        # Four separate repairs to the strip's own layout could not have fixed
        # this, because the strip was being overwritten from outside by a
        # function that predates it and does not know it exists.
        #
        try { Set-ConsoleStatusLayout } catch { }

        # THE "Use Connect below." PROMPT USED TO BE SHOWN OR HIDDEN HERE, from
        # whether any relevant service was unavailable. Removed 2026-09-13 as
        # redundant against the status pills; the probe result it read is still
        # computed and still drives the button greying below.

        # Same probe, same answer: repaint the tool buttons from the map that
        # just produced these labels.
        Update-CapabilityState
    } catch {
        $lblExo.Text = 'ExO'
        $lblExo.ForeColor = [System.Drawing.Color]::DimGray
        $lblSpo.Text = 'SPO'
        $lblSpo.ForeColor = [System.Drawing.Color]::DimGray
        $lblEms.Text = 'EMS'
        $lblEms.ForeColor = [System.Drawing.Color]::DimGray
        $lblAd.Text = 'AD'
        $lblAd.ForeColor = [System.Drawing.Color]::DimGray
        # The prompt was forced visible here when the probe threw. Nothing to
        # force now -- the four labels above are already greyed, which says the
        # same thing without a sentence.
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
    # NOT WHILE ARRANGING. Every tool button is deliberately enabled in
    # Arrange mode so that a greyed-out tool can still be dragged; a lookup or
    # a background capability refresh landing mid-drag would grey it again
    # under the operator's mouse and the drag would die on the spot.
    # Set-ArrangeMode calls this itself when Arrange goes off, which is where
    # the real state comes back.
    if ($script:Arrange -and $script:Arrange.On) { return }
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

    # LOOK UP IS UNIFIED LOOKUP. Whatever was typed -- "glenn", an alias, a
    # full address -- goes to Unified Lookup as typed, which searches every
    # directory by fragment. That is the lookup. It runs FIRST and always.
    #
    # The console's own resolution (which greys the grid and hands every other
    # tool its address) runs afterwards, and ONLY for a full address. It used
    # to run on everything, and on "glenn" it appended the mail domain, went
    # looking for glenn@..., and announced it could not find him -- while
    # Unified Lookup had already found six of him. Reported 2026-09-12.
    $typed = "$($txtInput.Text)".Trim()
    if ($typed) {
        try { Open-UnifiedLookupFor -Text $typed } catch { }
        if ($typed -notmatch '@') {
            $script:LastSmtp = $null
            $script:LastADServer = $null
            $script:LastEnvironment = $null
            $script:LastType = $null
            $lblResult.Text = "Searched '$typed' in Unified Lookup. To point the tools below at one object, type its full address and Look Up."
            Update-ToolButtonState -DetectedType $null
            Set-ToolStatus -Form $form -Done -Text $lblResult.Text
            $btnRun.Enabled = $true
            return
        }
    }

    # Full address: resolve quietly. Reasons for a miss go to the status bar,
    # not a dialog -- Unified Lookup is already open with the real answer.
    $script:QuietOutput = $true
    Set-ToolStatus -Form $form -Busy -Text 'Looking up ...'
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
        Set-ToolStatus -Form $form -Text "Looking up $smtp ..."

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
            Set-ToolStatus -Form $form -Text "Looking up $smtp in the on-premises directory ..."
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

        Set-ToolStatus -Form $form -Text "Resolving $smtp in Exchange Online ..."
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
        # -Done keeps whatever the branches wrote in $lblResult (the
        # "Found: ..." line) and stamps the elapsed time beside it.
        Set-ToolStatus -Form $form -Done -Text $lblResult.Text
        $btnRun.Enabled = $true
        $script:QuietOutput = $false
    }
})

function Open-UnifiedLookupFor {
    <#
        Launch Unified Lookup with the typed text, through Start-ToolScript so
        the launch is recorded and the status bar narrates it like any other.
        Taken from Tools.json when it is there, so a renamed button still
        works; built by hand when it is not, so removing the button does not
        remove the behaviour.
    #>
    param([string]$Text)
    $q = "$Text".Trim()
    if (-not $q) { return }
    $tool = $null
    foreach ($t in @($script:AllTools)) {
        if ("$($t.Script)" -match '(?i)Unified-Lookup\.ps1$') { $tool = $t; break }
    }
    if (-not $tool) {
        if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'Unified-Lookup.ps1'))) { return }
        $tool = [PSCustomObject]@{ Name = 'Unified Lookup'; Script = '.\Unified-Lookup.ps1'; Section = ''; ObjectTypes = @() }
    }
    Start-ToolScript -Tool $tool -SmtpOverride $q
    Update-ConnectionButtonColors
}
# Anything written into the hidden output box becomes a dialog. Blanking it
# (the two Clear paths) says nothing. Information, not Error: most of what
# lands here is "not found" with the directory's reasons, which is an answer.
$txtOut.Add_TextChanged({
    try {
        $msg = "$($txtOut.Text)".Trim()
        if (-not $msg) { return }
        # During Look Up the detail is a footnote, not news: first line to
        # the status bar and no dialog.
        if ($script:QuietOutput) {
            $first = ($msg -split "`r?`n")[0]
            Set-ToolStatus -Form $form -Text $first
            return
        }
        $null = Show-ToolDialog($msg, 'Admin Console', 'OK', 'Information')
    } catch { }
})

# The Clear and Close handlers lived here. Both buttons were removed on
# 2026-09-13 -- see the note where they used to be created, which records what
# Clear did beyond emptying the box.

# Filter as you type. SuspendLayout around it because the handler moves every
# panel and the version label on each keystroke, and WinForms would otherwise
# repaint the whole form per character.
$txtFind.Add_TextChanged({
    $form.SuspendLayout()
    try { Update-ConsoleSearch -Text $txtFind.Text }
    finally { $form.ResumeLayout() }
})

# Escape clears the filter.
#
# IT USED TO ALSO CLOSE THE CONSOLE when there was no filter, because
# $form.CancelButton was the Close button. That button was removed on
# 2026-09-13 and CancelButton with it, so Escape is now only ever "abandon the
# search" and the window closes from its own title bar. The guard below is kept
# anyway: it costs nothing and it leaves Escape available to a dialog opened
# over the console.
# CTRL+F FROM ANYWHERE, ESC FROM ANYWHERE. KeyPreview lets the form see the
# key before the focused control does. Ctrl+F is claimed outright; Esc is
# claimed only while there is a filter to clear, so a dialog or a button that
# wants Esc still gets it the rest of the time.
$form.KeyPreview = $true
$form.Add_KeyDown({
    param($s, $e)
    try {
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::F) {
            $txtFind.Focus(); $txtFind.SelectAll()
            $e.Handled = $true; $e.SuppressKeyPress = $true
            return
        }
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape -and $txtFind.Text) {
            $txtFind.Clear()
            $e.Handled = $true; $e.SuppressKeyPress = $true
        }
    } catch { }
})

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

function Set-ConsoleFindRowLayout {
    <#
        PLACE THE RIGHT-HAND END OF THE FIND ROW FROM WHAT THE CONTROLS
        ACTUALLY MEASURE, NOT FROM WHAT THEY WERE CREATED AS.

        THIS IS THE FIX FOR "Arrange is partially covered", and the reason it
        survived three attempts is that the arithmetic was never wrong -- it
        was STALE. The layout block near the top of this script computes X
        positions from literal widths (74, 90, 50). Set-ConsoleTheme then runs,
        applies the typeface and the DPI scaling, and BUTTONS CHANGE WIDTH.
        Every X derived from the old widths is then out by however much they
        grew, and the control to the left ends up underneath its neighbour.

        Reading .Width at the moment of placement cannot be stale. Nothing here
        assumes a size; it asks.

        CALL IT AFTER ANYTHING THAT CAN RESIZE THESE CONTROLS: the initial
        theme pass, and the end of every zoom pass. Last writer wins, which is
        why it does not matter that the zoom's own right-anchor logic has
        already had a go at them.
    #>
    # ---- EVERY BOX IN THIS HEADER IS SIZED FROM ITS OWN FONT ---------------
    #
    # THIS IS THE FIX FOR THE 200 % DISPLAY, and -ZoomDiag is what finally
    # showed it. From that log, on a 3432x1688 (192 dpi) session:
    #
    #     lblUnified   box 254x22   text 248x47   ->  98% / 214%   CLIPPED
    #     btnRun       box 108x34   text 150x47   -> 139% / 138%   CLIPPED
    #     chkArrange   box  89x34   text 144x47   -> 162% / 138%   CLIPPED
    #     lblResult   box 3348x24   text 330x47   ->  10% / 196%   CLIPPED
    #     txtInput    box 1896x48   (empty)       ->             ok
    #
    # THE LAST LINE IS THE WHOLE ANSWER. A single-line TextBox derives its own
    # height from its font, so the two input boxes were always right. Labels and
    # Buttons do not, so every one of them kept a height written as a 96 dpi
    # literal -- 18, 24, 34 -- while the theme set the font in POINTS and 192
    # dpi rendered it 47px tall. The zoom scaled the boxes by its 1.2 and the
    # text by the same 1.2, so the mismatch survived scaling untouched: it was
    # never a factor problem.
    #
    # So the labels and buttons now do what the text boxes already did: ask the
    # font. Nothing here is a design constant except the gaps between things.
    #
    # WHY IT LOOKED LIKE THREE OTHER BUGS FIRST. Display scaling, the Windows
    # text-size slider and a per-app DPI override were each argued from
    # screenshots and each wrong, and the machine reported "96 dpi / 100 %"
    # because the diagnostic that asked was DPI-unaware and Windows lied to it.
    # Measure inside the process, in the font the control is wearing.
    $gap = 8
    $pad = 6

    function Measure-HeaderText {
        param($Control)
        if (-not $Control -or $Control.IsDisposed) { return $null }
        # HEIGHT USES A 'Wg' FLOOR, WIDTH DOES NOT. An empty caption still has
        # to reserve a line -- the result line and the match count are both
        # blank until something happens, and a label collapsed to zero height
        # would take its row's height with it. But measuring the WIDTH of a
        # stand-in makes an empty label demand room for text it does not have,
        # which is how the count label ended up 114 % of itself.
        $t  = "$($Control.Text)"
        $w  = 0
        if ($t) { $w = [System.Windows.Forms.TextRenderer]::MeasureText($t, $Control.Font).Width }
        $th = 0
        if ($t) { $th = [System.Windows.Forms.TextRenderer]::MeasureText($t, $Control.Font).Height }
        $lh = [System.Windows.Forms.TextRenderer]::MeasureText('Wg', $Control.Font).Height
        return [pscustomobject]@{ W = $w; H = [Math]::Max($th, $lh) }
    }

    # Buttons first: their height sets the row height, and their width has to
    # hold their caption. GROW-ONLY on width so a short caption does not shrink
    # a button below the size it was designed at.
    $btnH = 0
    foreach ($b in @($btnRun, $btnFindClear, $btnView, $chkArrange)) {
        $m = Measure-HeaderText -Control $b
        if (-not $m) { continue }
        $want = $m.W + ($pad * 3)
        if ($want -gt $b.Width) { $b.Width = $want }
        $h = $m.H + ($pad * 2)
        if ($h -gt $btnH) { $btnH = $h }
    }
    if ($btnH -gt 0) {
        foreach ($b in @($btnRun, $btnFindClear, $btnView, $chkArrange)) {
            if ($b -and -not $b.IsDisposed -and $b.Height -lt $btnH) { $b.Height = $btnH }
        }
    }

    # Labels and the result line: height from the font, always -- these are the
    # ones that were 214 % over.
    foreach ($l in @($lblUnified, $lblFind, $lblFindCount, $lblResult)) {
        $m = Measure-HeaderText -Control $l
        if (-not $m) { continue }
        $l.Height = $m.H + 2
    }

    # THE MATCH COUNT NEEDS ITS WIDTH TOO. It carries '24 of 100' at 50 design
    # px, which is the one label that is neither full-width nor sized from the
    # shared label column -- and it is right-anchored, so growing it moves it
    # leftwards without pushing anything. Grow-only.
    if ($lblFindCount -and -not $lblFindCount.IsDisposed) {
        $m = Measure-HeaderText -Control $lblFindCount
        if ($m -and ($m.W + 4) -gt $lblFindCount.Width) { $lblFindCount.Width = $m.W + 4 }
    }

    $right = $form.ClientSize.Width - 12
    # 'Show all' IS NO LONGER IN THIS LIST. From 2026-09-13 the boxes are 70% of
    # the span and the trailing button follows the box end, so pinning it to the
    # right edge would leave it stranded half a window away from the box it
    # belongs to. The count, View and Arrange are still right-anchored.
    $row   = @($chkArrange, $btnView, $lblFindCount)
    foreach ($c in $row) {
        if (-not $c -or $c.IsDisposed) { continue }
        $right = $right - $c.Width
        $c.Location = New-Object System.Drawing.Point($right, $c.Top)
        $right = $right - $gap
    }

    # ---- both rows, as one column layout ---------------------------------
    #
    # THE LABEL COLUMN IS MEASURED FROM THE CAPTIONS, not from the 104 they
    # were created with. Set-ConsoleTheme applies the typeface and the DPI
    # scaling after that literal was written, so on a scaled display 104 is
    # either too narrow -- clipping 'Unified lookup' -- or too wide, which
    # pushes both boxes right for no reason. TextRenderer.MeasureText answers
    # for the font the label is actually wearing, at the moment of placement.
    $labelW = 0
    foreach ($l in @($lblUnified, $lblFind)) {
        if (-not $l -or $l.IsDisposed) { continue }
        $m = [System.Windows.Forms.TextRenderer]::MeasureText("$($l.Text)", $l.Font).Width
        if ($m -gt $labelW) { $labelW = $m }
    }
    if ($labelW -le 0) { $labelW = 104 }
    $labelW = $labelW + 6
    foreach ($l in @($lblUnified, $lblFind)) {
        if (-not $l -or $l.IsDisposed) { continue }
        $l.Location = New-Object System.Drawing.Point(12, $l.Top)
        $l.Width    = $labelW
    }

    $boxLeft = 12 + $labelW + $gap

    # 70% OF THE SPAN, not all of it -- see $script:ConsoleBoxWidthShare where
    # it is set. $right is the left edge of the count label less a gap, so the
    # room a row actually has is that, less the trailing button and its gap.
    $btnW = 90
    if ($btnFindClear -and -not $btnFindClear.IsDisposed) { $btnW = $btnFindClear.Width }
    $share = 0.7
    if ($script:ConsoleBoxWidthShare) { $share = [double]$script:ConsoleBoxWidthShare }

    $full = $right - $gap - $boxLeft - $btnW
    $boxW = [int]($full * $share)
    # A floor, for the same reason as before: the console is resizable and a box
    # squeezed to nothing is worse than one that asks for a wider window.
    if ($boxW -lt 120) { $boxW = 120 }
    # And a ceiling, so a narrow window cannot push the button under the count
    # label -- the floor above wins over the share, and this keeps that honest.
    if ($boxW -gt $full -and $full -ge 120) { $boxW = $full }

    foreach ($t in @($txtInput, $txtFind)) {
        if (-not $t -or $t.IsDisposed) { continue }
        $t.Width = $boxW
    }

    # ---- STACK THE ROWS FROM THE HEIGHTS THEY ACTUALLY HAVE ----------------
    #
    # The Y positions were literals too -- 68 and 98, 30px apart, which is a
    # 28px button and 2px of air. At 192 dpi the buttons are 59px and the rows
    # overlapped each other as well as clipping their own text.
    #
    # Measured, in order, each row below the one above it. The title is AutoSize
    # so its Bottom is font-driven and can be trusted as the starting point.
    $topY = 12
    if ($lblTitle -and -not $lblTitle.IsDisposed) { $topY = $lblTitle.Bottom + $gap }

    $row1 = @($lblUnified, $txtInput, $btnRun)          | Where-Object { $_ -and -not $_.IsDisposed }
    $row2 = @($lblFind, $txtFind, $btnFindClear, $lblFindCount, $btnView, $chkArrange) | Where-Object { $_ -and -not $_.IsDisposed }

    $h1 = 0; foreach ($c in $row1) { if ($c.Height -gt $h1) { $h1 = $c.Height } }
    $h2 = 0; foreach ($c in $row2) { if ($c.Height -gt $h2) { $h2 = $c.Height } }

    # Vertically centred inside its row, so a 22px label sits on the same
    # baseline as a 59px button instead of at the top of the gap.
    foreach ($c in $row1) { $c.Top = $topY + [int](($h1 - $c.Height) / 2) }
    $row2Y = $topY + $h1 + 4
    foreach ($c in $row2) { $c.Top = $row2Y + [int](($h2 - $c.Height) / 2) }

    # Both trailing buttons sit at the end of their own box, one above the other.
    $btnX = $boxLeft + $boxW + $gap
    foreach ($b in @($btnRun, $btnFindClear)) {
        if (-not $b -or $b.IsDisposed) { continue }
        $b.Left = $btnX
    }
    foreach ($t in @($txtInput, $txtFind)) {
        if (-not $t -or $t.IsDisposed) { continue }
        $t.Left = $boxLeft
    }

    $y = $row2Y + $h2 + $gap
    if ($lblResult -and -not $lblResult.IsDisposed) {
        $lblResult.Top = $y
        $y = $y + $lblResult.Height + 4
    }

    # PUBLISHED, because the grid has to start below whatever this came to.
    # Update-ConsoleZoom and Get-ConsoleZoomFit both read it: the header is no
    # longer a known 164px, it is however tall its own text made it, and on a
    # 192 dpi session that is closer to 300.
    $script:HeaderBottom = $y
}

function Set-ArrangeMode {
    param([bool]$On)
    $script:Arrange.On = $On
    $script:Arrange.Drag = $null
    $chkArrange.Text = $(if ($On) { 'Arrange ON' } else { 'Arrange' })
    # Accent while on, the theme's plain button otherwise; without a theme,
    # the system colours.
    if ($script:Theme) {
        $bk = ConvertTo-ThemeColor $(if ($On) { $script:Theme.AccentBack } else { $script:Theme.ButtonBack })
        $fg = ConvertTo-ThemeColor $(if ($On) { $script:Theme.AccentText } else { $script:Theme.ButtonText })
        if ($bk) { $chkArrange.BackColor = $bk }
        if ($fg) { $chkArrange.ForeColor = $fg }
    } else {
        $chkArrange.BackColor = $(if ($On) { [System.Drawing.SystemColors]::Highlight } else { [System.Drawing.SystemColors]::Control })
        $chkArrange.ForeColor = $(if ($On) { [System.Drawing.SystemColors]::HighlightText } else { [System.Drawing.SystemColors]::ControlText })
    }
    # A DISABLED BUTTON RAISES NO MOUSE EVENTS AT ALL, so a greyed-out tool
    # could not be dragged -- Arrange worked on whatever this operator happened
    # to be connected for and the rest of the grid was nailed down. Reported
    # 2026-09-12. Arranging is not running: it moves a button, it launches
    # nothing, and which directories are reachable this minute has no bearing
    # on where a button belongs.
    #
    # So EVERY tool button is enabled while Arrange is on, and the real state
    # is RECOMPUTED when it goes off rather than parked and restored. Parking
    # was the first attempt and it is the wrong shape: a lookup or a
    # connection finishing mid-arrange would make the parked answer stale, and
    # the restore would then put back something that was true five minutes
    # ago. Update-ToolButtonState already knows how to answer this from
    # scratch, so it answers it.
    #
    # Update-ToolButtonState stands down while Arrange is on -- see the guard
    # at the top of it -- so nothing re-greys a button mid-drag.
    #
    # Click is already refused while Arrange is on (the handlers return early),
    # so an enabled-but-unavailable button still cannot be launched from here.
    foreach ($n in $script:SectionNames) {
        $p = $script:SectionPanels[$n]
        if (-not $p) { continue }
        $p.Cursor = $(if ($script:Arrange.On) { [System.Windows.Forms.Cursors]::SizeAll } else { [System.Windows.Forms.Cursors]::Default })
        if (-not $script:Arrange.On) { continue }
        foreach ($c in @($p.Controls)) {
            if ($c -is [System.Windows.Forms.Button] -and $c.Tag) { $c.Enabled = $true }
        }
    }
    if (-not $script:Arrange.On) {
        # Back to the truth, from the current lookup and the current
        # capabilities. Runs AFTER Arrange.On is false or the guard would
        # turn it into a no-op.
        try { Update-ToolButtonState -DetectedType $script:LastType } catch { }
    }
    $lblResult.Text = $(if ($script:Arrange.On) {
        'ARRANGE: drag a tool onto another to reorder, or into another section to move it. Every tool can be moved, including ones that are greyed out. Buttons do not launch until you press Arrange again.'
    } else { 'Arrange off. Layout is in Tools.json.' })
}
$chkArrange.Add_Click({ Set-ArrangeMode -On (-not $script:Arrange.On) })

#region Views -----------------------------------------------------------------
# THE SAME TOOLS, DRAWN DIFFERENT WAYS.
#
# A VIEW IS A WINDOW, NOT A MODE. Every alternative view opens as its own
# maximised form over the console and closes back to it. The obvious-looking
# alternative -- hiding the button panels and re-laying the console out in
# place -- would have to cooperate with the zoom walk, AutoScroll, the docked
# status strip, the header band and Set-ConsoleInitialFit, all of which took
# real effort to get right and none of which needs to learn about views. A
# separate form cannot break any of them.
#
# EVERY VIEW READS ONE MODEL. Get-ConsoleToolModel is the single answer to
# "what tools are there, which can run, and when did I last use one". Four
# views each deriving that from $script:AllTools and the capability tables
# would eventually disagree about whether a tool is available, and they would
# disagree on the day it mattered rather than the day it was written.
#
# TO ADD A VIEW: write a Show-Console<Name>View function and add one entry to
# $script:ConsoleViews. Nothing else changes.

# ---- per-operator state -----------------------------------------------------
#
# WHY PER-OPERATOR AND NOT ONE FILE. The version footer already says this
# console runs on a jump box several people share. "Recently used" and
# "pinned" are the first things in this toolkit that describe a PERSON rather
# than the tenant, and merging three engineers' habits into one list produces
# something that is nobody's. Keyed on Get-LogOperator, which is the same
# identity the audit log writes, so the two cannot drift apart.
#
# A MISSING OR CORRUPT FILE IS NOT AN ERROR. It means "no history yet", which
# is exactly what a new operator has. Nothing here may throw: a view that
# cannot open because a preferences file has a stray comma would be a
# ridiculous way to lose the console.
function Get-ConsoleStatePath {
    $who = 'operator'
    try { $who = "$(Get-LogOperator)" } catch { }
    $who = ($who -replace '[^A-Za-z0-9._-]', '_')
    if (-not $who) { $who = 'operator' }
    return (Get-ConfigFilePath -Name "ConsoleState.$who.json")
}

function Get-ConsoleState {
    # Collapsed and Window were added later. An older file has neither, and
    # an absent key means "nothing folded, default position" -- the same as a
    # new operator, which is the only safe reading of a missing preference.
    $default = [PSCustomObject]@{ Pinned = @(); Runs = @(); Collapsed = @(); Window = $null }
    $p = Get-ConsoleStatePath
    if (-not (Test-Path -LiteralPath $p)) { return $default }
    try {
        $j = Get-Content -LiteralPath $p -Raw -ErrorAction Stop | ConvertFrom-Json
        $pin = @(); $runs = @(); $col = @(); $win = $null
        if ($j.PSObject.Properties['Pinned'])    { $pin  = @($j.Pinned | ForEach-Object { "$_" } | Where-Object { $_ }) }
        if ($j.PSObject.Properties['Runs'])      { $runs = @($j.Runs) }
        if ($j.PSObject.Properties['Collapsed']) { $col  = @($j.Collapsed | ForEach-Object { "$_" } | Where-Object { $_ }) }
        if ($j.PSObject.Properties['Window'] -and $j.Window) { $win = $j.Window }
        return [PSCustomObject]@{ Pinned = $pin; Runs = $runs; Collapsed = $col; Window = $win }
    } catch { return $default }
}

function Save-ConsoleState {
    param([Parameter(Mandatory)]$State)
    try { $State | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Get-ConsoleStatePath) -Encoding UTF8 -ErrorAction Stop }
    catch { }
}

function Add-ConsoleToolRun {
    <#
        Record that a tool was launched.

        CALLED FROM Start-ToolScript, which is the ONE place every launch goes
        through -- the button grid, the list view, the palette and the group
        chooser all end up there. Recording at each call site instead would
        mean a view added later silently stops counting.
    #>
    param([Parameter(Mandatory)][string]$Name)
    try {
        $s = Get-ConsoleState
        $runs = @($s.Runs | Where-Object { "$($_.Name)" -ne $Name })
        $count = 1
        $prev = @($s.Runs | Where-Object { "$($_.Name)" -eq $Name } | Select-Object -First 1)
        if ($prev.Count) { $count = [int]$prev[0].Count + 1 }
        $runs = @([PSCustomObject]@{ Name = $Name; Count = $count; Last = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') }) + $runs
        # Capped. This is a convenience list, not an audit trail -- the audit
        # trail is CompletedTasks.csv and it is not this file's job to grow
        # without limit alongside it.
        if ($runs.Count -gt 200) { $runs = @($runs[0..199]) }
        $s.Runs = $runs
        Save-ConsoleState -State $s
    } catch { }
}

function Set-ConsoleSectionCollapsed {
    param([Parameter(Mandatory)][string]$Name, [bool]$Collapsed)
    try {
        $s = Get-ConsoleState
        $c = @($s.Collapsed | Where-Object { "$_" -ne $Name })
        if ($Collapsed) { $c = @($c) + @($Name) }
        $s.Collapsed = @($c)
        Save-ConsoleState -State $s
    } catch { }
}

function Save-ConsoleWindowBounds {
    <#
        Remember where the operator left the window. Only a Normal window is
        worth remembering: a maximized one restores to whatever RestoreBounds
        says, and a minimized one has a location of -32000,-32000, which
        would put the console off every screen on the next launch.
    #>
    try {
        if ($form.WindowState -ne [System.Windows.Forms.FormWindowState]::Normal) { return }
        $s = Get-ConsoleState
        $s.Window = [PSCustomObject]@{
            X = $form.Location.X; Y = $form.Location.Y
            W = $form.Size.Width; H = $form.Size.Height
        }
        Save-ConsoleState -State $s
    } catch { }
}

function Restore-ConsoleWindowBounds {
    <#
        Put the window back where it was, IF that place still exists.

        Clamped to the screen that contains the saved point, or the primary
        screen when no screen does -- a monitor that was there last week and
        is not there today must not take the console with it. Size is
        restored only when the operator had zoomed (bigger than design);
        smaller than design would clip the grid, and Update-ConsoleZoom is
        what reflows into the larger size, so it runs after.
    #>
    try {
        $s = Get-ConsoleState
        $w = $s.Window
        if (-not $w) { return $false }
        $x = [int]$w.X; $y = [int]$w.Y; $cw = [int]$w.W; $ch = [int]$w.H
        $pt = New-Object System.Drawing.Point($x, $y)
        $scr = [System.Windows.Forms.Screen]::FromPoint($pt)
        $area = $scr.WorkingArea
        if (-not $area.Contains($pt)) { $area = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea; $x = $area.X; $y = $area.Y }
        # Returns $true ONLY when a size was applied. A location-only restore
        # must not count as "the operator chose this size", or the initial
        # fit would leave a design-size window with scrollbars in it.
        $sized = $false
        if ($cw -gt $form.Width -and $ch -gt $form.Height) {
            $cw = [Math]::Min($cw, $area.Width); $ch = [Math]::Min($ch, $area.Height)
            $form.Size = New-Object System.Drawing.Size($cw, $ch)
            $sized = $true
        }
        if (($x + $form.Width)  -gt ($area.X + $area.Width))  { $x = $area.X + $area.Width  - $form.Width }
        if (($y + $form.Height) -gt ($area.Y + $area.Height)) { $y = $area.Y + $area.Height - $form.Height }
        if ($x -lt $area.X) { $x = $area.X }
        if ($y -lt $area.Y) { $y = $area.Y }
        $form.StartPosition = 'Manual'
        $form.Location = New-Object System.Drawing.Point($x, $y)
        return $sized
    } catch { return $false }
}

function Set-ConsoleToolPinned {
    param([Parameter(Mandatory)][string]$Name, [bool]$Pinned)
    try {
        $s = Get-ConsoleState
        $p = @($s.Pinned | Where-Object { "$_" -ne $Name })
        if ($Pinned) { $p = @($p) + @($Name) }
        $s.Pinned = @($p)
        Save-ConsoleState -State $s
    } catch { }
}

# ---- the model every view reads ---------------------------------------------
function Get-ConsoleToolModel {
    <#
        One row per tool: what it is, whether it can run, and this operator's
        history with it.

        READS Get-ToolCapabilityBlock RATHER THAN RE-DECIDING. There is exactly
        one definition of "this tool cannot run right now" and it lives in that
        function. A second definition here would drift from it.
    #>
    $st = Get-ConsoleState
    $runs = @{}
    foreach ($r in @($st.Runs)) { if ("$($r.Name)") { $runs["$($r.Name)"] = $r } }
    $pinned = @{}
    foreach ($n in @($st.Pinned)) { if ("$n") { $pinned["$n"] = $true } }

    $out = @()
    foreach ($name in $script:SectionNames) {
        foreach ($t in @($script:AllTools | Where-Object { (Get-ToolSection $_) -eq $name })) {
            $blocked = ''
            try {
                $b = Get-ToolCapabilityBlock -Tool $t
                if ($b) { $blocked = "$($b.Reason)"; if (-not $blocked) { $blocked = 'Not available right now' } }
            } catch { }
            $types = @($t.ObjectTypes | Where-Object { "$_".Trim() })
            $detail = $(if ($types.Count) { "works on: $($types -join ', ')" } else { 'any object' })
            $grp = Get-ToolGroupName -Tool $t
            if ($grp) { $detail = "in group '$grp'   --   $detail" }
            $r = $runs["$($t.Name)"]
            $out += [PSCustomObject]@{
                Tool     = $t
                Name     = "$($t.Name)"
                Section  = $name
                Detail   = $detail
                Blocked  = $blocked
                Path     = (Join-Path $PSScriptRoot $t.Script)
                Pinned   = [bool]$pinned["$($t.Name)"]
                RunCount = $(if ($r) { [int]$r.Count } else { 0 })
                LastRun  = $(if ($r) { "$($r.Last)" } else { '' })
            }
        }
    }
    return $out
}

# ---- shared chrome ----------------------------------------------------------
function New-ConsoleViewShell {
    <#
        The window every view lives in: maximised, themed, a top bar with a
        title and a right-hand control cluster, a body, and a hint line.

        FOUR VIEWS SHARING ONE SHELL IS THE POINT. Each one re-implementing a
        top bar is four chances to lay a button under its neighbour -- the
        exact bug the console's own find row took three attempts to shake.
        Here it is solved once, in Update-ConsoleViewShell, by measuring.
    #>
    param([Parameter(Mandatory)][string]$Title, [string]$Hint = '')
    $v = New-Object System.Windows.Forms.Form
    $v.Text = "Admin Console -- $Title"
    $v.StartPosition = 'CenterScreen'
    $v.WindowState = [System.Windows.Forms.FormWindowState]::Maximized
    $v.KeyPreview = $true

    $baseSize = 9
    try { if ($form.Font.Size -gt 0) { $baseSize = [int]$form.Font.Size } } catch { }
    $v | Add-Member -NotePropertyName 'FontSize' -NotePropertyValue ([int]($baseSize + 4)) -Force

    $bar = New-Object System.Windows.Forms.Panel
    $bar.Dock = [System.Windows.Forms.DockStyle]::Top
    $bar.Height = 92
    $v.Controls.Add($bar)

    $lblT = New-Object System.Windows.Forms.Label
    $lblT.Location = New-Object System.Drawing.Point(16, 12)
    $lblT.Text = $Title
    $bar.Controls.Add($lblT)

    $body = New-Object System.Windows.Forms.Panel
    $body.Dock = [System.Windows.Forms.DockStyle]::Fill
    $v.Controls.Add($body)
    $body.BringToFront()

    $lblH = New-Object System.Windows.Forms.Label
    $lblH.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $lblH.Height = 34
    $lblH.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $lblH.Text = "  $Hint"
    $v.Controls.Add($lblH)

    $btnBack = New-Object System.Windows.Forms.Button
    $btnBack.Text = 'Button Layout'
    $bar.Controls.Add($btnBack)
    $btnBack.Add_Click({ $v.Close() })
    $v.Add_KeyDown({
        param($s, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { $v.Close() }
    })

    return [PSCustomObject]@{
        Form = $v; Bar = $bar; Title = $lblT; Body = $body; Hint = $lblH
        Back = $btnBack
        # Laid out right-to-left in this order. A view appends its own controls
        # and they land to the LEFT of Back without any view doing arithmetic.
        Right = [System.Collections.ArrayList]@($btnBack)
        Left  = [System.Collections.ArrayList]@()
    }
}

function Update-ConsoleViewShell {
    <#
        Re-font and re-place the shell from what its controls MEASURE.

        Right-hand controls are positioned by subtraction from the right edge
        using their actual .Width, which is the only version that survives a
        font change -- and a font change is the entire point of the A-/A+
        buttons, so here it is not an edge case, it is the normal path.
    #>
    param([Parameter(Mandatory)]$Shell)
    $v = $Shell.Form
    $sz = [int]$v.FontSize
    $f  = New-Object System.Drawing.Font('Microsoft Sans Serif', $sz)
    $v.Font = $f
    $Shell.Title.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', ([int]($sz + 6)), [System.Drawing.FontStyle]::Bold)
    $Shell.Title.Size = New-Object System.Drawing.Size(([int]($sz * 40)), ([int]($sz * 2.2)))
    $Shell.Hint.Font  = New-Object System.Drawing.Font('Microsoft Sans Serif', ([Math]::Max(8, $sz - 3)))
    $Shell.Hint.Height = [int]($sz * 2.4)

    $rowY = [int]($sz * 2.2) + 14
    $h    = [int]($sz * 2.8)
    $Shell.Bar.Height = $rowY + $h + 14

    foreach ($c in @($Shell.Right)) {
        if (-not $c -or $c.IsDisposed) { continue }
        $c.Font = $f
        if ($c -is [System.Windows.Forms.Button]) {
            $w = [System.Windows.Forms.TextRenderer]::MeasureText("$($c.Text)", $f).Width + [int]($sz * 2.4)
            $c.Size = New-Object System.Drawing.Size($w, $h)
        } elseif ($c -is [System.Windows.Forms.ComboBox]) {
            # FROM THE WIDEST ITEM, not a multiple of the font size. A fixed
            # 5-ems box fitted the arrow and about four characters, so any
            # dropdown with a real word in it opened showing an ellipsis. The
            # 34 is the drop arrow plus the padding either side of the text.
            $wid = 0
            foreach ($it in $c.Items) {
                $m = [System.Windows.Forms.TextRenderer]::MeasureText("$it", $f).Width
                if ($m -gt $wid) { $wid = $m }
            }
            $c.Size = New-Object System.Drawing.Size(([Math]::Max(($wid + 34), [int]($sz * 6))), $h)
        } elseif ($c -is [System.Windows.Forms.Label]) {
            $c.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
            $c.Size = New-Object System.Drawing.Size(([System.Windows.Forms.TextRenderer]::MeasureText("$($c.Text)", $f).Width + 8), $h)
        }
    }
    $right = $Shell.Bar.ClientSize.Width - 16
    foreach ($c in @($Shell.Right)) {
        if (-not $c -or $c.IsDisposed) { continue }
        $right = $right - $c.Width
        $c.Location = New-Object System.Drawing.Point($right, $rowY)
        $right = $right - 8
    }

    $x = 16
    foreach ($c in @($Shell.Left)) {
        if (-not $c -or $c.IsDisposed) { continue }
        $c.Font = $f
        if ($c -is [System.Windows.Forms.Label]) {
            $c.Size = New-Object System.Drawing.Size(([System.Windows.Forms.TextRenderer]::MeasureText("$($c.Text)", $f).Width + 10), $h)
            $c.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        } elseif ($c -is [System.Windows.Forms.TextBox]) {
            $c.Size = New-Object System.Drawing.Size(([int]($sz * 24)), ([int]($sz * 2.2)))
        }
        $c.Location = New-Object System.Drawing.Point($x, $rowY)
        $x = $c.Right + 10
    }
}

function Add-ConsoleViewFontStepper {
    <# A- / A+ on the shell's right cluster, wired to rebuild. #>
    param([Parameter(Mandatory)]$Shell, [Parameter(Mandatory)][scriptblock]$OnChange)
    $v = $Shell.Form
    $plus = New-Object System.Windows.Forms.Button
    $plus.Text = 'A +'; $plus.Tag = 'Plain'
    $Shell.Bar.Controls.Add($plus)
    $minus = New-Object System.Windows.Forms.Button
    $minus.Text = 'A -'; $minus.Tag = 'Plain'
    $Shell.Bar.Controls.Add($minus)
    [void]$Shell.Right.Add($plus)
    [void]$Shell.Right.Add($minus)
    # Clamped 10-28: below 10 a large-text view has no reason to exist, above
    # 28 two words fill a cell.
    #
    # THE HANDLER CANNOT SEE THIS FUNCTION'S VARIABLES. It runs from the
    # message loop, whose current scope is the VIEW function's (the one that
    # called ShowDialog), not this one's -- so $Shell and $OnChange were $null
    # by the time a button was pressed, Update-ConsoleViewShell refused its
    # mandatory parameter, and the font only caught up when something else
    # in the view happened to call Update-ConsoleViewShell. Reported
    # 2026-09-12 as "the catalogue font changes only after moving to a new
    # tool". Everything the handler needs now rides on the button and is read
    # back through $this, the sender. The OnChange block itself still runs in
    # the view function's scope, which is where $lst, $txt and $build live.
    foreach ($pair in @(@($plus, 2), @($minus, -2))) {
        $pair[0] | Add-Member -NotePropertyName 'Stepper' -NotePropertyValue ([PSCustomObject]@{
            Shell = $Shell; OnChange = $OnChange; Delta = [int]$pair[1] }) -Force
        $pair[0].Add_Click({
            try {
                $st = $this.Stepper
                $f  = $st.Shell.Form
                $f.FontSize = [Math]::Max(10, [Math]::Min(28, [int]$f.FontSize + $st.Delta))
                Update-ConsoleViewShell -Shell $st.Shell
                & $st.OnChange
            } catch { }
        })
    }
}

function Get-ConsoleViewColors {
    <#
        The four colours every view needs, resolved once with fallbacks.

        RESOLVED BEFORE ANYTHING IS ASSIGNED. ConvertTo-ThemeColor returns
        $null for a key a theme does not define -- the Console 2 themes have no
        AccentBack -- and assigning $null to a Color property throws AT THE
        POINT OF ASSIGNMENT. A guard that tests the property afterwards has
        already lost; this view crashed exactly that way on its first run.
    #>
    param([Parameter(Mandatory)]$Form)
    $hb = $null; $hf = $null; $tb = $null; $tf = $null
    if ($script:Theme) {
        $hb = ConvertTo-ThemeColor $script:Theme.PanelBack
        $hf = ConvertTo-ThemeColor $script:Theme.FormText
        if (-not $hf) { $hf = ConvertTo-ThemeColor $script:Theme.GroupText }
        $tb = ConvertTo-ThemeColor $script:Theme.ButtonBack
        $tf = ConvertTo-ThemeColor $script:Theme.ButtonText
    }
    if (-not $hb) { $hb = $Form.BackColor }
    if (-not $hf) { $hf = $Form.ForeColor }
    if (-not $tb) { $tb = [System.Drawing.SystemColors]::Control }
    if (-not $tf) { $tf = [System.Drawing.SystemColors]::ControlText }
    return [PSCustomObject]@{
        HeaderBack = $hb; HeaderFore = $hf; TileBack = $tb; TileFore = $tf
        Dim = [System.Drawing.Color]::FromArgb(150, 155, 165)
        Ok  = [System.Drawing.Color]::FromArgb(90, 220, 140)
        Bad = [System.Drawing.Color]::FromArgb(245, 115, 115)
        Warn = [System.Drawing.Color]::FromArgb(240, 175, 60)
    }
}

function Invoke-ConsoleViewTool {
    <#
        Run a tool from inside a view, and put the view back in step afterwards.

        A CONNECT SCRIPT CLOSES BOTH WINDOWS, for the reason the button handler
        explains at length: the sign-in prompt is created on this thread and a
        modal form disables every other window on it. A view is itself modal
        over the console, so it is two forms deep -- both have to go before the
        work at the bottom of this script can run.
    #>
    param([Parameter(Mandatory)]$ViewForm, [Parameter(Mandatory)]$Tool, [string]$Blocked, [scriptblock]$Rebuild)
    if ($Blocked) {
        $null = Show-ToolDialog("$($Tool.Name) cannot run right now.`r`n`r`n$Blocked", 'Not available', 'OK', 'Information')
        return
    }
    if (Get-ConnectButtonService -Tool $Tool) {
        $script:PendingConnect = $Tool
        $ViewForm.Close()
        $form.Close()
        return
    }
    $ViewForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Start-ToolScript -Tool $Tool
        Update-ConnectionButtonColors
        # The tool may have connected to something, so what is available has
        # changed -- and it has just moved to the top of the recent list.
        if ($Rebuild) { & $Rebuild }
    } catch { Show-HandlerError -ErrorRecord $_ -Context "Run $($Tool.Name) from a view" }
    finally { $ViewForm.Cursor = [System.Windows.Forms.Cursors]::Default }
}

function New-ConsoleToolTile {
    <#
        One tool as a full-width button, with the facts that used to be
        columns moved into its tooltip and its colour.
    #>
    param([Parameter(Mandatory)]$Row, [Parameter(Mandatory)]$Colors, [Parameter(Mandatory)][int]$Size,
          [Parameter(Mandatory)]$Tip, [string]$Prefix = '')
    $b = New-Object System.Windows.Forms.Button
    $b.Text = "$Prefix$($Row.Name)"
    $b.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', $Size)
    $b.Dock = [System.Windows.Forms.DockStyle]::Fill
    $b.Height = [int]($Size * 2.9)
    $b.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $b.Padding = New-Object System.Windows.Forms.Padding(([int]($Size * 0.8)), 0, 4, 0)
    $b.Margin = New-Object System.Windows.Forms.Padding(2)
    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $b.UseVisualStyleBackColor = $false
    $b.BackColor = $Colors.TileBack
    $b.ForeColor = $(if ($Row.Blocked) { $Colors.Dim } else { $Colors.TileFore })
    $b.FlatAppearance.BorderSize = 1
    $b.FlatAppearance.BorderColor = $Colors.HeaderBack
    # These tiles were already left-aligned with a padded caption, so an icon
    # drops into the gap that was there for it. -Fallback for the same reason
    # the main grid asks for it: a column of tiles with one blank slot reads as
    # a tool that failed rather than as a tool nobody had a picture for. The
    # gutter it sets is derived from the font, so the size stepper moves both.
    #
    # THE GLYPH IS RESOLVED HERE RATHER THAN LEFT TO Add-ButtonGlyph, because
    # this button's .Tag is set by the CALLER after it comes back -- so the
    # tool object, and with it any Tools.json 'Icon', is not on the control yet.
    # $Row.Name is used rather than the button's caption for the same reason it
    # has to be: $Prefix puts a pin marker in front of a favourite, and an
    # anchored rule cannot see past it.
    if (Test-ConsoleIconStyle) {
        Add-ButtonGlyph -Button $b -Glyph (Get-ToolGlyph -Name "$($Row.Name)" -Icon "$($Row.Tool.Icon)" -Fallback)
    }
    $seen = ''
    if ($Row.PSObject.Properties['LastRun'] -and "$($Row.LastRun)") { $seen = "`r`nlast run $($Row.LastRun)  ($($Row.RunCount)x)" }
    $Tip.SetToolTip($b, $(if ($Row.Blocked) { "$($Row.Name)`r`n$($Row.Detail)$seen`r`n`r`nNOT AVAILABLE: $($Row.Blocked)" } else { "$($Row.Name)`r`n$($Row.Detail)$seen" }))
    return $b
}

# ---- 1. list view -----------------------------------------------------------
function Show-ConsoleListView {
    <#
        FULL SCREEN, BIG TEXT, TOOLS IN COLUMNS, GROUPED BY SECTION.

        NAMES ONLY. The first cut carried 'Runs on' and 'Status' columns beside
        every tool. They were honest and they were wasted space: a list whose
        job is to be readable across the room should not spend two thirds of
        its width on text nobody is reading. Both facts are still there -- in
        the tooltip and in the colour -- and dropping them buys the width to
        run three tools across, which cuts the scrolling to a third.

        A TableLayoutPanel RATHER THAN A ListView. A ListView in Details mode
        gives one item per ROW; three tools across a row would mean three tools
        sharing one selection, and clicking would have to hit-test which cell
        was meant. A table of real buttons has none of that.
    #>
    $sh = New-ConsoleViewShell -Title 'Tool list' -Hint 'Click or press Enter on a tool to run it.  RIGHT-CLICK to pin it (pinned tools show a *).  Tab moves between tools.  Esc returns to the button layout.'
    $v = $sh.Form
    $v | Add-Member -NotePropertyName 'Cols' -NotePropertyValue 3 -Force
    $v | Add-Member -NotePropertyName 'Rows' -NotePropertyValue @() -Force
    $tip = New-Object System.Windows.Forms.ToolTip
    $tip.AutoPopDelay = 20000

    $cbo = New-Object System.Windows.Forms.ComboBox
    $cbo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    foreach ($n in 1, 2, 3, 4, 5) { [void]$cbo.Items.Add("$n") }
    $cbo.SelectedItem = '3'
    $sh.Bar.Controls.Add($cbo)
    $lblCols = New-Object System.Windows.Forms.Label
    $lblCols.Text = 'Columns:'
    $sh.Bar.Controls.Add($lblCols)
    [void]$sh.Right.Add($cbo); [void]$sh.Right.Add($lblCols)

    $lblFind = New-Object System.Windows.Forms.Label; $lblFind.Text = 'Find:'
    $txtFind = New-Object System.Windows.Forms.TextBox
    $lblCount = New-Object System.Windows.Forms.Label
    foreach ($c in @($lblFind, $txtFind, $lblCount)) { $sh.Bar.Controls.Add($c); [void]$sh.Left.Add($c) }

    $scroll = New-Object System.Windows.Forms.Panel
    $scroll.Dock = [System.Windows.Forms.DockStyle]::Fill
    $scroll.AutoScroll = $true
    $sh.Body.Controls.Add($scroll)
    # Docked TOP, not Fill: filled it would be clipped to the visible area and
    # AutoScroll would have nothing to scroll to.
    $tbl = New-Object System.Windows.Forms.TableLayoutPanel
    $tbl.Dock = [System.Windows.Forms.DockStyle]::Top
    $tbl.AutoSize = $true
    $tbl.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $tbl.ColumnCount = 3
    $tbl.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 24)
    $scroll.Controls.Add($tbl)

    $build = {
        param([string]$Filter)
        $sz = [int]$v.FontSize
        $cols = [int]$v.Cols
        $col = Get-ConsoleViewColors -Form $v
        $fh = New-Object System.Drawing.Font('Microsoft Sans Serif', ([int]($sz + 1)), [System.Drawing.FontStyle]::Bold)
        $tbl.SuspendLayout()
        try {
            # Dispose, not just Clear. These are real controls with real
            # handles; clearing the collection alone leaks one per rebuild, and
            # this rebuilds on every keystroke in the Find box.
            $old = @($tbl.Controls); $tbl.Controls.Clear(); $tbl.RowStyles.Clear()
            foreach ($c in $old) { try { $c.Dispose() } catch { } }
            $tbl.ColumnCount = $cols
            $tbl.ColumnStyles.Clear()
            for ($i = 0; $i -lt $cols; $i++) {
                [void]$tbl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, (100.0 / $cols))))
            }
            $flt = "$Filter".Trim()
            $row = 0; $c0 = 0; $shown = 0; $lastSec = ''
            foreach ($m in $v.Rows) {
                if ($flt -and "$($m.Name)" -notmatch [regex]::Escape($flt)) { continue }
                if ($m.Section -ne $lastSec) {
                    if ($c0 -ne 0) { $row++; $c0 = 0 }
                    $h = New-Object System.Windows.Forms.Label
                    $h.Text = "$($m.Section)"; $h.Font = $fh
                    $h.BackColor = $col.HeaderBack; $h.ForeColor = $col.HeaderFore
                    $h.Dock = [System.Windows.Forms.DockStyle]::Fill
                    $h.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                    $h.Height = [int]($sz * 2.6)
                    $h.Margin = New-Object System.Windows.Forms.Padding(2, ([int]($sz * 0.9)), 2, 2)
                    $tbl.Controls.Add($h, 0, $row); $tbl.SetColumnSpan($h, $cols)
                    $row++; $c0 = 0; $lastSec = $m.Section
                }
                $b = New-ConsoleToolTile -Row $m -Colors $col -Size $sz -Tip $tip -Prefix $(if ($m.Pinned) { '* ' } else { '' })
                $b.Tag = $m
                $b.Add_Click({ Invoke-ConsoleViewTool -ViewForm $v -Tool $this.Tag.Tool -Blocked $this.Tag.Blocked -Rebuild { $v.Rows = Get-ConsoleToolModel; & $build $txtFind.Text } })
                # PIN FROM HERE TOO. The Recent view's hint told operators they
                # could right-click a tool "in this view or the list view", and
                # in the list view nothing happened -- the handler was only ever
                # wired in Recent. A hint that names a gesture the window does
                # not support is worse than no hint.
                $b.Add_MouseUp({
                    param($s, $e)
                    if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Right) { return }
                    Set-ConsoleToolPinned -Name $s.Tag.Name -Pinned (-not $s.Tag.Pinned)
                    $v.Rows = Get-ConsoleToolModel
                    & $build $txtFind.Text
                })
                $tbl.Controls.Add($b, $c0, $row)
                $shown++; $c0++
                if ($c0 -ge $cols) { $c0 = 0; $row++ }
            }
            $total = @($v.Rows).Count
            $lblCount.Text = $(if ($flt) { "$shown of $total tools" } else { "$total tools" })
        } finally { $tbl.ResumeLayout() }
        Update-ConsoleViewShell -Shell $sh
    }

    $txtFind.Add_TextChanged({ & $build $txtFind.Text })
    $cbo.Add_SelectedIndexChanged({ $v.Cols = [int]"$($cbo.SelectedItem)"; & $build $txtFind.Text })
    Add-ConsoleViewFontStepper -Shell $sh -OnChange { & $build $txtFind.Text }
    $v.Add_Resize({ try { Update-ConsoleViewShell -Shell $sh; $w = $scroll.ClientSize.Width; if ($w -gt 100) { $tbl.Width = $w } } catch { } })
    $scroll.Add_Resize({ try { $w = $scroll.ClientSize.Width; if ($w -gt 100) { $tbl.Width = $w } } catch { } })

    Set-ConsoleTheme -Form $v
    $v.Rows = Get-ConsoleToolModel
    Update-ConsoleViewShell -Shell $sh
    & $build ''
    $sh.Title.Text = "Tool list  --  $(@($v.Rows).Count) tools"
    [void]$txtFind.Focus()
    [void]$v.ShowDialog($form)
}

# ---- 2. recent and pinned ---------------------------------------------------
function Show-ConsoleRecentView {
    <#
        THE FIFTEEN TOOLS YOU ACTUALLY USE, out of a hundred and twenty seven.

        THREE BANDS, AND THE ORDER IS THE ARGUMENT. Pinned first, because it is
        the only one the operator chose deliberately. Recent second, because
        "what was I just doing" is the commonest question after a tool closes.
        Most used last, because a frequency count is a description of the past
        few months and ages slowly -- useful, but never the thing you wanted
        this second.

        PIN IS A RIGHT-CLICK, not a column of checkboxes. Pinning happens about
        twice a year; giving it permanent screen furniture would cost width on
        every row forever to serve a decision nobody is making today.
    #>
    $sh = New-ConsoleViewShell -Title 'Recent and pinned' -Hint 'Click to run.  RIGHT-CLICK a tool to pin or unpin it.  Esc returns to the button layout.'
    $v = $sh.Form
    $v | Add-Member -NotePropertyName 'Cols' -NotePropertyValue 3 -Force
    $tip = New-Object System.Windows.Forms.ToolTip; $tip.AutoPopDelay = 20000

    $scroll = New-Object System.Windows.Forms.Panel
    $scroll.Dock = [System.Windows.Forms.DockStyle]::Fill; $scroll.AutoScroll = $true
    $sh.Body.Controls.Add($scroll)
    $tbl = New-Object System.Windows.Forms.TableLayoutPanel
    $tbl.Dock = [System.Windows.Forms.DockStyle]::Top
    $tbl.AutoSize = $true; $tbl.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $tbl.ColumnCount = 3
    $tbl.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 24)
    $scroll.Controls.Add($tbl)

    $build = {
        $sz = [int]$v.FontSize
        $cols = [int]$v.Cols
        $col = Get-ConsoleViewColors -Form $v
        $fh = New-Object System.Drawing.Font('Microsoft Sans Serif', ([int]($sz + 1)), [System.Drawing.FontStyle]::Bold)
        $model = Get-ConsoleToolModel

        $pinned = @($model | Where-Object { $_.Pinned } | Sort-Object Name)
        $recent = @($model | Where-Object { $_.LastRun } | Sort-Object { $_.LastRun } -Descending | Select-Object -First 15)
        $often  = @($model | Where-Object { $_.RunCount -gt 0 } | Sort-Object RunCount -Descending | Select-Object -First 15)

        $bands = @()
        $bands += [PSCustomObject]@{ Title = 'Pinned'; Rows = $pinned; Empty = 'Nothing pinned yet -- right-click any tool in this view or the list view to pin it.' }
        $bands += [PSCustomObject]@{ Title = 'Recently used'; Rows = $recent; Empty = 'No history yet. Tools you run from any view appear here.' }
        $bands += [PSCustomObject]@{ Title = 'Most used'; Rows = $often; Empty = '' }

        $tbl.SuspendLayout()
        try {
            $old = @($tbl.Controls); $tbl.Controls.Clear(); $tbl.RowStyles.Clear()
            foreach ($c in $old) { try { $c.Dispose() } catch { } }
            $tbl.ColumnCount = $cols; $tbl.ColumnStyles.Clear()
            for ($i = 0; $i -lt $cols; $i++) {
                [void]$tbl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, (100.0 / $cols))))
            }
            $row = 0
            foreach ($band in $bands) {
                if (@($band.Rows).Count -eq 0 -and -not $band.Empty) { continue }
                $h = New-Object System.Windows.Forms.Label
                $h.Text = "$($band.Title)"; $h.Font = $fh
                $h.BackColor = $col.HeaderBack; $h.ForeColor = $col.HeaderFore
                $h.Dock = [System.Windows.Forms.DockStyle]::Fill
                $h.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                $h.Height = [int]($sz * 2.6)
                $h.Margin = New-Object System.Windows.Forms.Padding(2, ([int]($sz * 0.9)), 2, 2)
                $tbl.Controls.Add($h, 0, $row); $tbl.SetColumnSpan($h, $cols); $row++
                if (@($band.Rows).Count -eq 0) {
                    $e = New-Object System.Windows.Forms.Label
                    $e.Text = "   $($band.Empty)"; $e.ForeColor = $col.Dim
                    $e.Dock = [System.Windows.Forms.DockStyle]::Fill
                    $e.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                    $e.Height = [int]($sz * 2.4)
                    $tbl.Controls.Add($e, 0, $row); $tbl.SetColumnSpan($e, $cols); $row++
                    continue
                }
                $c0 = 0
                foreach ($m in $band.Rows) {
                    $pre = $(if ($m.Pinned) { '* ' } else { '' })
                    $b = New-ConsoleToolTile -Row $m -Colors $col -Size $sz -Tip $tip -Prefix $pre
                    $b.Tag = $m
                    $b.Add_Click({ Invoke-ConsoleViewTool -ViewForm $v -Tool $this.Tag.Tool -Blocked $this.Tag.Blocked -Rebuild { & $build } })
                    $b.Add_MouseUp({
                        param($s, $e)
                        if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Right) { return }
                        Set-ConsoleToolPinned -Name $s.Tag.Name -Pinned (-not $s.Tag.Pinned)
                        & $build
                    })
                    $tbl.Controls.Add($b, $c0, $row)
                    $c0++
                    if ($c0 -ge $cols) { $c0 = 0; $row++ }
                }
                if ($c0 -ne 0) { $row++ }
            }
        } finally { $tbl.ResumeLayout() }
        Update-ConsoleViewShell -Shell $sh
        $sh.Title.Text = "Recent and pinned  --  $(@($pinned).Count) pinned, $(@($recent).Count) recent"
    }

    Add-ConsoleViewFontStepper -Shell $sh -OnChange { & $build }
    $v.Add_Resize({ try { Update-ConsoleViewShell -Shell $sh; $w = $scroll.ClientSize.Width; if ($w -gt 100) { $tbl.Width = $w } } catch { } })
    $scroll.Add_Resize({ try { $w = $scroll.ClientSize.Width; if ($w -gt 100) { $tbl.Width = $w } } catch { } })

    Set-ConsoleTheme -Form $v
    Update-ConsoleViewShell -Shell $sh
    & $build
    [void]$v.ShowDialog($form)
}

# ---- 3. command palette -----------------------------------------------------
function Show-ConsolePaletteView {
    <#
        TYPE THREE LETTERS, PRESS ENTER.

        THE ONE VIEW THAT IS NOT MAXIMISED, and that is the whole design. The
        list view is for browsing; this is for an operator who already knows
        the tool's name and wants it gone from the screen in under a second. A
        full-screen window would make that slower, not faster.

        RANKED, NOT FILTERED ALPHABETICALLY. A name that STARTS with what was
        typed beats one that merely contains it, and among equals the tool used
        most often wins. Typing 'ma' should offer Mailbox Health before Group
        Manager, and after a fortnight it should offer whichever of those this
        operator actually opens.
    #>
    $v = New-Object System.Windows.Forms.Form
    $v.Text = 'Admin Console -- Command palette'
    $v.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $v.StartPosition = 'CenterScreen'
    $v.MaximizeBox = $false; $v.MinimizeBox = $false
    $v.KeyPreview = $true
    $sz = 14
    try { if ($form.Font.Size -gt 0) { $sz = [int]$form.Font.Size + 5 } } catch { }
    $v.ClientSize = New-Object System.Drawing.Size(760, 520)
    $v.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', $sz)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(14, 14)
    $txt.Size = New-Object System.Drawing.Size(732, ([int]($sz * 2.4)))
    $txt.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', ([int]($sz + 4)))
    $v.Controls.Add($txt)

    $lst = New-Object System.Windows.Forms.ListBox
    $lst.Location = New-Object System.Drawing.Point(14, ([int]($sz * 2.4) + 22))
    $lst.Size = New-Object System.Drawing.Size(732, (520 - [int]($sz * 2.4) - 22 - 46))
    $lst.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', $sz)
    $lst.IntegralHeight = $false
    $v.Controls.Add($lst)

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Location = New-Object System.Drawing.Point(14, (520 - 34))
    $lblHint.Size = New-Object System.Drawing.Size(732, 26)
    $lblHint.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', ([Math]::Max(8, $sz - 4)))
    $lblHint.Text = 'Up / Down to choose.  Enter to run.  Esc to close.'
    $v.Controls.Add($lblHint)

    # STATE LIVES ON THE FORM, NOT IN A VARIABLE BESIDE IT.
    #
    # THIS IS WHY Enter DID NOTHING. '$shown = ...' inside the $fill
    # scriptblock creates $shown in $fill's OWN scope; the moment $fill
    # returns that copy is gone, and the outer one is still the empty array it
    # was initialised to. $run then read an empty list, decided the selection
    # was out of range, and returned -- silently, every time.
    #
    # A PROPERTY ASSIGNMENT ON AN OBJECT HAS NO SUCH PROBLEM: $v.Shown = ...
    # mutates the form that both scriptblocks can see. The list view worked
    # from its first run for exactly this reason -- it happened to use $v.Rows.
    $v | Add-Member -NotePropertyName 'Model' -NotePropertyValue (Get-ConsoleToolModel) -Force
    $v | Add-Member -NotePropertyName 'Shown' -NotePropertyValue @() -Force

    $fill = {
        param([string]$Filter)
        $f = "$Filter".Trim()
        $rank = @()
        foreach ($m in $v.Model) {
            if (-not $f) { $rank += [PSCustomObject]@{ M = $m; R = 2 }; continue }
            $n = "$($m.Name)"
            if ($n -imatch ('^' + [regex]::Escape($f))) { $rank += [PSCustomObject]@{ M = $m; R = 0 }; continue }
            if ($n -imatch [regex]::Escape($f))         { $rank += [PSCustomObject]@{ M = $m; R = 1 }; continue }
            # Section is searched too, so 'report' finds everything filed under
            # Reporting even when the tool's own name never says so.
            if ("$($m.Section)" -imatch [regex]::Escape($f)) { $rank += [PSCustomObject]@{ M = $m; R = 3 } }
        }
        $v.Shown = @($rank | Sort-Object @{ Expression = 'R' }, @{ Expression = { -$_.M.RunCount } }, @{ Expression = { $_.M.Name } } | ForEach-Object { $_.M })
        $lst.BeginUpdate()
        try {
            $lst.Items.Clear()
            foreach ($m in $v.Shown) {
                $mark = $(if ($m.Blocked) { '  (not available)' } else { '' })
                [void]$lst.Items.Add("$($m.Name)   --   $($m.Section)$mark")
            }
            if ($lst.Items.Count) { $lst.SelectedIndex = 0 }
        } finally { $lst.EndUpdate() }
        $lblHint.Text = "$($lst.Items.Count) match(es).   Up / Down to choose.  Enter to run.  Esc to close."
    }

    $run = {
        if ($lst.SelectedIndex -lt 0 -or $lst.SelectedIndex -ge @($v.Shown).Count) { return }
        $m = $v.Shown[$lst.SelectedIndex]
        Invoke-ConsoleViewTool -ViewForm $v -Tool $m.Tool -Blocked $m.Blocked -Rebuild {
            $v.Model = Get-ConsoleToolModel
            & $fill $txt.Text
        }
    }

    # ARROWS ARE HANDLED ON THE TEXTBOX so the caret never leaves it. Moving
    # focus to the list to scroll it would mean the next keystroke typed into
    # nothing, which is the classic way a palette feels broken.
    $txt.Add_KeyDown({
        param($s, $e)
        switch ($e.KeyCode) {
            ([System.Windows.Forms.Keys]::Down) {
                if ($lst.Items.Count) { $lst.SelectedIndex = [Math]::Min($lst.Items.Count - 1, $lst.SelectedIndex + 1) }
                $e.Handled = $true; $e.SuppressKeyPress = $true
            }
            ([System.Windows.Forms.Keys]::Up) {
                if ($lst.Items.Count) { $lst.SelectedIndex = [Math]::Max(0, $lst.SelectedIndex - 1) }
                $e.Handled = $true; $e.SuppressKeyPress = $true
            }
            ([System.Windows.Forms.Keys]::Enter) { $e.Handled = $true; $e.SuppressKeyPress = $true; & $run }
            ([System.Windows.Forms.Keys]::Escape) { $v.Close() }
        }
    })
    $txt.Add_TextChanged({ & $fill $txt.Text })
    $lst.Add_DoubleClick({ & $run })
    $v.Add_KeyDown({
        param($s, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { $v.Close() }
    })

    Set-ConsoleTheme -Form $v
    & $fill ''
    [void]$txt.Focus()
    [void]$v.ShowDialog($form)
}

# ---- 4. status board --------------------------------------------------------
function Show-ConsoleStatusView {
    <#
        NOT A VIEW OF TOOLS. A view of the ENVIRONMENT the tools run in --
        what is connected, how long PIM has left, whether this copy of the
        toolkit still matches its manifest, what is sitting in the retry queue,
        and what failed recently.

        THE MORNING CHECK. Every fact here is already somewhere in the console
        or a log file; the value is having them on one screen at the point
        where the answer changes what you do next. Read-only by design -- a
        board that can also change things is a control panel, and a control
        panel is what the other views are for.
    #>
    $sh = New-ConsoleViewShell -Title 'Status' -Hint 'Read-only.  Refresh re-runs every probe.  Esc returns to the button layout.'
    $v = $sh.Form
    $col = Get-ConsoleViewColors -Form $v

    $btnRef = New-Object System.Windows.Forms.Button
    $btnRef.Text = 'Refresh'
    $sh.Bar.Controls.Add($btnRef)
    [void]$sh.Right.Add($btnRef)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true; $txt.ReadOnly = $true; $txt.ScrollBars = 'Vertical'
    $txt.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txt.WordWrap = $false
    $sh.Body.Controls.Add($txt)

    $build = {
        $v.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $sb = New-Object System.Text.StringBuilder
        $line = { param([string]$s) [void]$sb.AppendLine($s) }
        try {
            & $line 'CONNECTIONS'
            & $line ''
            foreach ($p in @(
                @{ N = 'Exchange Online'; T = { (Test-ExoInProcess).Present } },
                @{ N = 'Exchange on-prem'; T = { (Test-EmsAvailable).Available } },
                @{ N = 'Active Directory'; T = { (Test-AdAvailable).Available } })) {
                $ok = $false
                try { $ok = [bool](& $p.T) } catch { }
                & $line ("  {0,-22} {1}" -f $p.N, $(if ($ok) { 'connected' } else { 'NOT connected' }))
            }
            $tenant = ''
            try { $tenant = "$(Get-ConnectedTenantDomain)" } catch { }
            & $line ("  {0,-22} {1}" -f 'Tenant', $(if ($tenant) { $tenant } else { '(not known)' }))
            $pim = ''
            try { $pim = "$(Get-PimStatusLine)" } catch { }
            & $line ("  {0,-22} {1}" -f 'PIM', $(if ($pim) { $pim } else { 'no activation recorded' }))
            & $line ''

            & $line 'TOOLKIT'
            & $line ''
            try {
                $vv = Test-ToolkitIntegrity
                if (-not $vv.HasManifest) {
                    & $line '  unversioned copy (no Version.json)'
                } else {
                    & $line ("  {0,-22} v{1}   released {2}" -f 'Version', $vv.Version, $vv.Released)
                    & $line ("  {0,-22} {1} verified, {2} modified, {3} missing, {4} extra" -f 'Files', $vv.Matched, @($vv.Modified).Count, @($vv.Missing).Count, @($vv.Extra).Count)
                    foreach ($m in @($vv.Modified | Select-Object -First 8)) { & $line "      modified: $m" }
                    foreach ($m in @($vv.Missing  | Select-Object -First 8)) { & $line "      MISSING : $m" }
                }
            } catch { & $line "  integrity check failed: $($_.Exception.Message)" }
            & $line ''

            & $line 'TOOLS'
            & $line ''
            $model = Get-ConsoleToolModel
            $blocked = @($model | Where-Object { $_.Blocked })
            & $line ("  {0,-22} {1}" -f 'Total', @($model).Count)
            & $line ("  {0,-22} {1}" -f 'Blocked right now', @($blocked).Count)
            foreach ($g in @($blocked | Group-Object Blocked | Sort-Object Count -Descending)) {
                & $line ("      {0,-3} {1}" -f $g.Count, $g.Name)
            }
            $missing = @($model | Where-Object { -not (Test-Path -LiteralPath $_.Path) })
            if ($missing.Count) {
                & $line ("  {0,-22} {1}" -f 'Script file missing', $missing.Count)
                foreach ($m in @($missing | Select-Object -First 10)) { & $line "      $($m.Name)  ->  $(Split-Path -Leaf $m.Path)" }
            }
            & $line ''

            & $line 'RETRY QUEUE'
            & $line ''
            try {
                $q = Get-RetryQueuePath
                $files = @()
                if ($q -and (Test-Path -LiteralPath $q)) { $files = @(Get-ChildItem -LiteralPath $q -Filter '*.csv' -File -ErrorAction SilentlyContinue) }
                if ($files.Count -eq 0) { & $line '  empty' }
                else { foreach ($fl in $files) { & $line ("  {0,-34} {1} row(s)   {2}" -f $fl.Name, (@(Import-Csv -LiteralPath $fl.FullName -ErrorAction SilentlyContinue).Count), $fl.LastWriteTime) } }
            } catch { & $line "  could not be read: $($_.Exception.Message)" }
            & $line ''

            & $line 'RECENT ACTIVITY   (from the audit log)'
            & $line ''
            try {
                $lp = Join-Path (Get-LogPath) 'CompletedTasks.csv'
                if (-not (Test-Path -LiteralPath $lp)) { & $line '  no audit log yet' }
                else {
                    $rows = @(Import-Csv -LiteralPath $lp -ErrorAction Stop)
                    $tail = @($rows | Select-Object -Last 12)
                    [array]::Reverse($tail)
                    foreach ($r in $tail) {
                        & $line ("  {0}  {1,-10} {2,-26} {3,-8} {4}" -f "$($r.Timestamp)", "$($r.Operator)", "$($r.Target)", "$($r.Result)", "$($r.Action)")
                    }
                    $fails = @($rows | Where-Object { "$($_.Result)" -match 'Fail' })
                    & $line ''
                    & $line ("  {0} row(s) in the log, {1} recorded as failed." -f $rows.Count, $fails.Count)
                }
            } catch { & $line "  could not be read: $($_.Exception.Message)" }
        } catch {
            & $line ''
            & $line "Status build failed: $($_.Exception.Message)"
        } finally {
            $v.Cursor = [System.Windows.Forms.Cursors]::Default
        }
        $txt.Text = $sb.ToString()
        $txt.Select(0, 0)
    }

    $btnRef.Add_Click({ & $build })
    Add-ConsoleViewFontStepper -Shell $sh -OnChange {
        $txt.Font = New-Object System.Drawing.Font('Consolas', ([int]$v.FontSize))
    }
    $v.Add_Resize({ try { Update-ConsoleViewShell -Shell $sh } catch { } })

    Set-ConsoleTheme -Form $v
    Update-ConsoleViewShell -Shell $sh
    # Monospaced, and not negotiable: every line above is column-aligned with
    # composite formatting, which a proportional font turns into a mess.
    $txt.Font = New-Object System.Drawing.Font('Consolas', ([int]$v.FontSize))
    & $build
    [void]$v.ShowDialog($form)
}

# ---- 5. catalogue -----------------------------------------------------------
function Show-ConsoleCatalogueView {
    <#
        WHAT DOES THIS TOOLKIT ACTUALLY DO. Every tool with the help text from
        its own header, for training somebody new or for the tool you have
        never opened.

        HELP IS READ ON SELECTION, NOT UP FRONT. Reading the TOOLKIT-HELP block
        out of 127 files to build this view would take seconds and would need a
        cache, and a cache needs invalidating -- which is a bug waiting on a
        script somebody edits. One file read per click is instant and is always
        current. The cheapest correct design here is also the fastest one.
    #>
    $sh = New-ConsoleViewShell -Title 'Catalogue' -Hint 'Pick a tool to read what it does.  Run opens it.  Esc returns to the button layout.'
    $v = $sh.Form
    $col = Get-ConsoleViewColors -Form $v

    $lblFind = New-Object System.Windows.Forms.Label; $lblFind.Text = 'Find:'
    $txtFind = New-Object System.Windows.Forms.TextBox
    foreach ($c in @($lblFind, $txtFind)) { $sh.Bar.Controls.Add($c); [void]$sh.Left.Add($c) }

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = 'Run this tool'
    $sh.Bar.Controls.Add($btnRun)
    [void]$sh.Right.Add($btnRun)

    $split = New-Object System.Windows.Forms.SplitContainer
    $split.Dock = [System.Windows.Forms.DockStyle]::Fill
    $split.Orientation = [System.Windows.Forms.Orientation]::Vertical
    $sh.Body.Controls.Add($split)

    $lst = New-Object System.Windows.Forms.ListBox
    $lst.Dock = [System.Windows.Forms.DockStyle]::Fill
    $lst.IntegralHeight = $false
    $split.Panel1.Controls.Add($lst)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true; $txt.ReadOnly = $true; $txt.ScrollBars = 'Vertical'
    $txt.Dock = [System.Windows.Forms.DockStyle]::Fill
    $split.Panel2.Controls.Add($txt)

    # ON THE FORM, NOT BESIDE IT -- see the note in the palette. A '$shown'
    # assigned inside $fillList never reaches $showHelp, which is why the help
    # pane kept showing whichever tool it was built with and the Run button
    # kept its caption no matter what was clicked.
    # SORT, AND THE DEFAULT IS NAMED SO IT IS NOT A MYSTERY.
    #
    # The catalogue has always been in console order -- section by section,
    # then the order the buttons sit in -- which is a defensible default and
    # was invisible: nothing on screen said so, and a list of a hundred tools
    # in no announced order reads as no order at all. Reported 2026-09-12 as
    # "I have no idea how it's sorted".
    #
    # Modified is the file's LastWriteTime. It answers "what changed recently"
    # without leaving the console, which is the question after any delivery.
    $lblSort = New-Object System.Windows.Forms.Label
    $lblSort.Text = 'Sort:'
    $sh.Bar.Controls.Add($lblSort)

    $cboSort = New-Object System.Windows.Forms.ComboBox
    $cboSort.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$cboSort.Items.Add('Console order')
    [void]$cboSort.Items.Add('Name')
    [void]$cboSort.Items.Add('Modified')
    $cboSort.SelectedIndex = 0
    $sh.Bar.Controls.Add($cboSort)

    # THE BOX BEFORE THE LABEL, and that is not a typo. The shell places its
    # right-hand cluster right to left, so the first thing added sits
    # furthest right. Adding the label first would read "[Console order] Sort:"
    # on screen.
    [void]$sh.Right.Add($cboSort)
    [void]$sh.Right.Add($lblSort)

    $v | Add-Member -NotePropertyName 'Model' -NotePropertyValue (Get-ConsoleToolModel) -Force
    $v | Add-Member -NotePropertyName 'Shown' -NotePropertyValue @() -Force

    # STAMPED ONCE, HERE. One stat per tool is nothing; one per keystroke in
    # the filter box, on a hundred files, is a stutter the operator would feel
    # and never connect to a sort they are not using.
    foreach ($m in @($v.Model)) {
        $when = $null
        try { $when = (Get-Item -LiteralPath $m.Path -ErrorAction Stop).LastWriteTime } catch { }
        $m | Add-Member -NotePropertyName 'Modified' -NotePropertyValue $when -Force
    }

    $fillList = {
        param([string]$Filter)
        $f = "$Filter".Trim()
        $rows = @($v.Model | Where-Object { -not $f -or "$($_.Name)" -imatch [regex]::Escape($f) -or "$($_.Section)" -imatch [regex]::Escape($f) })

        # Console order is the model's own order, so it is the one case that
        # sorts by nothing. Modified is newest first -- the reason to ask is
        # almost always "what moved", not "what is oldest" -- and a file that
        # could not be read sorts last rather than crashing the comparison.
        switch ("$($cboSort.SelectedItem)") {
            'Name'     { $rows = @($rows | Sort-Object { "$($_.Name)" }) }
            'Modified' {
                $rows = @($rows | Sort-Object -Descending { if ($_.Modified) { $_.Modified } else { [datetime]::MinValue } })
            }
        }
        $v.Shown = $rows

        $byDate = ("$($cboSort.SelectedItem)" -eq 'Modified')
        $lst.BeginUpdate()
        try {
            $lst.Items.Clear()
            foreach ($m in $v.Shown) {
                # The date is IN the row when that is what you sorted by, and
                # absent otherwise: a date on every line when it is not the
                # ordering is a column of noise.
                $label = "$($m.Name)"
                if ($byDate) {
                    $d = $(if ($m.Modified) { $m.Modified.ToString('yyyy-MM-dd') } else { '----------' })
                    $label = "$d   $label"
                }
                [void]$lst.Items.Add($label)
            }
            if ($lst.Items.Count) { $lst.SelectedIndex = 0 }
            else { $txt.Text = 'Nothing matches that.' }
        } finally { $lst.EndUpdate() }
        $sh.Title.Text = "Catalogue  --  $(@($v.Shown).Count) of $(@($v.Model).Count) tools, by $("$($cboSort.SelectedItem)".ToLower())"
    }

    # SelectedIndexChanged, not Click: a dropdown list changes by keyboard too.
    $cboSort.Add_SelectedIndexChanged({
        try { & $fillList $txtFind.Text } catch { }
    })

    $showHelp = {
        if ($lst.SelectedIndex -lt 0 -or $lst.SelectedIndex -ge @($v.Shown).Count) { return }
        $m = $v.Shown[$lst.SelectedIndex]
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("$($m.Name)")
        [void]$sb.AppendLine(('=' * [Math]::Max(8, "$($m.Name)".Length)))
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("Section : $($m.Section)")
        [void]$sb.AppendLine("Script  : $(Split-Path -Leaf $m.Path)")
        [void]$sb.AppendLine("Modified: $(if ($m.Modified) { $m.Modified.ToString('yyyy-MM-dd HH:mm') } else { 'unknown' })")
        [void]$sb.AppendLine("Applies : $($m.Detail)")
        if ($m.LastRun) { [void]$sb.AppendLine("History : last run $($m.LastRun), $($m.RunCount) time(s)") }
        [void]$sb.AppendLine("Status  : $(if ($m.Blocked) { "NOT AVAILABLE -- $($m.Blocked)" } else { 'available' })")
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('---------------------------------------------------------------')
        [void]$sb.AppendLine('')
        $help = $null
        try { $help = Get-ToolHelpText -Path $m.Path } catch { }
        if ($help) { [void]$sb.AppendLine(($help -join "`r`n")) }
        else {
            [void]$sb.AppendLine('No overview has been written for this tool yet.')
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('Add one between TOOLKIT-HELP-BEGIN and TOOLKIT-HELP-END in the')
            [void]$sb.AppendLine('script header and it will appear here.')
        }
        $txt.Text = $sb.ToString()
        $txt.Select(0, 0)
        $btnRun.Text = "Run $($m.Name)"
        Update-ConsoleViewShell -Shell $sh
    }

    $lst.Add_SelectedIndexChanged({ & $showHelp })
    $txtFind.Add_TextChanged({ & $fillList $txtFind.Text })
    $btnRun.Add_Click({
        if ($lst.SelectedIndex -lt 0 -or $lst.SelectedIndex -ge @($v.Shown).Count) { return }
        $m = $v.Shown[$lst.SelectedIndex]
        Invoke-ConsoleViewTool -ViewForm $v -Tool $m.Tool -Blocked $m.Blocked -Rebuild {
            $v.Model = Get-ConsoleToolModel
            & $fillList $txtFind.Text
        }
    })
    Add-ConsoleViewFontStepper -Shell $sh -OnChange {
        $f = New-Object System.Drawing.Font('Microsoft Sans Serif', ([int]$v.FontSize))
        $lst.Font = $f
        $txt.Font = New-Object System.Drawing.Font('Consolas', ([int]$v.FontSize))
    }
    $v.Add_Resize({
        try {
            Update-ConsoleViewShell -Shell $sh
            if ($split.Width -gt 400) { $split.SplitterDistance = [int]($split.Width * 0.32) }
        } catch { }
    })

    Set-ConsoleTheme -Form $v
    Update-ConsoleViewShell -Shell $sh
    $lst.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', ([int]$v.FontSize))
    $txt.Font = New-Object System.Drawing.Font('Consolas', ([int]$v.FontSize))
    & $fillList ''
    & $showHelp
    try { if ($split.Width -gt 400) { $split.SplitterDistance = [int]($split.Width * 0.32) } } catch { }
    [void]$v.ShowDialog($form)
}

# ---- 6. 3D buttons ----------------------------------------------------------
function New-Console3DTile {
    <#
        One tool as a raised, bevelled button.

        OWNER-DRAWN, AND THE CAPTION IS DRAWN HERE TOO. The button's own Text
        is left empty on purpose: a Flat button still paints its caption, so
        leaving it set would draw the words twice -- once by WinForms in the
        system position, once by this code in the bevelled one -- and the
        offset between them reads as a rendering fault.

        EVERY COLOUR IS PRECOMPUTED ONTO THE TAG, NOT LOOKED UP IN THE PAINT
        HANDLER. A Paint handler runs from the message loop, where the scope it
        was written in is not the scope it executes in -- the lesson this
        console has now learned from the status strip, the settings cache and
        the palette view. $s.Tag is a property on the sender, so it means the
        same thing from anywhere, and painting cannot fail because a variable
        went out of reach.
    #>
    param(
        [Parameter(Mandatory)]$Row,
        [Parameter(Mandatory)]$Colors,
        [Parameter(Mandatory)][int]$Size,
        [Parameter(Mandatory)]$Tip
    )
    $face = $Colors.TileBack
    $b = New-Object System.Windows.Forms.Button
    $b.Text = ''
    $b.Dock = [System.Windows.Forms.DockStyle]::Fill
    $b.Height = [int]($Size * 4.2)
    $b.Margin = New-Object System.Windows.Forms.Padding(5)
    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $b.FlatAppearance.BorderSize = 0
    $b.FlatAppearance.MouseOverBackColor = $face
    $b.FlatAppearance.MouseDownBackColor = $face
    $b.UseVisualStyleBackColor = $false
    $b.BackColor = $face

    $b.Tag = [PSCustomObject]@{
        Row     = $Row
        Tool    = $Row.Tool
        Name    = $Row.Name
        Blocked = "$($Row.Blocked)"
        Caption = "$(if ($Row.Pinned) { '* ' } else { '' })$($Row.Name)"
        Font    = New-Object System.Drawing.Font('Microsoft Sans Serif', $Size, [System.Drawing.FontStyle]::Bold)
        # A raised edge is the face lightened; a shadow is the face darkened
        # -- derived rather than fixed, so the bevel reads correctly on a dark
        # theme and a light one without either being a special case.
        Top     = Get-BlendedColor -From $face -To ([System.Drawing.Color]::White) -Amount 0.22
        Bottom  = Get-BlendedColor -From $face -To ([System.Drawing.Color]::Black) -Amount 0.20
        Light   = Get-BlendedColor -From $face -To ([System.Drawing.Color]::White) -Amount 0.55
        Dark    = Get-BlendedColor -From $face -To ([System.Drawing.Color]::Black) -Amount 0.45
        Fore    = $(if ($Row.Blocked) { $Colors.Dim } else { $Colors.TileFore })
        Pressed = $false
        Hot     = $false
    }

    $b.Add_Paint({
        param($s, $e)
        $t = $s.Tag
        if (-not $t) { return }
        $g = $e.Graphics
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::None
        $w = $s.ClientSize.Width; $h = $s.ClientSize.Height
        if ($w -le 2 -or $h -le 2) { return }
        $rect = New-Object System.Drawing.Rectangle(0, 0, $w, $h)

        # PRESSED INVERTS THE BEVEL RATHER THAN MOVING THE BUTTON. Shifting the
        # control itself would relayout the whole table on every click; swapping
        # which edges are lit costs one repaint and is what the eye actually
        # reads as "pushed in".
        $c1 = $(if ($t.Pressed) { $t.Bottom } else { $t.Top })
        $c2 = $(if ($t.Pressed) { $t.Top }    else { $t.Bottom })
        if ($t.Hot -and -not $t.Pressed) {
            $c1 = Get-BlendedColor -From $c1 -To ([System.Drawing.Color]::White) -Amount 0.12
            $c2 = Get-BlendedColor -From $c2 -To ([System.Drawing.Color]::White) -Amount 0.12
        }

        $br = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, $c1, $c2, 90.0)
        try { $g.FillRectangle($br, $rect) } finally { $br.Dispose() }

        $hi = $(if ($t.Pressed) { $t.Dark }  else { $t.Light })
        $lo = $(if ($t.Pressed) { $t.Light } else { $t.Dark })
        $pHi = New-Object System.Drawing.Pen($hi, 1)
        $pLo = New-Object System.Drawing.Pen($lo, 1)
        try {
            # Two-pixel bevel: the outer line carries the contrast, the inner
            # one softens it. One pixel alone reads as a border, not a shape.
            $g.DrawLine($pHi, 0, 0, ($w - 1), 0)
            $g.DrawLine($pHi, 0, 0, 0, ($h - 1))
            $g.DrawLine($pHi, 1, 1, ($w - 2), 1)
            $g.DrawLine($pHi, 1, 1, 1, ($h - 2))
            $g.DrawLine($pLo, 0, ($h - 1), ($w - 1), ($h - 1))
            $g.DrawLine($pLo, ($w - 1), 0, ($w - 1), ($h - 1))
            $g.DrawLine($pLo, 1, ($h - 2), ($w - 2), ($h - 2))
            $g.DrawLine($pLo, ($w - 2), 1, ($w - 2), ($h - 2))
        } finally { $pHi.Dispose(); $pLo.Dispose() }

        $pad = [int]($s.Font.Size * 1.2)
        $off = $(if ($t.Pressed) { 1 } else { 0 })
        $tr = New-Object System.Drawing.Rectangle(($pad + $off), $off, ($w - ($pad * 2)), $h)
        $flags = [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor
                 [System.Windows.Forms.TextFormatFlags]::Left -bor
                 [System.Windows.Forms.TextFormatFlags]::EndEllipsis -bor
                 [System.Windows.Forms.TextFormatFlags]::NoPrefix
        [System.Windows.Forms.TextRenderer]::DrawText($g, "$($t.Caption)", $t.Font, $tr, $t.Fore, $flags)

        if ($s.Focused) {
            $fr = New-Object System.Drawing.Rectangle(3, 3, ($w - 7), ($h - 7))
            [System.Windows.Forms.ControlPaint]::DrawFocusRectangle($g, $fr)
        }
    })

    $b.Add_MouseDown({ param($s, $e) $s.Tag.Pressed = $true;  $s.Invalidate() })
    $b.Add_MouseUp({   param($s, $e) $s.Tag.Pressed = $false; $s.Invalidate() })
    $b.Add_MouseEnter({ $this.Tag.Hot = $true;  $this.Invalidate() })
    $b.Add_MouseLeave({ $this.Tag.Hot = $false; $this.Tag.Pressed = $false; $this.Invalidate() })
    $b.Add_GotFocus({  $this.Invalidate() })
    $b.Add_LostFocus({ $this.Invalidate() })

    $seen = ''
    if ("$($Row.LastRun)") { $seen = "`r`nlast run $($Row.LastRun)  ($($Row.RunCount)x)" }
    $Tip.SetToolTip($b, $(if ($Row.Blocked) { "$($Row.Name)`r`n$($Row.Detail)$seen`r`n`r`nNOT AVAILABLE: $($Row.Blocked)" } else { "$($Row.Name)`r`n$($Row.Detail)$seen" }))
    return $b
}

function Show-Console3DView {
    <#
        THE SAME TOOLS AS RAISED BUTTONS, full screen and grouped by section.

        Fewer across than the list view by default, because the whole point is
        that each one looks like something you press -- and a bevel needs room
        to read as a bevel. Four columns of chunky buttons is a different thing
        from ten columns of labels, and squeezing these to the same density
        would just be the list view with decoration.
    #>
    $sh = New-ConsoleViewShell -Title '3D buttons' -Hint 'Click or press Enter to run.  RIGHT-CLICK to pin (pinned tools show a *).  Esc returns to the button layout.'
    $v = $sh.Form
    $v | Add-Member -NotePropertyName 'Cols' -NotePropertyValue 4 -Force
    $v | Add-Member -NotePropertyName 'Rows' -NotePropertyValue @() -Force
    $tip = New-Object System.Windows.Forms.ToolTip
    $tip.AutoPopDelay = 20000

    $cbo = New-Object System.Windows.Forms.ComboBox
    $cbo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    foreach ($n in 2, 3, 4, 5, 6) { [void]$cbo.Items.Add("$n") }
    $cbo.SelectedItem = '4'
    $sh.Bar.Controls.Add($cbo)
    $lblCols = New-Object System.Windows.Forms.Label
    $lblCols.Text = 'Columns:'
    $sh.Bar.Controls.Add($lblCols)
    [void]$sh.Right.Add($cbo); [void]$sh.Right.Add($lblCols)

    $lblFind = New-Object System.Windows.Forms.Label; $lblFind.Text = 'Find:'
    $txtFind = New-Object System.Windows.Forms.TextBox
    $lblCount = New-Object System.Windows.Forms.Label
    foreach ($c in @($lblFind, $txtFind, $lblCount)) { $sh.Bar.Controls.Add($c); [void]$sh.Left.Add($c) }

    $scroll = New-Object System.Windows.Forms.Panel
    $scroll.Dock = [System.Windows.Forms.DockStyle]::Fill
    $scroll.AutoScroll = $true
    $sh.Body.Controls.Add($scroll)
    $tbl = New-Object System.Windows.Forms.TableLayoutPanel
    $tbl.Dock = [System.Windows.Forms.DockStyle]::Top
    $tbl.AutoSize = $true
    $tbl.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $tbl.ColumnCount = 4
    $tbl.Padding = New-Object System.Windows.Forms.Padding(10, 10, 10, 28)
    $scroll.Controls.Add($tbl)

    $build = {
        param([string]$Filter)
        $sz = [int]$v.FontSize
        $cols = [int]$v.Cols
        $col = Get-ConsoleViewColors -Form $v
        $fh = New-Object System.Drawing.Font('Microsoft Sans Serif', ([int]($sz + 2)), [System.Drawing.FontStyle]::Bold)
        $tbl.SuspendLayout()
        try {
            # Disposed, not just cleared. Each tile owns a Font and a Tag and
            # is rebuilt on every keystroke in Find; clearing the collection
            # alone leaks one handle per tool per character typed.
            $old = @($tbl.Controls); $tbl.Controls.Clear(); $tbl.RowStyles.Clear()
            foreach ($c in $old) { try { $c.Dispose() } catch { } }
            $tbl.ColumnCount = $cols
            $tbl.ColumnStyles.Clear()
            for ($i = 0; $i -lt $cols; $i++) {
                [void]$tbl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, (100.0 / $cols))))
            }
            $flt = "$Filter".Trim()
            $row = 0; $c0 = 0; $shown = 0; $lastSec = ''
            foreach ($m in $v.Rows) {
                if ($flt -and "$($m.Name)" -notmatch [regex]::Escape($flt)) { continue }
                if ($m.Section -ne $lastSec) {
                    if ($c0 -ne 0) { $row++; $c0 = 0 }
                    $h = New-Object System.Windows.Forms.Label
                    $h.Text = "  $($m.Section)"; $h.Font = $fh
                    $h.BackColor = $col.HeaderBack; $h.ForeColor = $col.HeaderFore
                    $h.Dock = [System.Windows.Forms.DockStyle]::Fill
                    $h.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                    $h.Height = [int]($sz * 2.8)
                    $h.Margin = New-Object System.Windows.Forms.Padding(5, ([int]($sz * 1.1)), 5, 4)
                    $tbl.Controls.Add($h, 0, $row); $tbl.SetColumnSpan($h, $cols)
                    $row++; $c0 = 0; $lastSec = $m.Section
                }
                $b = New-Console3DTile -Row $m -Colors $col -Size $sz -Tip $tip
                $b.Add_Click({ Invoke-ConsoleViewTool -ViewForm $v -Tool $this.Tag.Tool -Blocked $this.Tag.Blocked -Rebuild { $v.Rows = Get-ConsoleToolModel; & $build $txtFind.Text } })
                $b.Add_MouseUp({
                    param($s, $e)
                    if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Right) { return }
                    Set-ConsoleToolPinned -Name $s.Tag.Name -Pinned (-not $s.Tag.Row.Pinned)
                    $v.Rows = Get-ConsoleToolModel
                    & $build $txtFind.Text
                })
                $tbl.Controls.Add($b, $c0, $row)
                $shown++; $c0++
                if ($c0 -ge $cols) { $c0 = 0; $row++ }
            }
            $total = @($v.Rows).Count
            $lblCount.Text = $(if ($flt) { "$shown of $total tools" } else { "$total tools" })
        } finally { $tbl.ResumeLayout() }
        Update-ConsoleViewShell -Shell $sh
    }

    $txtFind.Add_TextChanged({ & $build $txtFind.Text })
    $cbo.Add_SelectedIndexChanged({ $v.Cols = [int]"$($cbo.SelectedItem)"; & $build $txtFind.Text })
    Add-ConsoleViewFontStepper -Shell $sh -OnChange { & $build $txtFind.Text }
    $v.Add_Resize({ try { Update-ConsoleViewShell -Shell $sh; $w = $scroll.ClientSize.Width; if ($w -gt 100) { $tbl.Width = $w } } catch { } })
    $scroll.Add_Resize({ try { $w = $scroll.ClientSize.Width; if ($w -gt 100) { $tbl.Width = $w } } catch { } })

    Set-ConsoleTheme -Form $v
    $v.Rows = Get-ConsoleToolModel
    Update-ConsoleViewShell -Shell $sh
    & $build ''
    $sh.Title.Text = "3D buttons  --  $(@($v.Rows).Count) tools"
    [void]$txtFind.Focus()
    [void]$v.ShowDialog($form)
}
# ---- the registry and the button --------------------------------------------
# ---- 8. what I can run right now ---------------------------------------------
function Show-ConsoleReadyView {
    <#
        THE SAME TOOLS, GROUPED BY WHAT THEY NEED, WITH THE CONNECTED GROUPS
        FIRST.

        Built for a console with a hundred tools across three kinds of
        connection. An operator who has only signed in to Exchange Online does
        not want to read the grid and work out which buttons are grey; they
        want the tools they can use right now at the top and everything else
        below, labelled with what it is waiting for.

        THE GROUPING COMES FROM THE SAME ANSWER THAT GREYS THE BUTTONS.
        Get-ToolAvailability is what the grid asks, so this view cannot say a
        tool is ready while the grid says it is not. It is the list view with
        a different sort key and nothing else changed, on purpose: the tiles,
        the pinning, the find box and the font stepper are all inherited.
    #>
    $sh = New-ConsoleViewShell -Title 'What I can run' -Hint 'Tools you can run right now are at the top. Each group says what it needs and whether that is connected.  Click to run.  RIGHT-CLICK to pin.  Esc returns to the button layout.'
    $v = $sh.Form
    $v | Add-Member -NotePropertyName 'Cols' -NotePropertyValue 3 -Force
    $v | Add-Member -NotePropertyName 'Rows' -NotePropertyValue @() -Force
    $tip = New-Object System.Windows.Forms.ToolTip
    $tip.AutoPopDelay = 20000

    $cbo = New-Object System.Windows.Forms.ComboBox
    $cbo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    foreach ($n in 1, 2, 3, 4, 5) { [void]$cbo.Items.Add("$n") }
    $cbo.SelectedItem = '3'
    $sh.Bar.Controls.Add($cbo)
    $lblCols = New-Object System.Windows.Forms.Label
    $lblCols.Text = 'Columns:'
    $sh.Bar.Controls.Add($lblCols)
    [void]$sh.Right.Add($cbo); [void]$sh.Right.Add($lblCols)

    $lblFind = New-Object System.Windows.Forms.Label; $lblFind.Text = 'Find:'
    $txtFind = New-Object System.Windows.Forms.TextBox
    $lblCount = New-Object System.Windows.Forms.Label
    foreach ($c in @($lblFind, $txtFind, $lblCount)) { $sh.Bar.Controls.Add($c); [void]$sh.Left.Add($c) }

    $scroll = New-Object System.Windows.Forms.Panel
    $scroll.Dock = [System.Windows.Forms.DockStyle]::Fill
    $scroll.AutoScroll = $true
    $sh.Body.Controls.Add($scroll)
    $tbl = New-Object System.Windows.Forms.TableLayoutPanel
    $tbl.Dock = [System.Windows.Forms.DockStyle]::Top
    $tbl.AutoSize = $true
    $tbl.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $tbl.ColumnCount = 3
    $tbl.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 24)
    $scroll.Controls.Add($tbl)

    # One row per tool, each carrying its Needs label and whether every need
    # is met. Unknown availability (never classified) counts as ready, which
    # is the grid's fail-open rule too.
    $classify = {
        $map = $null
        try { $map = Get-CapabilityStateMap } catch { }
        $rows = @()
        foreach ($m in @(Get-ConsoleToolModel)) {
            $needs = @(); $ok = $true
            $a = $null
            try { $a = $script:ToolAvail["$($m.Name)"] } catch { }
            if ($a) {
                foreach ($id in @($a.Requires)) {
                    $lbl = "$id"; $up = $true
                    try {
                        if ($map -and $map[$id]) {
                            if ("$($map[$id].Label)") { $lbl = "$($map[$id].Label)" }
                            $up = [bool]$map[$id].Available
                        }
                    } catch { }
                    $needs += $lbl
                    if (-not $up) { $ok = $false }
                }
                if (-not $a.Enabled) { $ok = $false }
            }
            $needText = $(if ($needs.Count) { ($needs | Sort-Object -Unique) -join ' + ' } else { 'No connection needed' })
            $m | Add-Member -NotePropertyName 'Needs' -NotePropertyValue $needText -Force
            $m | Add-Member -NotePropertyName 'Ready' -NotePropertyValue $ok -Force
            $rows += $m
        }
        # Ready groups first, then by what they need, then by name.
        return @($rows | Sort-Object @{ Expression = { -not $_.Ready } }, Needs, Name)
    }

    $build = {
        param([string]$Filter)
        $sz = [int]$v.FontSize
        $cols = [int]$v.Cols
        $col = Get-ConsoleViewColors -Form $v
        $fh = New-Object System.Drawing.Font('Microsoft Sans Serif', ([int]($sz + 1)), [System.Drawing.FontStyle]::Bold)
        $tbl.SuspendLayout()
        try {
            $old = @($tbl.Controls); $tbl.Controls.Clear(); $tbl.RowStyles.Clear()
            foreach ($c in $old) { try { $c.Dispose() } catch { } }
            $tbl.ColumnCount = $cols
            $tbl.ColumnStyles.Clear()
            for ($i = 0; $i -lt $cols; $i++) {
                [void]$tbl.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, (100.0 / $cols))))
            }
            $flt = "$Filter".Trim()
            $row = 0; $c0 = 0; $shown = 0; $ready = 0; $lastKey = $null
            foreach ($m in $v.Rows) {
                if ($flt -and "$($m.Name)" -notmatch [regex]::Escape($flt)) { continue }
                $key = "$($m.Ready)|$($m.Needs)"
                if ($key -ne $lastKey) {
                    if ($c0 -ne 0) { $row++; $c0 = 0 }
                    $h = New-Object System.Windows.Forms.Label
                    $h.Text = $(if ($m.Ready) { "READY  --  needs: $($m.Needs)" } else { "NOT CONNECTED  --  needs: $($m.Needs)" })
                    $h.Font = $fh
                    $h.BackColor = $col.HeaderBack
                    $h.ForeColor = $(if ($m.Ready) { $col.HeaderFore } else { $col.Dim })
                    $h.Dock = [System.Windows.Forms.DockStyle]::Fill
                    $h.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
                    $h.Height = [int]($sz * 2.6)
                    $h.Margin = New-Object System.Windows.Forms.Padding(2, ([int]($sz * 0.9)), 2, 2)
                    $tbl.Controls.Add($h, 0, $row); $tbl.SetColumnSpan($h, $cols)
                    $row++; $c0 = 0; $lastKey = $key
                }
                $b = New-ConsoleToolTile -Row $m -Colors $col -Size $sz -Tip $tip -Prefix $(if ($m.Pinned) { '* ' } else { '' })
                $b.Tag = $m
                $b.Add_Click({ Invoke-ConsoleViewTool -ViewForm $v -Tool $this.Tag.Tool -Blocked $this.Tag.Blocked -Rebuild { $v.Rows = & $classify; & $build $txtFind.Text } })
                $b.Add_MouseUp({
                    param($s, $e)
                    if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Right) { return }
                    Set-ConsoleToolPinned -Name $s.Tag.Name -Pinned (-not $s.Tag.Pinned)
                    $v.Rows = & $classify
                    & $build $txtFind.Text
                })
                $tbl.Controls.Add($b, $c0, $row)
                $shown++; $c0++
                if ($m.Ready) { $ready++ }
                if ($c0 -ge $cols) { $c0 = 0; $row++ }
            }
            $total = @($v.Rows).Count
            $lblCount.Text = $(if ($flt) { "$shown of $total tools" } else { "$ready of $total ready" })
        } finally { $tbl.ResumeLayout() }
        Update-ConsoleViewShell -Shell $sh
    }

    $txtFind.Add_TextChanged({ & $build $txtFind.Text })
    $cbo.Add_SelectedIndexChanged({ $v.Cols = [int]"$($cbo.SelectedItem)"; & $build $txtFind.Text })
    Add-ConsoleViewFontStepper -Shell $sh -OnChange { & $build $txtFind.Text }
    $v.Add_Resize({ try { Update-ConsoleViewShell -Shell $sh; $w = $scroll.ClientSize.Width; if ($w -gt 100) { $tbl.Width = $w } } catch { } })
    $scroll.Add_Resize({ try { $w = $scroll.ClientSize.Width; if ($w -gt 100) { $tbl.Width = $w } } catch { } })

    Set-ConsoleTheme -Form $v
    $v.Rows = & $classify
    Update-ConsoleViewShell -Shell $sh
    & $build ''
    $readyN = @($v.Rows | Where-Object { $_.Ready }).Count
    $sh.Title.Text = "What I can run  --  $readyN of $(@($v.Rows).Count) ready"
    [void]$txtFind.Focus()
    [void]$v.ShowDialog($form)
}

$script:ConsoleViews = @(
    [PSCustomObject]@{ Id = 'Buttons';   Name = 'Button Layout';                    Note = 'The grid of buttons. The default.'; Show = $null },
    [PSCustomObject]@{ Id = 'List';      Name = 'Tool list - large text, columns';   Note = 'Every tool as a full-width row, grouped by section, in a bigger font.'; Show = { Show-ConsoleListView } },
    [PSCustomObject]@{ Id = 'Recent';    Name = 'Recent and pinned';                 Note = 'The handful you actually use. Right-click to pin.'; Show = { Show-ConsoleRecentView } },
    [PSCustomObject]@{ Id = 'Palette';   Name = 'Command palette';                   Note = 'Type three letters, press Enter. For when you know the name.'; Show = { Show-ConsolePaletteView } },
    [PSCustomObject]@{ Id = 'Status';    Name = 'Status board';                      Note = 'Connections, PIM, integrity, retry queue, recent activity.'; Show = { Show-ConsoleStatusView } },
    [PSCustomObject]@{ Id = 'Catalogue'; Name = 'Catalogue - what each tool does';   Note = 'Every tool with the help text from its own header.'; Show = { Show-ConsoleCatalogueView } },
    [PSCustomObject]@{ Id = 'Buttons3D';  Name = '3D buttons';                          Note = 'Raised, bevelled tiles. Big targets, grouped by section.'; Show = { Show-Console3DView } },
    [PSCustomObject]@{ Id = 'Ready';      Name = 'What I can run right now';            Note = 'Grouped by what each tool needs, connected groups first.'; Show = { Show-ConsoleReadyView } }
)

function Show-ConsoleViewMenu {
    <#
        The View button's menu, built from $script:ConsoleViews every time it
        opens rather than once at startup, so a view added to the list appears
        without anything else being touched.
    #>
    $menu = New-Object System.Windows.Forms.ContextMenuStrip
    foreach ($vw in $script:ConsoleViews) {
        $item = New-Object System.Windows.Forms.ToolStripMenuItem
        $item.Text = $vw.Name
        $item.ToolTipText = $vw.Note
        $item.Checked = ($vw.Id -eq 'Buttons')   # the console itself is what is on screen
        $item.Tag = $vw
        $item.Add_Click({
            $view = $this.Tag
            if (-not $view.Show) { return }      # Buttons: already here
            try { & $view.Show }
            catch { Show-HandlerError -ErrorRecord $_ -Context "Open the '$($view.Name)' view" }
        })
        [void]$menu.Items.Add($item)
    }
    # COLOURED BY HAND, NOT THROUGH Set-ConsoleTheme. That function walks a
    # form's Controls collection; a ContextMenuStrip holds Items instead, so
    # passing one in does nothing useful and would leave a system-white menu
    # hanging off a dark console.
    if ($script:Theme) {
        $bk = ConvertTo-ThemeColor $script:Theme.PanelBack
        $fg = ConvertTo-ThemeColor $script:Theme.FormText
        if ($bk) { $menu.BackColor = $bk }
        if ($fg) { $menu.ForeColor = $fg }
        foreach ($mi in $menu.Items) {
            if ($bk) { $mi.BackColor = $bk }
            if ($fg) { $mi.ForeColor = $fg }
        }
    }
    $menu.Show($btnView, (New-Object System.Drawing.Point(0, $btnView.Height)))
}
$btnView.Add_Click({ Show-ConsoleViewMenu })
#endregion

# -Refresh: the resolved theme is cached in a GLOBAL so tools inherit it,
# and a global outlives this script. A console reopened from the same
# PowerShell window (a launcher, or just re-running the script) would
# otherwise start with whatever was resolved LAST time, and a theme picked
# in Theme Manager in between would appear not to load.
$script:Theme = Get-ConsoleTheme -Refresh
if ($script:Theme) {
    Set-ConsoleTheme -Form $form -Theme $script:Theme
    # THE THEME HAS JUST RESIZED THE BUTTONS. Re-place the find row from what
    # they now measure; see Set-ConsoleFindRowLayout for why this call is the
    # whole fix rather than a tidy-up.
    try { Set-ConsoleFindRowLayout } catch { }
    try {
        $tt = ConvertTo-ThemeColor $script:Theme.FormText; if ($tt) { $lblTitle.ForeColor = $tt }
        # Get-ConsoleTitleFont, NOT Get-ThemeUiFont. The theme pass runs after
        # the title was built, so re-facing it from the theme alone would undo
        # a configured brand font every launch -- it would appear correctly for
        # the instant before the theme applied, which is the most confusing
        # possible outcome.
        $lblTitle.Font = (Get-ConsoleTitleFont -Size 15)
    } catch { }
    # lblWarn was coloured here -- it was NoTheme-tagged, so the theme pass
    # skipped it and it needed doing by hand. Removed with the label on
    # 2026-09-13.

    # The "signed in as ..." caption was re-placed and re-coloured here. Removed
    # with the label on 2026-09-13.
}

# SECTION ACCENTS, APPLIED AFTER THE THEME AND NOT BEFORE. Set-ConsoleTheme
# sets every Label to the theme's text colour and every panel to its panel
# colour, so anything set earlier is overwritten. Doing it here means an accent
# lands on top of whatever the theme chose: the heading takes the accent
# outright, and the panel behind the buttons is tinted a tenth of the way
# toward it -- enough to read as "this band is Reporting", not enough to fight
# the buttons sitting on it. Sections without an accent are untouched.
foreach ($secName in $script:SectionNames) {
    $hex = "$($script:SectionAccents[$secName])"
    if (-not $hex) { continue }
    $ac = ConvertTo-ThemeColor $hex
    if (-not $ac) { continue }
    try {
        $sl = $script:SectionLabels[$secName]
        $sp = $script:SectionPanels[$secName]
        if ($sl) {
            $sl.ForeColor = $ac
            $sl.Font = New-Object System.Drawing.Font($sl.Font, [System.Drawing.FontStyle]::Bold)
        }
        if ($sp) { $sp.BackColor = Get-BlendedColor -From $sp.BackColor -To $ac -Amount 0.10 }
    } catch { }
}
# NO RE-STACK HERE, DELIBERATELY, AND THE ATTEMPT IS WORTH RECORDING.
#
# The heading boxes have to be sized for the font the theme applies, and the
# obvious move is to re-stack once the theme has run -- Update-ConsoleSearch is
# right there and already owns this layout. Doing that broke the console: it
# left every section panel hidden and the window collapsed to 761px, and the
# zoom factor then ran away to 3.0 against a fit that wanted 1.2.
#
# Update-ConsoleSearch is not a pure layout pass. It also decides Visible for
# every button, and running it this early -- before Update-CapabilityState, in
# a half-initialised state -- makes those decisions against the wrong inputs.
#
# Set-SectionHeadBox measures the face the theme WILL use instead, so the build
# loop gets the final answer first time and nothing needs re-stacking at all.

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

#region Console zoom ----------------------------------------------------------
<#
    THE CONSOLE ZOOMS BY REFLOWING, NOT BY UNIFORM SCALING.

    Enable-FormAutoScale (Common.ps1) gives every tool window a single scale
    factor -- the SMALLER of the width and height ratios, so proportions are
    preserved exactly. That is right for a tool window and wrong for this one.

    WHY IT IS WRONG HERE, WITH THE ACTUAL NUMBERS. The console is designed at
    roughly 1010x954. Maximised on a 1845x1204 desktop that is a width ratio of
    1.83 and a height ratio of 1.26, so uniform scaling takes 1.26 -- and the
    height is then FULLY consumed while 550px of width sits empty on the right.
    Uniform zoom is always capped by the shorter axis, and this window is tall
    on monitors that are wide. No choice of factor fixes that; the LAYOUT has to
    change.

    So it reflows. The section panels are FlowLayoutPanels, which already wrap
    on their own -- widen one and it fits more buttons per row without being
    asked. More columns means fewer rows, fewer rows means a shorter console,
    and a shorter console raises the ceiling on the zoom factor. The width gets
    spent twice: once on columns, once on the larger factor that shortens the
    grid allows.

    THE FACTOR IS SEARCHED, NOT CALCULATED. Columns depend on the factor (bigger
    buttons, fewer fit) and the height depends on the columns, so there is no
    closed form -- it is a fixed point. Walking candidate factors from largest to
    smallest and taking the first that fits is a handful of arithmetic per
    resize and is obviously correct, which a clever closed form would not be.

    THE FINAL LAYOUT IS MEASURED, NOT PREDICTED. The search predicts row counts
    to choose a factor; the apply pass then asks each panel how tall it actually
    became and positions from that. Prediction picks the candidate, measurement
    places the controls -- so a caption that wraps differently than the
    arithmetic assumed cannot push the last section off the bottom.
#>

$script:Zoom = $null

function Initialize-ConsoleZoom {
    <#
        Record the design layout once, on Shown.

        Shown rather than now, because Add-ToolHelpButton can still widen the
        form after the tool has finished laying it out, and a baseline captured
        before that pins the console to a width it no longer has.
    #>
    if ($script:Zoom) { return }

    $z = @{
        BtnW0   = $script:BtnW
        BtnH0   = $script:BtnH
        Pad0    = $script:BtnPad
        StartY0 = $script:SectionStartY
        DesignW = $form.ClientSize.Width
        DesignH = $form.ClientSize.Height
        Tops    = New-Object System.Collections.ArrayList
        Fonts   = New-Object System.Collections.ArrayList
        Applied = 1.0
        Busy    = $false
    }

    # Every control that is NOT a section panel or a section label. These are the
    # header, the input row, the output box and the version line -- the fixed
    # furniture above and below the grid.
    $panels = @($script:SectionPanels.Values)
    $labels = @($script:SectionLabels.Values)
    foreach ($c in $form.Controls) {
        if ($panels -contains $c) { continue }
        if ($labels -contains $c) { continue }
        # A docked control is placed by Dock, and SETTING Anchor below would
        # silently clear Dock -- which is how the status strip ended up
        # anchored to a spot below the fold instead of docked to the bottom.
        if ($c.Dock -ne [System.Windows.Forms.DockStyle]::None) { continue }
        # RIGHT-EDGE FURNITURE TRACKS THE RIGHT EDGE.
        #
        # The '?' and the logo sit against the right margin. Scaling their X from
        # the baseline would march them toward the middle as the window widens,
        # because their position is a distance from the RIGHT, not the left.
        #
        # THE WIDTH TEST IS NOT OPTIONAL, AND LEAVING IT OUT IS WHAT BROKE THE
        # MAXIMISED WINDOW.
        #
        # A small gap to the right edge does not mean "anchored right" -- it also
        # describes anything that SPANS the window. txtInput, lblResult, txtOut
        # and lblVersion are all 12px from both edges, so a gap test alone
        # flagged every one of them. They were then positioned as though they
        # were the '?' button: x = ClientWidth - scaledWidth - gap, which put a
        # 1430px-wide box at x=398, and then widened to the full panel width from
        # there. The result runs a quarter of the way off the right edge, raises
        # a horizontal scrollbar, and drags the whole console sideways.
        #
        # Real right-edge furniture is SMALL. Nothing that spans half the window
        # is decoration in the corner.
        $gap = $z.DesignW - ($c.Left + $c.Width)
        [void]$z.Tops.Add([pscustomobject]@{
            Control    = $c
            Bounds     = $c.Bounds
            # The 60px rule catches the '?' and the logo. Anything that sits
            # LEFT of a right-anchored neighbour (the match count beside
            # Arrange) has a bigger gap and used to scale from the left --
            # drifting right, at any factor above 1.0, under the button it
            # was meant to sit beside. Those say so with Tag 'RightAnchored'.
            RightAnchored = (($gap -ge 0 -and $gap -lt 60 -and $c.Width -lt ($z.DesignW / 2)) -or
                             ("$($c.Tag)" -like '*RightAnchored*'))
            RightGap   = $gap
        })

        # Anchors off at capture time, for the same reason as in
        # Initialize-FormScaleBaseline: WinForms must stop moving these controls
        # BEFORE the first resize, so that Update-ConsoleZoom is the only thing
        # positioning them. Clearing anchors during the resize instead leaves a
        # layout pass to reconcile and the controls come back stretched.
        try {
            $c.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                        [System.Windows.Forms.AnchorStyles]::Left
        } catch { }
    }

    # Fonts are captured for EVERY control including the buttons inside panels,
    # which the Tops list deliberately excludes.
    $stack = New-Object System.Collections.Stack
    $stack.Push($form)
    while ($stack.Count -gt 0) {
        $p = $stack.Pop()
        foreach ($c in $p.Controls) {
            [void]$z.Fonts.Add([pscustomobject]@{
                Control = $c
                Size    = $c.Font.Size
                Fam     = $c.Font.FontFamily
                Style   = $c.Font.Style
                Unit    = $c.Font.Unit
                Made    = $null
            })
            if ($c.Controls.Count -gt 0) { $stack.Push($c) }
        }
    }

    $script:Zoom = $z
}

function Get-ConsoleZoomFit {
    <#
        The largest factor at which the whole console still fits the window.

        Coarse steps on purpose. A finer search buys sub-pixel differences
        nobody can see and multiplies the arithmetic on every drag.
    #>
    $z = $script:Zoom
    if (-not $z) { return 1.0 }

    $availW = $form.ClientSize.Width  - 24
    $availH = $form.ClientSize.Height - $(if ($script:StatusStripH) { $script:StatusStripH } else { 0 })
    if ($availW -le 0 -or $availH -le 0) { return 1.0 }

    # Visible button count per section, taken once -- it does not vary with the
    # candidate factor, only the column count does.
    $counts = @{}
    foreach ($name in $script:SectionNames) {
        $panel = $script:SectionPanels[$name]
        if (-not $panel) { continue }
        $n = 0
        foreach ($c in $panel.Controls) { if ($c.Tag -and $c.Visible) { $n++ } }
        $counts[$name] = $n
    }

    # THIS ANSWERS ONE QUESTION ONLY: HOW TALL MAY A ROW BE?
    #
    # Width is not part of it any more. The column count is fixed at PerRow
    # and the buttons are stretched to fill the row exactly (see
    # Update-ConsoleZoom), so the width always fits by construction -- there
    # is nothing to search for. What is left to decide is the vertical scale:
    # row height, padding and type size, all of which are bounded by the
    # window's HEIGHT and not its width.
    #
    # Separating the two is the fix for both halves of one complaint on
    # 2026-09-12: maximised, the grid sat in the middle with dead space each
    # side, because a single factor could not grow past what the height
    # allowed; made narrow, it fell back to seven columns, because the factor
    # bottomed out at 1.0 and the panel stayed wider than the window.
    #
    # THE FLOOR IS 0.6, NOT 1.0. A window narrower than the design has to
    # shrink the buttons -- the alternative is dropping a column, which is the
    # behaviour that was reported as a bug.
    $cols = [int]$script:PerRow
    if ($cols -lt 1) { $cols = 1 }

    $best = 0.6
    for ($f = 3.0; $f -ge 0.6; $f -= 0.05) {
        $cell = ($z.BtnW0 + ($z.Pad0 * 2)) * $f
        if ($cell -le 0) { continue }

        # THE HEADER DOES NOT SHRINK, so the budget it takes is fixed. See the
        # $fTop note in Update-ConsoleZoom for why -- in short, fonts stop
        # getting smaller around 8pt while bounds keep scaling linearly, so a
        # shrunk header is a header whose text no longer fits it.
        #
        # AND IT BARELY COSTS ANYTHING. The header is 164px of a ~1170px stack.
        # On the 844px-high session this was reported from, letting it shrink to
        # 0.72 saved 46px of the 326px that had to be found -- 14 % of the
        # saving, for all of the legibility. The grid is where the height is.
        # AND IT IS WHATEVER THE HEADER MEASURED, not 164. Set-ConsoleFindRowLayout
        # publishes $script:HeaderBottom from the real font heights; on a 192 dpi
        # session that is nearly double the design figure, and reserving 164
        # there is what let the grid start underneath the header.
        $startY = [Math]::Max($z.StartY0, $z.StartY0 * $f)
        if ($script:HeaderBottom -and $script:HeaderBottom -gt $startY) { $startY = $script:HeaderBottom }
        $need = $startY
        foreach ($name in $script:SectionNames) {
            if (-not $counts.ContainsKey($name)) { continue }
            if ($counts[$name] -le 0) { continue }
            $rows = [int][Math]::Ceiling($counts[$name] / [double]$cols)
            # SecHeadRes, not 22: the heading is now sized from its own text and
            # is taller than the old fixed box.
            #
            # PLUS A FEW PIXELS, BECAUSE SCALING A FONT DOES NOT SCALE ITS PIXEL
            # HEIGHT LINEARLY. This reserve is design-space * $f, but the
            # stacking pass MEASURES the scaled face, and a point size times 1.3
            # does not land on 1.3x the pixels -- it rounds up. Measured here:
            # reserved 44*1.3 = 57, actual 61. Four pixels a section, six
            # sections, and the version footer ended up under the status strip.
            # It is the same trap as the $fTop note below: fonts and bounds do
            # not scale in step. Over-reserving costs a hair of zoom; under-
            # reserving costs the bottom of the window.
            $need += [Math]::Ceiling($script:SecHeadRes * $f) + $script:SecHeadSlack +
                     (($rows * (($z.BtnH0 + ($z.Pad0 * 2)) * $f)) + (12 * $f)) + ($script:SecAfter * $f)
        }
        $need += (26 * $f) + 8

        if ($need -le $availH) { $best = $f; break }
    }
    return [Math]::Round($best, 2)
}

function Set-ConsoleInitialFit {
    <#
        OPEN WITHOUT SCROLLBARS. On a display shorter than the design stack
        the console opened clamped to the work area with a vertical
        scrollbar, whose 17px then forced a horizontal one; and because the
        status strip is docked inside an AutoScroll form it scrolled away
        with the content. Dragging the window wider fixed all of it: the
        grid reflows to more columns and the stack gets shorter.

        This does that drag, at Shown, one column at a time, and MEASURES
        rather than estimates: after each widening the real reflow runs and
        the form is asked whether it still needs to scroll. A display that
        fits the design as built never enters the loop.
    #>
    $z = $script:Zoom
    if (-not $z) { return }
    try {
        $work = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea
        $cell = $z.BtnW0 + ($z.Pad0 * 2)
        if ($cell -le 0) { return }
        # Measured from the controls themselves, not from the scrollbar
        # state: ScrollableControl only decides on its bars during its own
        # layout pass, which at Shown may not have run yet, and a stale
        # "no scrollbar" here is exactly the wrong answer.
        $needsScroll = {
            $form.PerformLayout()
            try { $form.AutoScrollPosition = New-Object System.Drawing.Point(0, 0) } catch { }
            $bottom = 0
            foreach ($c in $form.Controls) {
                if (-not $c.Visible -or $c.Dock -ne [System.Windows.Forms.DockStyle]::None) { continue }
                if ($c.Bottom -gt $bottom) { $bottom = $c.Bottom }
            }
            $strip = $(if ($script:StatusStripH) { $script:StatusStripH } else { 0 })
            return (($bottom + 8 + $strip) -gt $form.ClientSize.Height) -or $form.VerticalScroll.Visible
        }
        # Fill the height first: a taller window may be all it takes.
        # A size the operator saved is left alone ONLY while it shows
        # everything; the moment it would need a scrollbar it is grown like
        # any other, because "opens with scrollbars" is the one outcome this
        # function exists to prevent.
        if ($form.Height -lt $work.Height -and (-not $script:WindowRestored -or (& $needsScroll))) {
            $form.Height = $work.Height
            $form.Top    = $work.Top
            Update-ConsoleZoom
        }
        $steps = 0
        while ((& $needsScroll) -and $steps -lt 10) {
            $steps++
            $newW = $form.Width + $cell
            if ($newW -gt $work.Width) { $form.Width = $work.Width; Update-ConsoleZoom; break }
            $form.Width = $newW
            Update-ConsoleZoom
        }
        if ($form.Right -gt $work.Right) { $form.Left = [Math]::Max($work.Left, $work.Right - $form.Width) }
        # GIVE BACK ANY HEIGHT THE REFLOW FREED, always -- including for a
        # window restored from saved bounds.
        #
        # This used to skip a restored window on the grounds that its size was
        # the operator's choice. That is right about WIDTH and position and
        # wrong about height: the saved height came from a session with a
        # different number of rows in it, and restoring it left a band of dead
        # space under the version line that nobody asked for and nobody could
        # get rid of except by resizing by hand every launch. Reported
        # 2026-09-12. Height follows the content; width and position stay
        # where they were put.
        try {
            if (-not (& $needsScroll)) {
                $bottom = 0
                foreach ($c in $form.Controls) {
                    if (-not $c.Visible -or $c.Dock -ne [System.Windows.Forms.DockStyle]::None) { continue }
                    if ($c.Bottom -gt $bottom) { $bottom = $c.Bottom }
                }
                $strip = $(if ($script:StatusStripH) { $script:StatusStripH } else { 0 })
                $used = $bottom + 8 + $strip
                if ($used -gt 200 -and $used -lt $form.ClientSize.Height) {
                    $form.ClientSize = New-Object System.Drawing.Size($form.ClientSize.Width, [int]$used)
                }
            }
        } catch { }
        Set-ConsoleStatusLayout
        Write-StartupBreadcrumb -Text ('initial fit: {0} widening step(s), client {1}x{2}, scroll={3}' -f $steps, $form.ClientSize.Width, $form.ClientSize.Height, (& $needsScroll))
        # Arrange is the top-right corner's most crowded neighbour. Whatever
        # lands on it after a reflow goes BEHIND it, and is named in the log
        # so the layout can be fixed rather than papered over.
        try {
            $chkArrange.BringToFront()
            $btnView.BringToFront()
            foreach ($c in $form.Controls) {
                if ($c -eq $chkArrange -or -not $c.Visible) { continue }
                if ($c.Bounds.IntersectsWith($chkArrange.Bounds)) {
                    $txt = "$($c.Text)"; if ($txt.Length -gt 30) { $txt = $txt.Substring(0, 30) }
                    Write-StartupBreadcrumb -Text ('overlaps Arrange: {0} "{1}" at {2} (Arrange at {3})' -f $c.GetType().Name, $txt, $c.Bounds, $chkArrange.Bounds)
                }
            }
        } catch { }
    } catch {
        Write-StartupBreadcrumb -Text "initial fit failed: $($_.Exception.Message)"
    }
}

function Update-ConsoleZoom {
    <#
        Apply the current best factor: resize the furniture, widen the panels,
        let them reflow, then place everything from the heights they report.

        Runs in full every time rather than short-circuiting when the factor has
        not changed. Dragging the window wider at the SAME factor still has to
        move the '?' and re-wrap the grid, and an early exit on an unchanged
        factor is exactly how the help button gets left behind.
    #>
    $z = $script:Zoom
    if (-not $z -or $z.Busy) { return }
    if ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Minimized) { return }

    $z.Busy = $true

    # MEASURE WITH THE SCROLLBARS OUT OF THE WAY.
    #
    # $form.AutoScroll is $true, and a scrollbar takes ~17px out of ClientSize
    # the moment it appears. Laying out against a client area that is shrinking
    # underneath you is a feedback loop: panels sized to the pre-scrollbar width
    # overflow, which raises a horizontal scrollbar, which steals height, which
    # can raise the vertical one. That is the second half of what the
    # not-maximised window was showing.
    #
    # Off for the duration means ClientSize is stable and every number below is
    # measured against the real window. Switching it back on at the end lets
    # WinForms decide honestly whether a bar is needed against the finished
    # layout -- which is the only point at which that question has a real answer.
    # ORDER IS LOAD-BEARING: PARK THE SCROLL, THEN TURN SCROLLING OFF.
    #
    # The first version did these the other way round and made things worse.
    # Setting AutoScrollPosition on a form whose AutoScroll is already $false is
    # a no-op -- the property has nothing to act on -- so the offset from the
    # previous view survived, and every subsequent Location was written into a
    # still-shifted coordinate space. That is the console sliding off the left
    # edge with 'Look Up' and 'ServiceNow Queue' half gone.
    #
    # Scrolling to the origin WHILE scrolling still works actually moves the
    # view. Only then is it safe to switch off.
    $hadScroll = $form.AutoScroll
    if ($hadScroll) {
        $form.AutoScrollPosition = New-Object System.Drawing.Point(0, 0)
    }
    $form.AutoScroll = $false

    $f = Get-ConsoleZoomFit
    $form.SuspendLayout()
    try {
        # ---- fonts ------------------------------------------------------
        #
        # NOTHING IS DISPOSED HERE. See the matching note in Update-FormScale.
        # A control that never had its Font set explicitly shares its parent's
        # Font INSTANCE, so disposing "the one we made last time" hands every
        # child a dead GDI+ handle. The next measurement throws "Parameter is
        # not valid" out of Label.SetBoundsCore and the window paints as red X
        # boxes. Let the finalizers do it.
        # TYPE GROWS WITH THE WINDOW'S WIDTH TOO, within reason. $f alone is
        # the vertical fit, and on a wide-but-short window that leaves the
        # captions small in buttons that have plenty of room. The width factor
        # is allowed to lift it, capped at a fifth above the vertical fit so
        # the text cannot outgrow the row it sits in.
        $fWide = 1.0
        try {
            $cell0 = $z.BtnW0 + ($z.Pad0 * 2)
            if ($cell0 -gt 0) {
                $perRowNow = [Math]::Max(1, [int]$script:PerRow)
                $fWide = ($form.ClientSize.Width - 24) / [double]($cell0 * $perRowNow)
            }
        } catch { }
        $fFont = [Math]::Min([Math]::Max($fWide, $f), $f * 1.2)
        if ($fFont -lt 0.6) { $fFont = 0.6 }

        # THE HEADER FURNITURE NEVER SCALES DOWN. $fTop is the factor for
        # everything in $z.Tops -- the header, the two input rows, the result
        # line, the version line.
        #
        # WHY, MEASURED. Reported 2026-09-14: one operator's header was crushed
        # -- captions cut mid-word, the result line sliced through the middle of
        # its glyphs -- while his button grid was perfect. Same server as a
        # colleague whose console was flawless, same 96 dpi, same 8.25pt system
        # font, no compatibility overrides. The only difference was the session:
        # 1716x844 against 2268x1131.
        #
        # At 844px the stack does not fit, so the factor search engaged for him
        # and not for her. And at f = 0.70:
        #
        #     8.25pt scaled to 7.13pt still renders at 13px. Identical.
        #     An 18px label row scaled to 0.70 is 13px.
        #
        # FONTS STOP SHRINKING AND BOUNDS DO NOT. Below roughly 8pt the pixel
        # metrics round to the same value, so the text kept 100 % of its height
        # while its box lost 30 %. The grid survived because a 28px button has
        # 65 % headroom; the 18px label rows had 72 % and were squeezed to
        # exactly 100 %, which is what a sliced glyph is.
        #
        # The old comment above claimed the 1.2 cap meant "the text cannot
        # outgrow the row it sits in". That is true only while the row is
        # scaling with it. It stops being true the moment the font hits its
        # floor, and nothing said so.
        $fTop = [Math]::Max($f, 1.0)
        $fFontTop = [Math]::Max($fFont, 1.0)
        $topControls = @($z.Tops | ForEach-Object { $_.Control })

        foreach ($it in $z.Fonts) {
            $c = $it.Control
            if ($null -eq $c -or $c.IsDisposed) { continue }
            try {
                # Header furniture gets $fFontTop, everything else $fFont.
                $use = $(if ($topControls -contains $c) { $fFontTop } else { $fFont })
                $pt = [float]($it.Size * $use)
                if ($pt -gt 0 -and $pt -le 96) {
                    $c.Font = New-Object System.Drawing.Font(
                        $it.Fam, $pt, $it.Style, $it.Unit)
                }
            } catch { }
        }

        $panelW = $form.ClientSize.Width - 24
        if ($panelW -lt 200) { $panelW = 200 }

        # ---- fixed furniture --------------------------------------------
        foreach ($it in $z.Tops) {
            $c = $it.Control
            if ($null -eq $c -or $c.IsDisposed) { continue }
            $c.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor `
                        [System.Windows.Forms.AnchorStyles]::Left
            # $fTop, NOT $f -- the header keeps its design size when the grid
            # has to shrink. See the note where $fTop is worked out. The Y
            # positions have to use the same factor or the rows would close up
            # on each other while keeping their heights.
            $w = [int][Math]::Round($it.Bounds.Width  * $fTop)
            $h = [int][Math]::Round($it.Bounds.Height * $fTop)
            $y = [int][Math]::Round($it.Bounds.Y * $fTop)

            # An AutoSize label -- lblPim and lblInput are both AutoSize -- sets
            # its own width and height from the text. Writing them forces a
            # GetPreferredSize pass, and that is the exact call in the crash
            # stack: Label.SetBoundsCore -> GetPreferredSize -> MeasureString.
            # Give it a position and let it size itself.
            $auto = $false
            try { $auto = [bool]$c.AutoSize } catch { }
            if ($auto) { $w = $c.Width; $h = $c.Height }

            if ($it.RightAnchored) {
                $x = $form.ClientSize.Width - $w - [int][Math]::Round($it.RightGap * $fTop)
            } else {
                $x = [int][Math]::Round($it.Bounds.X * $fTop)
            }

            if ($auto) { $c.Location = New-Object System.Drawing.Point($x, $y) }
            else       { $c.SetBounds($x, $y, $w, $h) }
        }

        # The controls that span the grid have to match its width exactly, or
        # the console looks lopsided -- the same complaint that
        # Update-SectionPanelWidth exists to prevent.
        #
        # $txtInput was removed from this list on 2026-09-13. It is now the same
        # width as the find box rather than the width of the grid, and
        # Set-ConsoleFindRowLayout -- which runs immediately below -- is the one
        # place that decides what that is.
        foreach ($c in @($lblResult, $txtOut, $lblVersion)) {
            if ($c -and -not $c.IsDisposed) { $c.Width = $panelW }
        }

        # LAST WORD ON THE FIND ROW. The generic right-anchor pass above has
        # already placed Arrange, View and the count label from their captured
        # gaps; this re-places all four from their CURRENT widths, which is the
        # only version that survives the theme having resized them. Running
        # after the generic pass rather than instead of it keeps the zoom code
        # free of special cases.
        try { Set-ConsoleFindRowLayout } catch { }

        # The header caption was re-placed here too, because the font pass above
        # changes the title's width on every zoom. Removed with the caption on
        # 2026-09-13.

        # ---- the grid ----------------------------------------------------
        # HEIGHT FROM THE VERTICAL FIT, WIDTH FROM THE WINDOW.
        #
        # The row height, the padding and the type size come from $f, which is
        # what the window's HEIGHT allows. The button WIDTH is then whatever
        # divides the available width into exactly PerRow cells -- so the grid
        # fills the row edge to edge at any window size, and the space that
        # used to sit unused either side of a maximised console goes into the
        # buttons instead.
        #
        # Derived in this order because the cell is the button plus its two
        # margins: the padding has to be known before the width that is left
        # for the button can be worked out.
        $script:BtnH   = [int][Math]::Round($z.BtnH0 * $f)
        $script:BtnPad = [int][Math]::Round($z.Pad0  * $f)
        if ($script:BtnPad -lt 1) { $script:BtnPad = 1 }

        $cols = [int]$script:PerRow
        if ($cols -lt 1) { $cols = 1 }
        # The panel's own border and the slack Get-ConsolePanelWidth adds are
        # taken off first, or the last column would be a few pixels wide and
        # wrap.
        $usable = $panelW - 2 - $script:ConsolePanelSlack
        $cellW  = [int][Math]::Floor($usable / [double]$cols)
        $script:BtnW = $cellW - ($script:BtnPad * 2)
        # A floor, because a button narrower than this cannot show a caption
        # at all. Below it the grid is allowed to overflow and the window's
        # own scrollbar takes over -- visible, and better than invisible
        # buttons.
        if ($script:BtnW -lt 60) { $script:BtnW = 60 }
        $script:PanelW = $panelW

        $btnSize = New-Object System.Drawing.Size($script:BtnW, $script:BtnH)
        $margin  = New-Object System.Windows.Forms.Padding($script:BtnPad)
        foreach ($panel in $script:SectionPanels.Values) {
            if (-not $panel) { continue }
            $panel.Width = $panelW
            foreach ($c in $panel.Controls) {
                if (-not $c.Tag) { continue }
                $c.Size   = $btnSize
                $c.Margin = $margin
            }
            # AFTER the buttons are resized: the padding is computed from the
            # new cell size, so it has to be the last thing done to the panel.
            Set-SectionPanelGrid -Panel $panel
        }

        # ---- the header, BEFORE the grid is placed -------------------------
        # It has to run first now: it measures its own text, re-stacks its rows
        # and publishes $script:HeaderBottom, which is what the grid starts
        # below. Called again at the end of this function for the right-hand
        # end of the find row; running it twice is cheap and idempotent.
        try { Set-ConsoleFindRowLayout } catch { }

        # ---- place, from what the panels ACTUALLY became -------------------
        # THE GRID STARTS BELOW THE MEASURED HEADER. $fTop keeps the design
        # figure from shrinking; HeaderBottom overrides it whenever the real
        # text needs more, which on a high-dpi session it always does.
        # Get-ConsoleZoomFit reserves the same number when it picks the factor,
        # so the two cannot disagree.
        $y = [int][Math]::Round($z.StartY0 * $fTop)
        if ($script:HeaderBottom -and $script:HeaderBottom -gt $y) { $y = [int]$script:HeaderBottom }
        foreach ($name in $script:SectionNames) {
            $panel = $script:SectionPanels[$name]
            $lbl   = $script:SectionLabels[$name]
            if (-not $panel) { continue }
            # Folded: the heading is placed, the panel is not (it is hidden
            # already; Update-ConsoleSearch owns that). Same rule as unzoomed.
            if ($lbl -and $lbl.Visible -and -not $panel.Visible -and $script:CollapsedSections[$name]) {
                $lbl.Location = New-Object System.Drawing.Point([int][Math]::Round(12 * $f), $y)
                $y += (Get-SectionHeadOffset -Label $lbl -Factor $f) + [int][Math]::Round(6 * $f)
                continue
            }
            if (-not $panel.Visible) { continue }

            $panel.PerformLayout()
            $rowH = $script:BtnH + ($script:BtnPad * 2)
            $rows = 1
            foreach ($c in $panel.Controls) {
                if (-not $c.Tag -or -not $c.Visible) { continue }
                $r = [int][Math]::Floor($c.Top / [double][Math]::Max(1, $rowH)) + 1
                if ($r -gt $rows) { $rows = $r }
            }
            $panelH = ($rows * $rowH) + [int][Math]::Round(12 * $f)

            # THE HEADING BOX IS RE-SIZED HERE, and this is the call that fixes
            # the zoomed case. This function scaled the heading FONT by $f and
            # scaled the old 22px offset with it, but left the label's Height at
            # the 20 the build loop set -- so the bigger the zoom, the more of
            # every heading was cut off.
            $headOff = Get-SectionHeadOffset -Label $lbl -Factor $f
            if ($lbl) { $lbl.Location = New-Object System.Drawing.Point([int][Math]::Round(12 * $f), $y) }
            $panel.Location = New-Object System.Drawing.Point(
                [int][Math]::Round(12 * $f), ($y + $headOff))
            $panel.Height = $panelH
            $y += $headOff + $panelH + [int][Math]::Round($script:SecAfter * $f)
        }

        if ($lblVersion -and -not $lblVersion.IsDisposed) {
            $lblVersion.Location = New-Object System.Drawing.Point(
                [int][Math]::Round(12 * $f), ($y + [int][Math]::Round(4 * $f)))
            # AFTER the Tops pass has scaled its bounds by $fTop, or that pass
            # would put the too-small height straight back.
            Expand-LabelToFont -Label $lblVersion
        }

        $z.Applied = $f
        try { $chkArrange.BringToFront(); $btnView.BringToFront() } catch { }
    } finally {
        # AutoScroll goes back on in the finally, not after the try. If anything
        # above throws, a console left permanently unable to scroll would hide
        # whatever the exception left off-screen -- and the operator would be
        # looking at a broken window with no way to reach the bottom of it.
        $form.AutoScroll = $hadScroll
        $form.ResumeLayout($true)

        # Park at the origin again once scrolling is back and the layout is
        # final. Re-enabling AutoScroll can restore a remembered offset, and the
        # top-left is where the operator expects to be looking after a resize --
        # not two hundred pixels into the middle of their own console.
        if ($hadScroll) {
            try { $form.AutoScrollPosition = New-Object System.Drawing.Point(0, 0) } catch { }
        }
        $z.Busy = $false
    }
}
#endregion

Set-FormScreenCorner -Form $form -Corner 'TopLeft'

$script:PendingConnect = $null
# The console's own '?', in the same corner as every tool's. x = ClientSize-40
# lands at 784..812, which is clear -- lblTenant ends at 660.
# OPT OUT OF THE GENERIC SCALER, BEFORE the help button wires it up.
#
# Enable-FormAutoScale reads this marker and returns. The console gets zoom from
# Update-ConsoleZoom instead, which reflows the grid rather than scaling it --
# see the region above for why uniform scaling leaves half this window empty.
#
# Tag is a plain string on this form. Add-ReportFileControls puts an ArrayList
# on Tag for report windows and would be broken by a string; the console is not
# a report window and does not call it.
$form.Tag = "$($form.Tag)|NoAutoScale|NoStatusBar"

Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Admin Console' | Out-Null

# ---------------------------------------------------------------------------
# BOTTOM STATUS STRIP. What Admin Console 2 had along its bottom edge, in
# WinForms: on the left the words / marquee / elapsed that Set-ToolStatus
# drives (Look Up and tool launches narrate through it); on the right one
# pill per connection -- ExO, SPO, EMS, AD -- then the tenant and the
# operator. The pills MIRROR the header labels: every probe that already
# updates $lblExo and friends updates the pill through TextChanged /
# ForeColorChanged / VisibleChanged, so there is exactly one place that
# decides what "connected" means and this strip never probes anything.
#
# Docked, so zoom leaves it alone (Update-ConsoleZoom skips docked
# controls) and AutoScroll cannot scroll it away; the form grows by its
# height so it covers nothing, and the zoom's available height subtracts it.
# ---------------------------------------------------------------------------
$script:StatusStripH = 26
# THE STRIP'S STATE IS GLOBAL, AND THAT IS NOT LAZINESS.
#
# Everything below reacts to EVENTS -- TextChanged, ForeColorChanged,
# VisibleChanged, Resize -- and the 60-second PIM tick raises all three of the
# first ones. An event handler is not a closure: PowerShell resolves its names
# when the event FIRES, and $script: means "the nearest SCRIPT scope on the call
# stack right now", not "the scope this line was written in".
#
# While a tool's ShowDialog() pumps messages on this thread, the nearest script
# scope is THE TOOL's. $script:StatusPills there is $null, so the handler ran
#     $null['lblPim']
# and the operator got a JIT-debugging dialog once a minute:
#     Exception setting "Visible": "Cannot index into a null array."
# raised from Timer.OnTick -> Update-PimLabel -> Label.OnVisibleChanged.
#
# This is the SAME fault the PIM timer had with $lblPim (see Update-PimLabel);
# it just reappeared one indirection further out. $global: has exactly one
# meaning from any scope on any thread, so it is the only qualifier that
# survives being called from inside somebody else's script. The controls ride
# here too, so Set-ConsoleStatusLayout never has to look a name up either.
$global:ConsoleStatus = @{
    StripH = $script:StatusStripH
    Muted  = $null
    Pills  = @{}
    # Placed RIGHT TO LEFT in this order, so the first entry sits nearest the
    # operator name and is the first to be dropped when the strip runs out of
    # room. Tenant leads deliberately: see the note in Set-ConsoleStatusLayout.
    Order  = @('lblTenant', 'lblAd', 'lblEms', 'lblSpo', 'lblExo', 'lblPim')
    Laying = $false
}
$pnlStatus = New-Object System.Windows.Forms.Panel
$pnlStatus.Name   = 'pnlConsoleStatus'
$pnlStatus.Height = $script:StatusStripH
$pnlStatus.Dock   = [System.Windows.Forms.DockStyle]::Bottom
$pnlStatus.Tag    = 'NoTheme'
$stripBack = [System.Drawing.Color]::FromArgb(233, 236, 240)
$stripText = [System.Drawing.Color]::FromArgb(31, 35, 40)
$stripMuted = [System.Drawing.Color]::FromArgb(107, 114, 128)
$stripLine = [System.Drawing.Color]::FromArgb(217, 221, 227)
$pillBack  = [System.Drawing.Color]::FromArgb(247, 248, 250)
if ($script:Theme) {
    $c = ConvertTo-ThemeColor $script:Theme.PanelBack; if ($c) { $stripBack = $c }
    $c = ConvertTo-ThemeColor $script:Theme.FormText;  if ($c) { $stripText = $c }
    $c = ConvertTo-ThemeColor $script:Theme.GroupText; if ($c) { $stripMuted = $c }
    $c = ConvertTo-ThemeColor $script:Theme.ButtonBorder; if ($c) { $stripLine = $c }
    $c = ConvertTo-ThemeColor $script:Theme.FormBack;  if ($c) { $pillBack = $c }
}
$pnlStatus.BackColor = $stripBack
$global:ConsoleStatus.Panel = $pnlStatus
$pnlStatusLine = New-Object System.Windows.Forms.Panel
$pnlStatusLine.Height = 1; $pnlStatusLine.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlStatusLine.BackColor = $stripLine; $pnlStatusLine.Tag = 'NoTheme'
$pnlStatus.Controls.Add($pnlStatusLine)

$lblStatusText = New-Object System.Windows.Forms.Label
$lblStatusText.Name = 'lblConsoleStatus'; $lblStatusText.Tag = 'NoTheme'
$lblStatusText.Text = 'Ready'
$lblStatusText.AutoSize = $false
$lblStatusText.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblStatusText.Location = New-Object System.Drawing.Point(10, 1)
$lblStatusText.Height = $script:StatusStripH - 1
$lblStatusText.ForeColor = $stripText
$lblStatusText.BackColor = $stripBack
$pnlStatus.Controls.Add($lblStatusText)
$global:ConsoleStatus.Text = $lblStatusText

$pbConsole = New-Object System.Windows.Forms.ProgressBar
$pbConsole.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
$pbConsole.MarqueeAnimationSpeed = 25
$pbConsole.Size = New-Object System.Drawing.Size(100, 10)
$pbConsole.Visible = $false
$pbConsole.Tag = 'NoTheme'
$pnlStatus.Controls.Add($pbConsole)
$global:ConsoleStatus.Bar = $pbConsole

$lblConsoleElapsed = New-Object System.Windows.Forms.Label
$lblConsoleElapsed.Tag = 'NoTheme'
$lblConsoleElapsed.AutoSize = $false
$lblConsoleElapsed.Size = New-Object System.Drawing.Size(64, ($script:StatusStripH - 1))
$lblConsoleElapsed.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblConsoleElapsed.ForeColor = $stripMuted
$lblConsoleElapsed.BackColor = $stripBack
$pnlStatus.Controls.Add($lblConsoleElapsed)
$global:ConsoleStatus.Elapsed = $lblConsoleElapsed

function New-ConsolePill {
    # A dot + the label ITSELF, moved down from the header into the strip.
    # The probes keep writing Text / ForeColor / Visible to the same label
    # they always did; the dot follows ForeColor, the pill follows Visible,
    # and nothing is shown twice.
    param([Parameter(Mandatory)]$Source, [string]$Name)
    $Source.Name = $Name
    $pill = New-Object System.Windows.Forms.Panel
    $pill.Tag = 'NoTheme'; $pill.Height = 18; $pill.BackColor = $pillBack
    $dot = New-Object System.Windows.Forms.Panel
    $dot.Tag = 'NoTheme'; $dot.Size = New-Object System.Drawing.Size(8, 8); $dot.Location = New-Object System.Drawing.Point(8, 5)
    $dot.BackColor = $Source.ForeColor
    $dot.Visible = $Source.Visible
    if ($Source.Parent) { $Source.Parent.Controls.Remove($Source) }
    $Source.AutoSize  = $false
    $Source.Height    = 18
    $Source.Location  = New-Object System.Drawing.Point(20, 0)
    $Source.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $Source.Font      = New-Object System.Drawing.Font($form.Font.FontFamily, 8)
    $Source.BackColor = $pillBack
    $pill.Controls.Add($dot); $pill.Controls.Add($Source)
    $pnlStatus.Controls.Add($pill)
    $global:ConsoleStatus.Pills[$Name] = @{ Panel = $pill; Dot = $dot; Text = $Source }
    # ALL THREE HANDLERS READ $global: AND SWALLOW THEIR OWN FAULTS.
    #
    # They are raised by the PIM timer while a tool may be modal on this
    # thread, so $script: would land in the TOOL's scope (see the note above
    # the strip). And a decoration that fails must not become a JIT-debugging
    # dialog: a dot that is briefly the wrong colour is a cosmetic defect,
    # whereas an unhandled exception in a paint-time handler stops the
    # console. Nothing in here is load-bearing.
    $Source.Add_TextChanged({ try { Set-ConsoleStatusLayout } catch { } })
    $Source.Add_ForeColorChanged({
        # DimGray ("not checked") is unreadable on a dark strip; re-tint to
        # the theme's muted text. Green / red / the PIM colours pass through.
        try {
            $s = $global:ConsoleStatus
            if (-not $s) { return }
            if ($s.Muted -and $this.ForeColor.ToArgb() -eq [System.Drawing.Color]::DimGray.ToArgb()) { $this.ForeColor = $s.Muted; return }
            $p = $s.Pills["$($this.Name)"]
            if ($p) { $p.Dot.BackColor = $this.ForeColor }
        } catch { }
    })
    $Source.Add_VisibleChanged({
        # The pill itself is never hidden -- a label whose parent is hidden
        # cannot raise VisibleChanged again when a probe re-shows it. The
        # dot follows the label and the layout gives a hidden pill no width.
        try {
            $s = $global:ConsoleStatus
            if (-not $s) { return }
            $p = $s.Pills["$($this.Name)"]
            if ($p) { $p.Dot.Visible = $this.Visible; Set-ConsoleStatusLayout }
        } catch { }
    })
    return $pill
}
$global:ConsoleStatus.Muted = $(if ($script:Theme) { $stripMuted } else { $null })
[void](New-ConsolePill -Source $lblPim    -Name 'lblPim')
[void](New-ConsolePill -Source $lblExo    -Name 'lblExo')
[void](New-ConsolePill -Source $lblSpo    -Name 'lblSpo')
[void](New-ConsolePill -Source $lblEms    -Name 'lblEms')
[void](New-ConsolePill -Source $lblAd     -Name 'lblAd')
[void](New-ConsolePill -Source $lblTenant -Name 'lblTenant')
$lblOperator = New-Object System.Windows.Forms.Label
$lblOperator.Tag = 'NoTheme'; $lblOperator.AutoSize = $false; $lblOperator.Height = $script:StatusStripH - 1
$lblOperator.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblOperator.Font = New-Object System.Drawing.Font($form.Font.FontFamily, 8)
$lblOperator.ForeColor = $stripMuted; $lblOperator.BackColor = $stripBack
try { $lblOperator.Text = Get-LogOperator } catch { $lblOperator.Text = "$env:USERNAME" }
$lblOperator.Width = [System.Windows.Forms.TextRenderer]::MeasureText($lblOperator.Text, $lblOperator.Font).Width + 12
$pnlStatus.Controls.Add($lblOperator)
$global:ConsoleStatus.Operator = $lblOperator

function Set-ConsoleStatusLayout {
    # Right-to-left: operator, tenant, AD, EMS, SPO, ExO; the status text
    # takes what is left. Called on strip resize and whenever a pill changes.
    #
    # EVERY CONTROL COMES OUT OF $global:ConsoleStatus, NOT OFF A NAME.
    # This runs from event handlers, and an event handler can fire while a
    # tool script is the current scope -- at which point both $script:StatusPills
    # and a bare $pnlStatus mean whatever that tool happens to have called
    # them, which is usually nothing. Reading one global hashtable is the
    # only lookup that gives the same answer from every scope.
    $s = $global:ConsoleStatus
    if (-not $s -or -not $s.Panel -or $s.Panel.IsDisposed) { return }

    # RE-ENTRANCY GUARD. This is called from TextChanged, VisibleChanged and
    # Resize, and it WRITES widths and locations -- several of which raise
    # those same events. Without the guard a single connection probe lays the
    # strip out five or six times, and it is the half-finished passes that get
    # painted. A strip that looks right only after the last event is a strip
    # that looks wrong most of the time.
    if ($s.Laying) { return }
    $s.Laying = $true
    try {
        $pad   = 10
        $gap   = 6
        $dotW  = 20     # left offset of the label inside its pill
        $tail  = 10     # slack on the pill's right
        # NO PILL MAY COLLAPSE BELOW THIS, and that is the fix for the strip
        # turning into a row of coloured squares. The width used to be derived
        # purely from MeasureText of the label's CURRENT text -- so at any
        # moment a probe had blanked a label before writing the new value, the
        # pill shrank to the width of its dot and stayed there until something
        # else happened to re-run this. A floor means a transient empty string
        # costs nothing.
        $minText = 30

        $availW = $s.Panel.ClientSize.Width
        if ($availW -le 0) { return }

        # SELF-HEALING: force every pill panel visible before measuring.
        # Hiding one is a one-way door (a hidden parent makes its label report
        # Visible = $false, so the next pass skips it and never un-hides it).
        # Nothing in this function hides them any more, but asserting it here
        # costs one property write per pill and means the strip recovers from
        # anything else that ever does -- including a future edit of this file.
        foreach ($n in $s.Order) {
            $p = $s.Pills[$n]
            if ($p -and $p.Panel -and -not $p.Panel.IsDisposed -and -not $p.Panel.Visible) { $p.Panel.Visible = $true }
        }

        # ---- measure everything FIRST, place nothing yet --------------------
        #
        # Measuring and placing in one pass is what made this fragile: each
        # control was positioned from the running total of the ones before it,
        # so one bad measurement moved everything to its left. Measure, decide
        # what fits, then place.
        $s.Operator.Width = [Math]::Max(60,
            ([System.Windows.Forms.TextRenderer]::MeasureText("$($s.Operator.Text)", $s.Operator.Font).Width + 12))

        # The elapsed field is MEASURED, not fixed at its design width. It is
        # written by Set-ToolStatus, which puts anything from '' to '2257.5 s'
        # in it -- a fixed 64px was sized for the short case and clipped the
        # long one.
        $elW = 0
        if ("$($s.Elapsed.Text)".Trim()) {
            $elW = [System.Windows.Forms.TextRenderer]::MeasureText("$($s.Elapsed.Text)", $s.Elapsed.Font).Width + 10
        }
        $s.Elapsed.Width = [Math]::Max(0, $elW)
        $barW = $(if ($s.Bar.Visible) { $s.Bar.Width + 8 } else { 0 })

        # A PILL PANEL IS NEVER HIDDEN. IT IS GIVEN ZERO WIDTH.
        #
        # Control.Visible reports EFFECTIVE visibility: a label whose parent is
        # hidden returns $false no matter what its own flag says. So hiding a
        # pill panel makes its label read as invisible, which makes the next
        # pass skip it, which leaves the panel hidden -- a one-way door. Hide
        # all six once and the strip never shows a pill again, which is exactly
        # what happened the moment a tool went busy.
        #
        # The original code said so in as many words and used Width = 0. This
        # pass added Visible = $false on top of it and broke the invariant the
        # comment existed to protect. Zero width is the only mechanism here.
        $wanted = @()
        foreach ($n in $s.Order) {
            $p = $s.Pills[$n]
            if (-not $p) { continue }
            if (-not $p.Text.Visible) { $p.Panel.Width = 0; continue }
            $tw = [System.Windows.Forms.TextRenderer]::MeasureText("$($p.Text.Text)", $p.Text.Font).Width
            if ($tw -lt $minText) { $tw = $minText }
            $wanted += [PSCustomObject]@{ Name = $n; Pill = $p; W = ($dotW + $tw + $tail); TextW = ($tw + 4) }
        }

        # ---- decide what fits ----------------------------------------------
        #
        # WHEN IT DOES NOT ALL FIT, SOMETHING IS DROPPED DELIBERATELY rather
        # than everything being squeezed. Tenant goes first: it answers a
        # question ("which tenant am I in") that is asked once a session, where
        # the connection pills answer one asked constantly. Dropping the least
        # useful item whole beats shrinking six items into illegibility.
        # AND IT CAN NEVER DROP THEM ALL. The loop stops with two pills left,
        # whatever the arithmetic says. A budget that comes out wrong -- because
        # the panel was measured mid-resize, or a font changed underneath it --
        # should cost the least useful pill, not the whole strip. Losing the
        # connection state is the one outcome this row exists to prevent.
        $minTextW = 140
        $budget = $availW - $pad - $s.Operator.Width - $pad - $barW - $s.Elapsed.Width - 8 - $minTextW - $s.Text.Left
        $total = 0
        foreach ($w in $wanted) { $total += $w.W + $gap }
        while ($total -gt $budget -and @($wanted).Count -gt 2) {
            $drop = $wanted[0]                     # Order[0] is the rightmost pill
            $total = $total - ($drop.W + $gap)
            $drop.Pill.Panel.Width = 0
            $wanted = @($wanted | Select-Object -Skip 1)
        }

        # ---- place, right to left ------------------------------------------
        $x = $availW - $pad
        $s.Operator.Location = New-Object System.Drawing.Point(($x - $s.Operator.Width), 1)
        $x = $s.Operator.Left - $pad
        foreach ($w in $wanted) {
            $p = $w.Pill
            $p.Text.Width  = $w.TextW
            $p.Panel.Width = $w.W
            $p.Panel.Location = New-Object System.Drawing.Point(($x - $w.W), 4)
            $x = $p.Panel.Left - $gap
        }
        if ($s.Elapsed.Width -gt 0) {
            $s.Elapsed.Location = New-Object System.Drawing.Point(($x - 8 - $s.Elapsed.Width), 1)
            $x = $s.Elapsed.Left
        }
        if ($s.Bar.Visible) {
            $s.Bar.Location = New-Object System.Drawing.Point(($x - 8 - $s.Bar.Width), [int](($s.StripH - $s.Bar.Height) / 2))
            $x = $s.Bar.Left
        }
        $s.Text.Width = [Math]::Max(80, ($x - 10 - $s.Text.Left))
    } finally { $s.Laying = $false }
}
$pnlStatus.Add_Resize({ try { Set-ConsoleStatusLayout } catch { } })

# THE TWO EVENTS THE STRIP USED TO IGNORE, AND THE REASON IT LOOKED FINE ON A
# CLEAN LAUNCH AND WRONG AFTER ANY WORK.
#
# Set-ToolStatus does exactly two things to this strip: it toggles the
# marquee's Visible and it writes the elapsed text. On the GENERIC strip that
# Add-ToolStatusBar builds, both controls are Anchor = Top|Right, so they hold
# their own position and nothing else needs to know. The console's strip is
# not anchored -- it is positioned by Set-ConsoleStatusLayout -- and that
# function was wired to pill changes and panel resizes only.
#
# So the first time a tool went busy, a bar appeared and the elapsed field
# grew, both at stale coordinates, on top of whatever was already there. It
# stayed wrong until something unrelated happened to re-run the layout.
$pbConsole.Add_VisibleChanged({ try { Set-ConsoleStatusLayout } catch { } })
$lblConsoleElapsed.Add_TextChanged({ try { Set-ConsoleStatusLayout } catch { } })
# Long messages end in '...' rather than stopping mid-word. The label is a
# fixed width by the time the words arrive, and a sentence that simply stops
# reads as a bug in the tool that wrote it.
$lblStatusText.AutoEllipsis = $true

$form.Controls.Add($pnlStatus)
# Grow so the strip covers nothing; re-clamp to the working area as the
# original sizing did.
$form.ClientSize = New-Object System.Drawing.Size($form.ClientSize.Width, ($form.ClientSize.Height + $script:StatusStripH))
$work = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea
if ($form.Height -gt $work.Height) { $form.Height = $work.Height }
Set-ConsoleStatusLayout
Add-ToolStatusBar -Form $form -Label $lblStatusText -Bar $pbConsole -Elapsed $lblConsoleElapsed | Out-Null
if ($script:CommonLoadWatch) {
    $lblConsoleElapsed.Text = ('svc {0:0.0} s' -f ($script:CommonLoadWatch.ElapsedMilliseconds / 1000.0))
}

# ---------------------------------------------------------------------------
# HEADER BAND. A theme that names HeaderBack gets a band across the top of
# the console behind the connection labels, the way Admin Console 2 had a
# dark top bar. Labels up there keep their own text colours (green / red /
# grey say what they mean on any background) but take the band's colour as
# background, because a WinForms Label is opaque and would otherwise sit in
# a light rectangle. Themes without the key get no band and nothing moves.
# ---------------------------------------------------------------------------
$script:HeaderBandH = 50
if ($script:Theme -and "$($script:Theme.PSObject.Properties['HeaderBack'].Value)".Trim()) {
    $hb = ConvertTo-ThemeColor "$($script:Theme.HeaderBack)"
    if ($hb) {
        $pnlHeader = New-Object System.Windows.Forms.Panel
        $pnlHeader.Name = 'pnlConsoleHeader'; $pnlHeader.Tag = 'NoTheme'
        $pnlHeader.Location = New-Object System.Drawing.Point(0, 0)
        $pnlHeader.Size = New-Object System.Drawing.Size($form.ClientSize.Width, $script:HeaderBandH)
        $pnlHeader.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
        $pnlHeader.BackColor = $hb
        $form.Controls.Add($pnlHeader)
        $pnlHeader.SendToBack()
        $ht = ConvertTo-ThemeColor "$($script:Theme.PSObject.Properties['HeaderText'].Value)"
        foreach ($c in @($form.Controls)) {
            if ($c -eq $pnlHeader -or $c.Dock -ne [System.Windows.Forms.DockStyle]::None) { continue }
            if ($c.Top -ge $script:HeaderBandH) { continue }
            if ($c -is [System.Windows.Forms.Label] -or $c -is [System.Windows.Forms.PictureBox]) {
                $c.BackColor = $hb
                if ($ht -and $c -is [System.Windows.Forms.Label] -and $c.ForeColor.ToArgb() -eq [System.Drawing.Color]::DimGray.ToArgb()) { $c.ForeColor = $ht }
            }
        }
        $lblTitle.ForeColor = [System.Drawing.Color]::White
        $lblTitle.BackColor = $hb
    }
}

# Zoom hooks. Shown captures the baseline; ResizeEnd fires once when a drag
# finishes; SizeChanged catches maximise and restore, which do NOT raise
# ResizeEnd -- without it, maximising gives a big window with small text, which
# is the whole complaint.
#
# SHOWN CAPTURES THE BASELINE AND DOES NOT APPLY ANYTHING.
#
# Applying at factor 1.0 would rebuild the default view from measured panel
# heights instead of the ones Build-ToolButtons produced. Those should agree --
# but "should agree" is not "does agree", and the default view is the one thing
# about this window that is already right. Nothing touches it until the operator
# actually resizes.
$form.Add_Shown({
    # BEFORE the zoom captures its baseline, so the baseline records the find
    # row as it will actually be drawn. Captured from the stale literal
    # positions instead, every later zoom pass would faithfully reproduce the
    # overlap at a larger size.
    try { Set-ConsoleFindRowLayout } catch { }
    # Folded sections come from the operator's state file. Seeded here, once,
    # and applied through the same reflow the Find box uses, so there is one
    # way sections get placed and not two. Runs before the zoom baseline so
    # the baseline is the layout the operator will actually see.
    try {
        $script:CollapsedSections = @{}
        foreach ($n in @((Get-ConsoleState).Collapsed)) {
            if ("$n" -and $script:SectionPanels.ContainsKey("$n")) { $script:CollapsedSections["$n"] = $true }
        }
        if ($script:CollapsedSections.Count -gt 0) { Update-ConsoleSearch -Text '' }
    } catch { }
    Initialize-ConsoleZoom
    # Where the operator left it, clamped to a screen that exists. After the
    # baseline so a restored (larger) size reflows through the zoom rather
    # than being recorded as the design size.
    $script:WindowRestored = $false
    try { if (Restore-ConsoleWindowBounds) { $script:WindowRestored = $true; Update-ConsoleZoom } } catch { }
    Set-ConsoleInitialFit
    $script:ConsoleFitDone = $true
    # Once more after the first paint: AutoScroll settles its bars a beat
    # after Shown, and a fit measured before that can be off by one column.
    #
    # THE DIAGNOSTIC RIDES ON THE SAME TIMER, and after the second fit rather
    # than before it -- the numbers worth having are the ones on screen, and a
    # fit that is about to be redone is not those. -ZoomDiag only.
    $t = New-Object System.Windows.Forms.Timer
    $t.Interval = 250
    $t.Add_Tick({
        $this.Stop(); $this.Dispose()
        Set-ConsoleInitialFit
        try { Write-ConsoleZoomDiagnostic } catch { }
    })
    $t.Start()
})
$form.Add_ResizeEnd({ Update-ConsoleZoom; Set-ConsoleStatusLayout })
$form.Add_SizeChanged({
    if ($script:ZoomLastState -ne $form.WindowState) {
        $script:ZoomLastState = $form.WindowState
        Update-ConsoleZoom
        Set-ConsoleStatusLayout
    }
})
$script:ZoomLastState = $form.WindowState

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
# Position and size are remembered at close, not on every move: a drag fires
# dozens of events and a file write per event is how a preference file gets
# corrupted by the next launch reading it mid-write.
$form.Add_FormClosing({ try { Save-ConsoleWindowBounds } catch { } })

# The two breadcrumbs that localise the fault.
#
# "about to ShowDialog" with nothing after it = the window never opened, and
# everything above this line is suspect. "ShowDialog returned" a fraction of a
# second later = the window DID open and something inside the message loop tore
# it down. Those are different investigations, and without these two lines
# there is no way to tell them apart after the fact.
# ---------------------------------------------------------------------------
# LAST-DITCH HANDLER FOR THE UI THREAD.
#
# WinForms already catches exceptions thrown inside the message loop -- that
# is what produces the grey ".NET Framework / Unhandled exception has occurred
# in a component in your application" box with Continue and Details. That box
# is the worst of both worlds: it interrupts the operator, it says nothing
# useful without expanding Details, and the stack it does show scrolls away
# unread. Worse, a fault in a once-a-minute timer raises it once a minute.
#
# Attaching a handler here REPLACES that dialog. The exception and its stack
# go to the startup log where they can be pasted somewhere useful, and the
# operator is told once per DISTINCT message per session -- so a repeating
# fault is a line in a file, not a box every sixty seconds.
#
# The handler reads $global: only. It runs on whatever scope the message loop
# happens to be in, which by definition is the scope that just failed; see the
# note above the status strip on why $script: cannot be trusted from here.
#
# This is a net, not a fix. Anything that lands in ConsoleStartup.log with
# "UI THREAD" on it is still a defect and should be read as one.
# ---------------------------------------------------------------------------
#
# HOOKED ONCE PER PROCESS, not once per run. A theme change relaunches the
# console inside this same process; attaching again would give two handlers
# and two message boxes for one fault. The log path and the seen-list are
# refreshed on every run, the subscription is not.
$global:ConsoleErrorLog  = $script:StartupLogPath
$global:ConsoleErrorSeen = @{}
try {
    if (-not $global:ConsoleThreadExceptionHooked) {
    $global:ConsoleThreadExceptionHooked = $true
    [System.Windows.Forms.Application]::add_ThreadException([System.Threading.ThreadExceptionEventHandler]{
        param($sender, $e)
        try {
            $ex  = $e.Exception
            $msg = "$($ex.GetType().FullName): $($ex.Message)"
            $log = $global:ConsoleErrorLog
            if ($log) {
                $dir = Split-Path -Parent $log
                if ($dir -and -not (Test-Path -LiteralPath $dir)) {
                    New-Item -Path $dir -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
                }
                $lines = @('{0}  pid {1}  UI THREAD: {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $PID, $msg)
                foreach ($l in ("$($ex.StackTrace)" -split "`r?`n")) { if ("$l".Trim()) { $lines += "    $l" } }
                Add-Content -LiteralPath $log -Value $lines -Encoding UTF8 -ErrorAction SilentlyContinue
            }
            $seen = $global:ConsoleErrorSeen
            if ($seen -isnot [hashtable]) { return }
            if ($seen.ContainsKey($msg)) { return }
            $seen[$msg] = $true
            $body = "$msg`r`n`r`nThe console is still running. This was written to:`r`n$log`r`n`r`nIf it repeats, that file has the detail."
            $null = Show-ToolDialog($body, 'Admin Console',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
        } catch { }
    })
    }
} catch { }

# BELT FOR THE SHOWN HANDLERS. On 2026-09-12 one launch in three opened with
# scrollbars and the startup log showed no 'form Shown' line at all -- the
# Shown handlers never ran (cause not established; it followed a Connect
# relaunch in the same process). Shown is the only place the initial fit
# runs, so a launch without it is a launch at design size under a taller
# grid. This timer fires once, shortly after the window is up, and runs the
# same sequence if the Shown path has not already marked itself done. Same
# work, same order, same guards; nothing new to get wrong.
$script:ConsoleFitDone = $false
$fitGuard = New-Object System.Windows.Forms.Timer
$fitGuard.Interval = 700
$fitGuard.Add_Tick({
    $this.Stop(); $this.Dispose()
    try {
        if ($script:ConsoleFitDone) { return }
        Write-StartupBreadcrumb -Text 'Shown handlers did not run -- fitting from the guard timer'
        try { Update-SectionPanelWidth } catch { }
        try { Set-ConsoleFindRowLayout } catch { }
        Initialize-ConsoleZoom
        $script:WindowRestored = $false
        try { if (Restore-ConsoleWindowBounds) { $script:WindowRestored = $true; Update-ConsoleZoom } } catch { }
        Set-ConsoleInitialFit
        $script:ConsoleFitDone = $true
        # Same diagnostic as the Shown path, so -ZoomDiag reports either way.
        try { Write-ConsoleZoomDiagnostic } catch { }
    } catch {
        Write-StartupBreadcrumb -Text "guard timer fit failed: $($_.Exception.Message)"
    }
})
$fitGuard.Start()

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
