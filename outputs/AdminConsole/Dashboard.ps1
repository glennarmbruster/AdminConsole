#Requires -Version 5.1
# TOOLKIT-HELP-RUNS: nothing new. It reads what the console already knows and
# the logs already on this machine.
# TOOLKIT-HELP-BEGIN
# The start-of-day page. Six panels answering the questions worth asking before
# you touch anything:
#
#   ENVIRONMENT      Am I connected to what I need, and is PIM still active?
#   TOOLKIT HEALTH   Is my copy complete and current?
#   TODAY            What has run today, and did any of it fail?
#   RETRY QUEUE      Is anything sitting unfinished from a disable run?
#   TOP TOOLS        What actually gets used.
#   FAILURE HOTSPOTS What keeps going wrong, so it can be fixed.
#
# READ-ONLY. Nothing here changes a mailbox, a group or a file. The buttons
# open other tools; they do not act.
#
# The connection and version panels call the SAME functions as the console
# header and footer rather than working it out again -- so the two cannot
# disagree with each other. A second opinion about your own state is worse
# than one.
#
# ABOUT "UNKNOWN" RESULTS: rows converted from the old two-column log have no
# Result, because that log never recorded one. They are counted as Unknown, not
# folded into Success. The count shrinks on its own as those rows age out.
# TOOLKIT-HELP-END
<#
    Dashboard.ps1
    Operational status page for the admin toolkit.

    Reasoning: ScriptNotes/Dashboard.ps1.notes.md

    Author   : Glenn Armbruster
    Created  : 2026-08-07
#>

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

#region Colours and layout ------------------------------------------------------
# Same values as the console header, for the same reason: on the dark navy
# theme, stock Green and Red are dark-on-dark.
$script:ColGood = [System.Drawing.Color]::FromArgb(80, 220, 130)
$script:ColBad  = [System.Drawing.Color]::FromArgb(255, 105, 97)
$script:ColWarn = [System.Drawing.Color]::FromArgb(230, 170, 60)
$script:ColDim  = [System.Drawing.Color]::DimGray

$script:ColW  = 446
$script:ColL  = 12
$script:ColR  = 470
#endregion

#region Log reading -------------------------------------------------------------
function Get-DashLogRows {
    <#
        Every row of the completed-task log, or an empty set with a reason.

        Never throws. A dashboard that dies because a CSV is mid-write is worse
        than one that says "could not read the log" -- the first hides all five
        other panels, the second hides one.
    #>
    $out = [PSCustomObject]@{ Rows = @(); Problem = '' }

    $folder = $null
    try { $folder = Get-LogPath } catch { $out.Problem = 'Log folder not resolvable.'; return $out }
    if (-not $folder -or -not (Test-Path -LiteralPath $folder)) {
        $out.Problem = 'No Logs folder yet.'
        return $out
    }

    $file = Join-Path $folder 'CompletedTasks.csv'
    if (-not (Test-Path -LiteralPath $file)) {
        $out.Problem = 'No CompletedTasks.csv yet.'
        return $out
    }

    $head = ''
    try { $head = (Get-Content -LiteralPath $file -TotalCount 1 -ErrorAction Stop) } catch { }
    if ("$head".Trim() -ne $script:LogHeaderLine) {
        # The old two-column file. Say so precisely, and say what fixes it --
        # this is the exact state Convert-LogHistory.ps1 exists for.
        #
        # THE PATH IS IN THE MESSAGE because the script left the console's
        # button list on 2026-09-20 for 'Diagnostic tools\'. 'Run
        # Convert-LogHistory.ps1' was actionable while it was a button; without
        # the folder it now sends the reader hunting.
        $out.Problem = 'Log is still in the OLD format. Run "Diagnostic tools\Convert-LogHistory.ps1" once, from a PowerShell prompt.'
        return $out
    }

    try { $out.Rows = @(Import-Csv -LiteralPath $file -ErrorAction Stop) }
    catch { $out.Problem = "Could not read the log: $($_.Exception.Message)" }
    return $out
}

function Get-RowDate {
    param($Row)
    $d = [datetime]::MinValue
    if ([datetime]::TryParse("$($Row.Timestamp)", [ref]$d)) { return $d }
    return $null
}

function Get-RowToolName {
    <#
        Script name, or a clearly-marked stand-in.

        Rows converted from the old log have no Script -- the old format did not
        record one. They are labelled rather than dropped, because silently
        omitting them would understate every total on this page.
    #>
    param($Row)
    $s = "$($Row.Script)".Trim()
    if ($s) { return ($s -replace '\.ps1$', '') }
    return '(not recorded)'
}
#endregion

#region Section helpers ---------------------------------------------------------
function New-DashSection {
    param([string]$Title, [int]$X, [int]$Y, [int]$W, [int]$H)
    $g = New-Object System.Windows.Forms.GroupBox
    $g.Text     = $Title
    $g.Location = New-Object System.Drawing.Point($X, $Y)
    $g.Size     = New-Object System.Drawing.Size($W, $H)
    $g.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
    return $g
}

function Add-DashLine {
    <#
        One "Label : value" row inside a section. Returns the value label so the
        caller can colour it.
    #>
    param(
        [System.Windows.Forms.GroupBox]$Box,
        [string]$Key, [string]$Value, [int]$Y,
        [System.Drawing.Color]$Colour = [System.Drawing.Color]::Empty,
        [int]$KeyWidth = 130
    )
    $k = New-Object System.Windows.Forms.Label
    $k.Text     = $Key
    $k.Location = New-Object System.Drawing.Point(12, $Y)
    $k.Size     = New-Object System.Drawing.Size($KeyWidth, 18)
    $k.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
    $Box.Controls.Add($k)

    $v = New-Object System.Windows.Forms.Label
    $v.Tag      = 'NoTheme'
    $v.Text     = $Value
    $v.Location = New-Object System.Drawing.Point((12 + $KeyWidth), $Y)
    $v.Size     = New-Object System.Drawing.Size(($Box.Width - $KeyWidth - 30), 18)
    $v.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
    $v.AutoEllipsis = $true
    if ($Colour -ne [System.Drawing.Color]::Empty) { $v.ForeColor = $Colour }
    $Box.Controls.Add($v)
    return $v
}

function New-DashList {
    param([System.Windows.Forms.GroupBox]$Box, [int]$Y, [int]$H, [string[]]$Columns, [int[]]$Widths)
    $lv = New-Object System.Windows.Forms.ListView
    $lv.Tag           = 'NoTheme'
    $lv.View          = [System.Windows.Forms.View]::Details
    $lv.FullRowSelect = $true
    $lv.GridLines     = $false
    $lv.Location      = New-Object System.Drawing.Point(12, $Y)
    $lv.Size          = New-Object System.Drawing.Size(($Box.Width - 24), $H)
    $lv.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
    for ($i = 0; $i -lt $Columns.Count; $i++) {
        [void]$lv.Columns.Add($Columns[$i], $Widths[$i])
    }
    $Box.Controls.Add($lv)
    return $lv
}

function Start-ToolDetached {
    <#
        Launch another tool in a NEW process.

        Not with '&' in this one. The dashboard runs inside ShowDialog(), and a
        modal form disables every other window on its thread -- a WinForms tool
        started from here would draw and then refuse input, which is section 14d
        and cost days the first time.

        Every path is quoted. Start-Process does not quote for you, so one space
        in a folder name turns a path into two arguments (PRC001).

        ONLY FOR TOOLS THAT NEED NO SESSION. A new process starts with no
        connections, so anything that talks to Exchange, Entra or SharePoint
        opens here disconnected and useless. Tester, the log viewer and the
        config editors qualify. Disable User does not, and used to be launched
        this way -- it opened with EMS, ExO and SPO all red. Use
        $script:OpenAfterClose for those instead.
    #>
    param([string]$ScriptName)
    $path = Join-Path $PSScriptRoot $ScriptName
    if (-not (Test-Path -LiteralPath $path)) {
        Show-ToolDialog("$ScriptName is not in this folder.", 'Not found') | Out-Null
        return
    }
    # THE CONSOLE WINDOW STAYS VISIBLE. THAT IS THE DECISION, NOT AN OVERSIGHT.
    #
    # Two attempts to tidy it away on 2026-08-20, both reverted:
    #
    #   -WindowStyle Minimized      sets STARTUPINFO.wShowWindow, which Windows
    #                               applies to the process's FIRST top-level
    #                               window -- the FORM, not the console. Every
    #                               tool opened minimized into the taskbar and
    #                               every button looked dead.
    #
    #   Minimizing from the child   worked on the console, and cost the tool its
    #   (GetConsoleWindow +         foreground. Minimizing the foreground window
    #    ShowWindow)                hands activation to the next one -- the
    #                               dashboard -- and the child cannot take it
    #                               back, so the tool opened BEHIND.
    #
    # The visible console is what gives the new process the foreground, which
    # the form then inherits from its own process. It is not decoration; it is
    # the reason you can see what you launched.
    #
    # Anyone tidying this again needs an answer for activation FIRST, not as an
    # afterthought -- and needs to test it, because none of this is visible from
    # reading the code.
    #
    # -Command rather than -File, so the window can be sized before the tool
    # draws. -File takes a path and nothing else; there is nowhere to put a
    # statement that has to run first.
    #
    # No -NoExit, unlike the console's launcher: these tools are windows, not
    # transcripts, and the console should close with them rather than pile up.
    $psExe = $null
    try { $psExe = (Get-Process -Id $PID).Path } catch { }
    if (-not $psExe) { $psExe = 'powershell.exe' }

    $call = "& '" + ($path -replace "'", "''") + "'"
    $argLine = '-NoProfile -ExecutionPolicy Bypass -Command "' +
               (Get-ConsoleResizeCommand) + '; ' + $call + '"'
    try {
        Start-Process -FilePath $psExe -ArgumentList $argLine | Out-Null
    } catch {
        Show-ToolDialog("Could not launch $ScriptName`r`n$($_.Exception.Message)", 'Launch failed') | Out-Null
    }
}
#endregion

#region Form --------------------------------------------------------------------
# Set by a button, read after ShowDialog returns. A tool that needs the live
# sessions cannot be started while this form is up -- a modal form disables the
# other windows on its thread -- so it is started once this one has gone.
$script:OpenAfterClose = ''

$form = New-Object System.Windows.Forms.Form
$form.Text          = 'Admin Dashboard'
$form.Size          = New-Object System.Drawing.Size(946, 860)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

# --- Row 1 ---------------------------------------------------------------
$boxEnv    = New-DashSection -Title 'Environment'    -X $script:ColL -Y 42 -W $script:ColW -H 196
$boxHealth = New-DashSection -Title 'Toolkit Health' -X $script:ColR -Y 42 -W $script:ColW -H 196
$form.Controls.Add($boxEnv)
$form.Controls.Add($boxHealth)

# THE ENVIRONMENT PANEL IS BUILT FROM Tenant.json, not from a fixed list.
#
# It used to show all six rows always. Run it against a cloud-only tenant and it
# reported "Exchange on-prem: Not connected", "Active Directory: Not connected"
# and a PIM line -- three red or amber statements about things that do not exist
# there. Every one of them is a fault the operator cannot fix and should not be
# looking at, and a dashboard read at a glance is exactly where a permanent
# false alarm does its damage: it teaches you to ignore the red.
#
# HIDDEN, NOT GREYED. Greying says "this is off"; absence says "this is not part
# of this tenant", which is the true statement. Same rule the console follows
# for capability-blocked tools.
#
# Rows are positioned as they are added rather than at fixed offsets, so a
# cloud-only tenant gets a short panel instead of a tall one with holes in it.
$script:HasEms = Test-TenantCapability -Name 'OnPremExchange'
$script:HasAd  = Test-TenantCapability -Name 'OnPremAD'
$script:HasPim = Test-TenantCapability -Name 'Pim'
$script:HasSpo = Test-TenantCapability -Name 'SharePointOnline'
$script:HasExo = Test-TenantCapability -Name 'ExchangeOnline'

$envY = 24
$lblTenant = Add-DashLine -Box $boxEnv -Key 'Tenant' -Value '...' -Y $envY
$envY += 20

$lblExo = $null
if ($script:HasExo) {
    $lblExo = Add-DashLine -Box $boxEnv -Key 'Exchange Online' -Value '...' -Y $envY
    $envY += 20
}
$lblSpo = $null
if ($script:HasSpo) {
    $lblSpo = Add-DashLine -Box $boxEnv -Key 'SharePoint' -Value '...' -Y $envY
    $envY += 20
}
$lblEms = $null
if ($script:HasEms) {
    $lblEms = Add-DashLine -Box $boxEnv -Key 'Exchange on-prem' -Value '...' -Y $envY
    $envY += 20
}
$lblAd = $null
if ($script:HasAd) {
    $lblAd = Add-DashLine -Box $boxEnv -Key 'Active Directory' -Value '...' -Y $envY
    $envY += 20
}

$lblPim = $null
if ($script:HasPim) {
    $envY += 4
    $lblPim = Add-DashLine -Box $boxEnv -Key 'PIM' -Value '...' -Y $envY
    $envY += 24
}

$envY += 4
$lblVer = Add-DashLine -Box $boxEnv -Key 'Version' -Value '...' -Y $envY
$envY += 28

# The panel shrinks to what it actually holds. Toolkit Health beside it keeps
# its own height, so the row is as tall as the taller of the two.
$boxEnv.Height = [Math]::Max(120, $envY)

$lblFiles   = Add-DashLine -Box $boxHealth -Key 'Files verified' -Value '...' -Y 24
$lblDrift   = Add-DashLine -Box $boxHealth -Key 'Differences'    -Value '...' -Y 44
$lblTools   = Add-DashLine -Box $boxHealth -Key 'Tools.json'     -Value '...' -Y 68
$lblQuick   = Add-DashLine -Box $boxHealth -Key 'QuickLaunch.json' -Value '...' -Y 88
$lblSet     = Add-DashLine -Box $boxHealth -Key 'Settings.json'  -Value '...' -Y 108
$lblRel     = Add-DashLine -Box $boxHealth -Key 'Released'       -Value '...' -Y 132

# Run Tester stays here; Tool Manager moved to Configure.
#
# The split is between reading state and changing it. Tester answers the
# question this panel asks -- is the toolkit intact -- so it belongs beside the
# answer. Tool Manager changes what the console IS, which is a different
# activity and now has its own block.
$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text = 'Run Tester'
$btnTest.Location = New-Object System.Drawing.Point(12, 158)
$btnTest.Size = New-Object System.Drawing.Size(110, 26)
$btnTest.Add_Click({ Start-ToolDetached -ScriptName 'Test-Toolkit.ps1' })
$boxHealth.Controls.Add($btnTest)

# --- Row 2 ---------------------------------------------------------------
$boxToday = New-DashSection -Title "Today"        -X $script:ColL -Y 246 -W $script:ColW -H 196
$boxRetry = New-DashSection -Title 'Retry Queue'  -X $script:ColR -Y 246 -W $script:ColW -H 196
$form.Controls.Add($boxToday)
$form.Controls.Add($boxRetry)

$lblActions = Add-DashLine -Box $boxToday -Key 'Actions'   -Value '...' -Y 24
$lblOk      = Add-DashLine -Box $boxToday -Key 'Success'   -Value '...' -Y 44
$lblPart    = Add-DashLine -Box $boxToday -Key 'Partial'   -Value '...' -Y 64
$lblFail    = Add-DashLine -Box $boxToday -Key 'Failed'    -Value '...' -Y 84
$lblUnk     = Add-DashLine -Box $boxToday -Key 'Unknown'   -Value '...' -Y 104
$lblLast    = Add-DashLine -Box $boxToday -Key 'Last action' -Value '...' -Y 128
$lblLastAt  = Add-DashLine -Box $boxToday -Key 'At'        -Value '...' -Y 148

$btnLog = New-Object System.Windows.Forms.Button
$btnLog.Text = 'Log Viewer'
$btnLog.Location = New-Object System.Drawing.Point(12, 168)
$btnLog.Size = New-Object System.Drawing.Size(110, 24)
$btnLog.Add_Click({ Start-ToolDetached -ScriptName 'LogViewer.ps1' })
$boxToday.Controls.Add($btnLog)

$lblRetryCount = Add-DashLine -Box $boxRetry -Key 'Pending' -Value '...' -Y 24
$lvRetry = New-DashList -Box $boxRetry -Y 48 -H 112 `
           -Columns @('Employee', 'Failed step', 'Tries', 'Last attempt') `
           -Widths  @(130, 120, 45, 110)

$btnRetry = New-Object System.Windows.Forms.Button
$btnRetry.Text = 'Open Disable'
$btnRetry.Location = New-Object System.Drawing.Point(12, 164)
$btnRetry.Size = New-Object System.Drawing.Size(110, 24)
# Not detached: Disable User needs the on-premises AND cloud sessions, and both
# live in THIS process. Instead the dashboard closes itself and opens the tool
# afterwards, below ShowDialog -- by then no modal form holds the thread, so the
# tool's own window behaves normally and it inherits the live connections.
$btnRetry.Add_Click({
    # Asked before the dashboard vanishes, not after. A window closing
    # unannounced reads as a crash, and this one closes on purpose.
    $go = Show-ToolDialog(
        ("This closes the dashboard and opens Disable User." + [Environment]::NewLine + [Environment]::NewLine +
         "It has to: Disable User needs the console's live connections, and it " +
         "cannot open while the dashboard is in front of it." + [Environment]::NewLine + [Environment]::NewLine +
         "Closing Disable User returns you to the console, not to here."),
        'Open Disable User', 'OKCancel', 'Information')
    if ($go -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $script:OpenAfterClose = 'Disable-User.ps1'
    $form.Close()
})
$boxRetry.Controls.Add($btnRetry)

# --- Row 3 ---------------------------------------------------------------
$boxTop  = New-DashSection -Title 'Top Tools - last 30 days'       -X $script:ColL -Y 450 -W $script:ColW -H 226
$boxHot  = New-DashSection -Title 'Failure Hotspots - last 30 days' -X $script:ColR -Y 450 -W $script:ColW -H 226
$form.Controls.Add($boxTop)
$form.Controls.Add($boxHot)

$lvTop = New-DashList -Box $boxTop -Y 24 -H 150 -Columns @('Tool', 'Runs') -Widths @(300, 80)
$lvHot = New-DashList -Box $boxHot -Y 24 -H 150 -Columns @('Tool', 'Failed', 'Partial') -Widths @(230, 80, 80)

$lblTopNote = Add-DashLine -Box $boxTop -Key '' -Value '' -Y 180 -KeyWidth 0
$lblTopNote.ForeColor = $script:ColDim
$lblTopNote.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8)

$lblHotNote = Add-DashLine -Box $boxHot -Key '' -Value '' -Y 180 -KeyWidth 0
$lblHotNote.ForeColor = $script:ColDim
$lblHotNote.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 8)

# --- Footer --------------------------------------------------------------
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Tag      = 'NoTheme'
$lblStatus.Location = New-Object System.Drawing.Point(12, 686)
$lblStatus.Size     = New-Object System.Drawing.Size(700, 18)
$lblStatus.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 8)
$lblStatus.ForeColor = $script:ColDim
$lblStatus.AutoEllipsis = $true
$form.Controls.Add($lblStatus)

# --- Configure ------------------------------------------------------------
#
# The four editors that change what the console is, rather than what the tenant
# is. Three of them had no home at all before this -- they were not in
# Tools.json, so they were opened from Explorer or from memory.
#
# They belong together because they are the same activity and because none of
# them is daily work. Putting them on the console's front page would cost four
# slots to buttons pressed once a month; putting them here costs nothing and
# makes them findable.
#
# Safe to launch this way: none of the four connects to anything, so none of
# them raises a sign-in window. That is the constraint that decides where a
# tool can be launched from, and it was checked rather than assumed.
# CREATED BEFORE THE FIRST CONTROL THAT WANTS A TIP, which now means before
# the Configure row rather than after it.
#
# It used to sit further down, with a comment saying it was "ABOVE every
# Update- function that calls SetToolTip" -- true when the only callers were
# below it. Giving the Configure buttons tooltips put a caller ABOVE the
# declaration, and $null.SetToolTip() throws. That is the identical fault the
# comment was written to record, reintroduced by adding a caller rather than by
# moving the declaration, which is the way this one always comes back.
$toolTip = New-Object System.Windows.Forms.ToolTip

# THREE ROWS SINCE 2026-09-19: the eleventh tool (Clear Additions File) starts
# row 2 at y=90, the note moves under it, and the box and the window each grow
# by one row pitch (32).
#
# STILL THREE ROWS WITH TWELVE (2026-09-21, Park Scripts). Five per row means
# row 2 holds indices 10 to 14, so the twelfth sits beside Clear Additions and
# neither the box nor the window changes size. The next three are free too;
# the thirteenth after that starts row 3 and this comment needs revisiting.
$boxConfig = New-DashSection -Title 'Configure' -X $script:ColL -Y 716 -W 904 -H 162
$form.Controls.Add($boxConfig)

# ORDER IS THE ORDER THEY ARE NEEDED IN, not alphabetical and not the order
# they were built in. Tool Manager and Tool Status are the two opened weekly;
# Settings is the one a NEW engineer opens first, which is why it sits third
# rather than last.
$configTools = @(
    @{ Text = 'Tool Manager';    Script = 'ToolManager.ps1';      Tip = 'Buttons, sections and columns -- what the console shows and in what order.' },
    @{ Text = 'Tool Status';     Script = 'ToolStatus.ps1';       Tip = 'Which scripts the console actually uses, and a way to park the ones it does not.' },
    # Script is EMPTY on purpose -- see the InProcess flag below.
    @{ Text = 'Settings';        Script = '';                     InProcess = $true;
       Tip = 'YOUR paths, email and Edge profiles. Per-engineer -- this file is not shipped with a deploy.' },
    @{ Text = 'Theme Manager';   Script = 'ThemeManager.ps1';     Tip = 'Colours, and which tenant gets which palette.' },
    @{ Text = 'QuickLaunch';     Script = 'QuickLaunchTool.ps1';  Tip = 'The shortcuts that open from the console header.' },
    @{ Text = 'Tenant Manager';  Script = 'TenantManager.ps1';    Tip = 'Tenant facts: capabilities, directories, licence groups.' },
    @{ Text = 'Config Backup';   Script = 'ConfigBackup.ps1';     Tip = 'Copy Config\*.json somewhere safe. Run this BEFORE copying a new console folder over this one.' }
    @{ Text = 'Park In Place';  Script = 'ParkInPlace.ps1';      Tip = 'Hide tools from THIS tenant without moving any files. Park a whole capability, or tick individual scripts.' }
    @{ Text = 'Save / Load Layout'; Script = 'Customizations.ps1'; Tip = 'Your console arrangement in a file you keep. Save it BEFORE taking a new build -- an upgrade can replace the shared config, and this survives it.' }
    @{ Text = 'View Manager';    Script = 'ViewManager.ps1';      Tip = 'Which view the console opens into for YOU -- the button grid, the large-text list, the palette, the status board. Yours alone; nobody else changes.' }
    @{ Text = 'Clear Additions File'; Script = 'ClearAdditions.ps1'; Tip = 'Clear Config\Additions-*.json once its tools are on buttons, so Publish-Toolkit stops asking about them. Files go to Config\Backup, not the bin.' }
    @{ Text = 'Park Scripts';     Script = 'ParkList.ps1';        Tip = 'Park the tools this build has absorbed, from the list that shipped with it. Recommended ones open ticked; untick anything you want to keep. Hides buttons only -- nothing is deleted or moved.' }
)

# SEVEN ACROSS THE SAME WIDTH. 114 wide on a 126 pitch puts the last button at
# x=894 inside a 904 box, where six at 130/142 ended at 882 -- so the row gained
# a button and the window did not grow by a pixel.
#
# THE CAPTIONS DID NOT NEED SHORTENING. 'Tenant Manager' is the longest at about
# 90px in this font, so 114 still leaves room; 'Tenant Mgr' would have bought
# nothing and cost a word. Measured before abbreviating rather than after.
# TWO ROWS OF FIVE, AND THE BUTTONS GOT WIDER AGAIN.
#
# Seven at 114px fitted; the eighth did not, and 'Park In Place' was the caption
# that paid for it -- reported 2026-09-14 as unreadable. Squeezing eight across
# would have taken every caption below what 'Tenant Manager' needs, so the row
# WRAPS instead and the window grows by 34px to hold it.
#
# 170 wide on a 176 pitch: five per row ends at x=886 inside a 904 box, and the
# widest caption now has headroom rather than four pixels. A ninth tool goes
# into a row that already has room for it.
$script:CfgPerRow = 5
$script:CfgPitchX = 176
$script:CfgPitchY = 32
$cfgIndex = 0
foreach ($ct in $configTools) {
    $cfgCol = $cfgIndex % $script:CfgPerRow
    $cfgRow = [Math]::Floor($cfgIndex / $script:CfgPerRow)
    $cfgIndex++
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $ct.Text
    $b.Location = New-Object System.Drawing.Point((12 + ($cfgCol * $script:CfgPitchX)), (26 + ($cfgRow * $script:CfgPitchY)))
    $b.Size = New-Object System.Drawing.Size(170, 26)

    # SETTINGS RUNS IN THIS PROCESS. Everything else still gets its own.
    #
    # Start-ToolDetached exists because of section 14d: this form is up inside
    # ShowDialog(), a modal disables every other window on its thread, and a
    # WinForms tool started with '&' would draw and then refuse input. That is
    # still true for the other six -- each is a separate script that ends in its
    # own ShowDialog.
    #
    # Settings is no longer one of those. Show-SettingsEditor moved into
    # Common.ps1, which this file already dot-sources, so it can be called
    # directly -- and a dialog opened from inside a modal is simply a NESTED
    # modal, which WinForms handles. That is not a guess: ToolManager has called
    # this exact function from inside its own ShowDialog since it was written.
    #
    # The visible difference is a PowerShell host window that no longer appears
    # behind the dialog for the one tool people open most.
    if ($ct.InProcess) {
        $b.Add_Click({
            # Redrawn on save, because the dashboard PRINTS two of the values
            # this dialog changes -- the Settings.json line in Toolkit Health
            # and the reports path. Leaving them stale would have the page
            # disagree with the file the operator just wrote.
            #
            # Update-Dashboard is defined further down this file. That is fine:
            # a handler body is resolved when it RUNS, and by then the whole
            # script has been parsed.
            if (Show-SettingsEditor) { Update-Dashboard }
        })
    } else {
        # Tag carries the script name so the handler reads it from the control
        # that was clicked. A shared handler closing over $ct would see whatever
        # the loop variable held when the loop ENDED -- every button would
        # launch the last tool in the list.
        $b.Tag = $ct.Script
        $b.Add_Click({ Start-ToolDetached -ScriptName "$($this.Tag)" })
    }
    $toolTip.SetToolTip($b, $ct.Tip)
    $boxConfig.Controls.Add($b)
}

$lblCfgNote = New-Object System.Windows.Forms.Label
$lblCfgNote.Text = 'Config is read when the console starts. Restart the console to see changes made here.'
# Below ALL THREE rows: row 0 at y=26, row 1 at y=58, row 2 at y=90, each 26
# tall, so 124 clears them.
$lblCfgNote.Location = New-Object System.Drawing.Point(12, 124)
$lblCfgNote.Size = New-Object System.Drawing.Size(880, 20)
$lblCfgNote.ForeColor = [System.Drawing.Color]::DimGray
$boxConfig.Controls.Add($lblCfgNote)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = 'Refresh'
$btnRefresh.Location = New-Object System.Drawing.Point(796, 680)
$btnRefresh.Size = New-Object System.Drawing.Size(110, 28)
$form.Controls.Add($btnRefresh)

# --- Feature Request ------------------------------------------------------
#
# IN THE 42px BAND ABOVE ROW 1, which was empty. Nothing below it moves, and
# the window does not grow -- the help '?' is the only other thing up here and
# Add-ToolHelpButton puts it at ClientSize-40, which is x880. A 150px button at
# x12 cannot reach it.
#
# ON THE DASHBOARD RATHER THAN THE CONSOLE GRID, deliberately. A console slot
# is worth more than this: the grid is for work, and an operator presses this
# a handful of times a year. The dashboard is already the page you look at when
# something is annoying you, which is the exact moment a feature request gets
# thought of and then forgotten.
#
# TOP, NOT IN 'Configure'. Configure is for tools that change what the console
# IS, and it is a place you go on purpose. This has to be visible to somebody
# who was not looking for it.
#
# CREATED AFTER $toolTip so the tip can be set in one place with the button
# rather than at a distance -- the ordering fault this very comment block warns
# about, three lines up.
# COLOURED, AND THEREFORE TAGGED NoTheme.
#
# Set-ConsoleTheme repaints every Button it finds, so a palette would quietly
# turn both of these back into the same blue as everything else -- which is the
# one thing they must not be. 'NoTheme' is the existing opt-out for controls
# whose colour carries meaning rather than decoration; the connection pills and
# the ServiceNow highlight already use it.
#
# WHITE ON DARK GREEN AND DARK RED, not the bright pair. Bright green and bright
# red are what a themed dark console uses for connection state, and reusing them
# here would make two navigation buttons look like an alarm. These are dark
# enough to carry white text at a readable contrast and to read as buttons.
$btnFeature = New-Object System.Windows.Forms.Button
$btnFeature.Text      = 'Feature Request'
$btnFeature.Location  = New-Object System.Drawing.Point(12, 8)
$btnFeature.Size      = New-Object System.Drawing.Size(150, 26)
$btnFeature.Tag       = 'NoTheme'
$btnFeature.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnFeature.FlatAppearance.BorderSize = 0
$btnFeature.BackColor = [System.Drawing.Color]::FromArgb(21, 115, 71)
$btnFeature.ForeColor = [System.Drawing.Color]::White
$btnFeature.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)

# Start-ToolDetached, and that is checked rather than assumed. The rule in that
# function is that a detached tool gets a NEW process with no connections, so
# only tools needing no session may use it. Feature Request talks to an SMTP
# relay and nothing else -- no Exchange, no Entra, no SharePoint -- so it opens
# fully working. The console guard still passes: ToolkitConsoleSession is an
# environment variable and child processes inherit it.
$btnFeature.Add_Click({ Start-ToolDetached -ScriptName 'FeatureRequest.ps1' })
$form.Controls.Add($btnFeature)
$toolTip.SetToolTip($btnFeature, 'Ask for something the console does not do yet. A copy is sent to you.')

# --- Bug Fix --------------------------------------------------------------
#
# BESIDE FEATURE REQUEST, not under it. They are the same kind of act -- tell
# somebody about the console -- and the only decision is which one this is.
# Putting them side by side makes that decision take no time; putting them in
# different places makes the operator find the right one first.
#
# x170 leaves 8px between the two and ends at 320. The help '?' is placed at
# ClientSize-40, which is x880, so there is no collision -- the same arithmetic
# that let Feature Request sit in this band in the first place.
#
# Start-ToolDetached is safe here for the same reason: Bug Fix talks to an SMTP
# relay and nothing else, so it opens fully working in a fresh process with no
# connections.
$btnBug = New-Object System.Windows.Forms.Button
$btnBug.Text      = 'Bug Fix'
$btnBug.Location  = New-Object System.Drawing.Point(170, 8)
$btnBug.Size      = New-Object System.Drawing.Size(150, 26)
$btnBug.Tag       = 'NoTheme'
$btnBug.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnBug.FlatAppearance.BorderSize = 0
$btnBug.BackColor = [System.Drawing.Color]::FromArgb(168, 32, 38)
$btnBug.ForeColor = [System.Drawing.Color]::White
$btnBug.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$btnBug.Add_Click({ Start-ToolDetached -ScriptName 'BugFix.ps1' })
$form.Controls.Add($btnBug)
$toolTip.SetToolTip($btnBug, 'Report something that is broken. Goes to the same place, and a copy is sent to you.')

# 30px of clear space at the top before the first panel. The help '?' is
# placed at ClientSize-40 by Add-ToolHelpButton, which put it at x880 y8 --
# straight through the Toolkit Health group box. Measured, then moved, rather
# than trusting the button's run-time search to find a gap (14l).
# THIS is the line that governs the window, not the $form.Size above it: it is
# assigned later, and the last assignment wins. Adding a section below y=720
# without changing this number renders it and then clips it off the bottom,
# which is exactly what happened when the Configure block was added.
#
# 858 = Configure block bottom (716 + 130) plus a 12px margin. The block grew
# by 34 when its buttons wrapped to a second row, so this grew by 34 too.
# 890 = the same sum with the block at 162 for its third row (2026-09-19).
$form.ClientSize = New-Object System.Drawing.Size(920, 890)
#endregion

#region Population --------------------------------------------------------------
function Update-EnvironmentPanel {
    param([switch]$Refresh)
    try {
        # The console's own probe, not a second implementation. -SkipSpo on the
        # first pass for the same reason the console skips it: Get-SPOTenant is
        # a network round trip and this page should open immediately.
        $map = if ($Refresh) { Get-CapabilityStateMap -Refresh } else { Get-CapabilityStateMap -SkipSpo }

        foreach ($pair in @(
            @{ Lbl = $lblExo; Key = 'ExO' },
            @{ Lbl = $lblEms; Key = 'EMS' },
            @{ Lbl = $lblAd;  Key = 'AD'  },
            @{ Lbl = $lblSpo; Key = 'SPO' })) {

            # A row this tenant does not have was never created, so there is
            # nothing to write to. Skipping on null rather than checking the
            # capability again keeps one source of truth: the panel decided at
            # build time, and this only fills in what it built.
            if (-not $pair.Lbl) { continue }

            $st = $map[$pair.Key]
            if (-not $st -or $st.Unknown) {
                $pair.Lbl.Text = 'Not checked'; $pair.Lbl.ForeColor = $script:ColDim
            } elseif ($st.Available) {
                $pair.Lbl.Text = 'Connected';   $pair.Lbl.ForeColor = $script:ColGood
            } else {
                $pair.Lbl.Text = 'Not connected'; $pair.Lbl.ForeColor = $script:ColBad
            }
        }
    } catch {
        foreach ($l in @($lblExo, $lblSpo, $lblEms, $lblAd)) {
            if (-not $l) { continue }
            $l.Text = 'Unknown'; $l.ForeColor = $script:ColDim
        }
    }

    # Same two calls the console header makes, in the same order. The mismatch
    # test matters more here than there: this page is read at a glance at the
    # start of the day, which is exactly when "connected to the wrong tenant"
    # is cheapest to catch and most expensive to miss.
    try {
        $d = Get-ConnectedTenantDomain
        if ($d) {
            $lblTenant.Text = "$d"; $lblTenant.ForeColor = $script:ColGood
        } else {
            $lblTenant.Text = 'Not connected'; $lblTenant.ForeColor = $script:ColDim
        }
        $mismatch = Test-TenantDomainMismatch
        if ($mismatch) {
            # Amber and non-blocking, matching the console: a tenant with several
            # accepted domains can legitimately disagree with its onmicrosoft
            # prefix, so this is a question, not a fault.
            $lblTenant.Text = "$($lblTenant.Text)  (?)"
            $lblTenant.ForeColor = $script:ColWarn
            $script:TenantNote = $mismatch
        } else {
            $script:TenantNote = ''
        }
        $toolTip.SetToolTip($lblTenant, "$script:TenantNote")
    } catch { $lblTenant.Text = 'Unknown'; $lblTenant.ForeColor = $script:ColDim }

    # PIM is skipped entirely where the tenant does not use it. Get-PimStatusLine
    # would otherwise report "no active assignments" -- true, and meaningless, in
    # a tenant with no PIM at all.
    if ($lblPim) {
        try {
            $lblPim.Text      = Get-PimStatusLine
            $lblPim.ForeColor = Get-PimStatusColour
        } catch { $lblPim.Text = 'Unknown'; $lblPim.ForeColor = $script:ColDim }
    }
}

function Update-HealthPanel {
    try {
        $v = Test-ToolkitIntegrity
        if (-not $v.HasManifest) {
            $lblFiles.Text = 'Not stamped'; $lblFiles.ForeColor = $script:ColDim
            $lblDrift.Text = 'Run Publish-Toolkit.ps1 on the master'; $lblDrift.ForeColor = $script:ColDim
            $lblVer.Text   = 'unversioned';  $lblVer.ForeColor = $script:ColDim
            $lblRel.Text   = '-'
        } else {
            $total = $v.Matched + $v.Modified.Count + $v.Missing.Count
            $lblFiles.Text = "$($v.Matched) of $total"
            $bad = $v.Modified.Count + $v.Missing.Count + $v.Extra.Count
            if ($bad -eq 0) {
                $lblFiles.ForeColor = $script:ColGood
                $lblDrift.Text = 'None'; $lblDrift.ForeColor = $script:ColGood
            } else {
                $lblFiles.ForeColor = $script:ColWarn
                $bits = @()
                if ($v.Modified.Count) { $bits += "$($v.Modified.Count) modified" }
                if ($v.Missing.Count)  { $bits += "$($v.Missing.Count) missing" }
                if ($v.Extra.Count)    { $bits += "$($v.Extra.Count) extra" }
                $lblDrift.Text = ($bits -join ', ')
                $lblDrift.ForeColor = $script:ColWarn
            }
            $lblVer.Text = "v$($v.Version)"
            $lblVer.ForeColor = $(if ($bad -eq 0) { $script:ColGood } else { $script:ColWarn })
            $lblRel.Text = "$($v.Released)"
        }
    } catch {
        $lblFiles.Text = 'Unknown'; $lblFiles.ForeColor = $script:ColDim
        $lblDrift.Text = 'Could not check'; $lblDrift.ForeColor = $script:ColDim
    }

    # Config files: present and parseable are different questions, and the
    # second is the one that bites -- a truncated JSON is present.
    foreach ($pair in @(
        @{ Lbl = $lblTools; Name = 'Tools.json' },
        @{ Lbl = $lblQuick; Name = 'QuickLaunch.json' },
        @{ Lbl = $lblSet;   Name = 'Settings.json' })) {
        try {
            $p = Get-ConfigFilePath -Name $pair.Name
            if (-not (Test-Path -LiteralPath $p)) {
                $pair.Lbl.Text = 'Missing'; $pair.Lbl.ForeColor = $script:ColWarn
                continue
            }
            $null = Get-Content -LiteralPath $p -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $pair.Lbl.Text = 'Loaded'; $pair.Lbl.ForeColor = $script:ColGood
        } catch {
            $pair.Lbl.Text = 'INVALID JSON'; $pair.Lbl.ForeColor = $script:ColBad
        }
    }
}

function Update-LogPanels {
    param($Log)

    $rows = @($Log.Rows)
    if ($Log.Problem) {
        foreach ($l in @($lblActions, $lblOk, $lblPart, $lblFail, $lblUnk)) {
            $l.Text = '-'; $l.ForeColor = $script:ColDim
        }
        $lblLast.Text = $Log.Problem; $lblLast.ForeColor = $script:ColWarn
        $lblLastAt.Text = ''
        $lblTopNote.Text = $Log.Problem
        $lblHotNote.Text = $Log.Problem
        return
    }

    # ---- Today ----------------------------------------------------------
    $today = (Get-Date).Date
    $todayRows = @($rows | Where-Object {
        $d = Get-RowDate -Row $_
        $d -and $d.Date -eq $today
    })

    if ($todayRows.Count -eq 0) {
        foreach ($l in @($lblOk, $lblPart, $lblFail, $lblUnk)) { $l.Text = '0'; $l.ForeColor = $script:ColDim }
        $lblActions.Text = '0'; $lblActions.ForeColor = $script:ColDim
        $lblLast.Text = 'No activity logged today.'; $lblLast.ForeColor = $script:ColDim
        $lblLastAt.Text = ''
    } else {
        $ok   = @($todayRows | Where-Object { $_.Result -eq 'Success' }).Count
        $part = @($todayRows | Where-Object { $_.Result -eq 'Partial' }).Count
        $fail = @($todayRows | Where-Object { $_.Result -eq 'Failed'  }).Count
        # Anything not one of the three. Counted, never absorbed into Success.
        $unk  = $todayRows.Count - $ok - $part - $fail

        $lblActions.Text = "$($todayRows.Count)"; $lblActions.ForeColor = [System.Drawing.Color]::Empty
        $lblOk.Text   = "$ok";   $lblOk.ForeColor   = $(if ($ok)   { $script:ColGood } else { $script:ColDim })
        $lblPart.Text = "$part"; $lblPart.ForeColor = $(if ($part) { $script:ColWarn } else { $script:ColDim })
        $lblFail.Text = "$fail"; $lblFail.ForeColor = $(if ($fail) { $script:ColBad }  else { $script:ColDim })
        $lblUnk.Text  = "$unk";  $lblUnk.ForeColor  = $(if ($unk)  { $script:ColDim }  else { $script:ColDim })

        $last = $todayRows | Sort-Object { Get-RowDate -Row $_ } | Select-Object -Last 1
        $desc = @("$($last.Action)", "$($last.Target)") | Where-Object { $_ }
        if (-not $desc) { $desc = @("$($last.Detail)") }
        $lblLast.Text = ((@(Get-RowToolName -Row $last) + $desc) -join ' - ')
        $lblLast.ForeColor = [System.Drawing.Color]::Empty
        $d = Get-RowDate -Row $last
        $lblLastAt.Text = $(if ($d) { $d.ToString('HH:mm') } else { '' })
        if ("$($last.Result)") { $lblLastAt.Text += "   ($($last.Result))" }
    }

    # ---- 30-day windows -------------------------------------------------
    $since = (Get-Date).Date.AddDays(-30)
    $recent = @($rows | Where-Object {
        $d = Get-RowDate -Row $_
        $d -and $d -ge $since
    })

    $lvTop.Items.Clear()
    $lvHot.Items.Clear()

    if ($recent.Count -eq 0) {
        $lblTopNote.Text = 'No tool usage found for the selected period.'
        $lblHotNote.Text = 'No failures found for the selected period.'
        return
    }

    $byTool = $recent | Group-Object { Get-RowToolName -Row $_ } |
              Sort-Object Count -Descending | Select-Object -First 5
    foreach ($g in $byTool) {
        $it = New-Object System.Windows.Forms.ListViewItem($g.Name)
        [void]$it.SubItems.Add("$($g.Count)")
        [void]$lvTop.Items.Add($it)
    }
    $noScript = @($recent | Where-Object { -not "$($_.Script)".Trim() }).Count
    $lblTopNote.Text = $(if ($noScript) {
        "$($recent.Count) actions. $noScript have no script name (converted rows)."
    } else {
        "$($recent.Count) actions in the last 30 days."
    })

    $bad = @($recent | Where-Object { $_.Result -eq 'Failed' -or $_.Result -eq 'Partial' })
    if ($bad.Count -eq 0) {
        $lblHotNote.Text = 'No failures found for the selected period.'
    } else {
        $byBad = $bad | Group-Object { Get-RowToolName -Row $_ } |
                 Sort-Object { @($_.Group | Where-Object { $_.Result -eq 'Failed' }).Count } -Descending |
                 Select-Object -First 5
        foreach ($g in $byBad) {
            $f = @($g.Group | Where-Object { $_.Result -eq 'Failed'  }).Count
            $p = @($g.Group | Where-Object { $_.Result -eq 'Partial' }).Count
            $it = New-Object System.Windows.Forms.ListViewItem($g.Name)
            [void]$it.SubItems.Add("$f")
            [void]$it.SubItems.Add("$p")
            [void]$lvHot.Items.Add($it)
        }
        $lblHotNote.Text = "$($bad.Count) failed or partial in the last 30 days."
    }
}

function Update-RetryPanel {
    $lvRetry.Items.Clear()
    try {
        $folder = Get-RetryQueuePath
        $csv = Join-Path $folder 'disable-retry.csv'
        if (-not (Test-Path -LiteralPath $csv)) {
            $lblRetryCount.Text = 'No pending retries.'
            $lblRetryCount.ForeColor = $script:ColDim
            return
        }
        $rows = @(Import-Csv -LiteralPath $csv -ErrorAction Stop)
        if ($rows.Count -eq 0) {
            $lblRetryCount.Text = 'No pending retries.'
            $lblRetryCount.ForeColor = $script:ColDim
            return
        }
        $lblRetryCount.Text = "$($rows.Count) waiting"
        $lblRetryCount.ForeColor = $script:ColWarn
        foreach ($r in ($rows | Select-Object -First 6)) {
            $it = New-Object System.Windows.Forms.ListViewItem("$($r.Employee)")
            [void]$it.SubItems.Add("$($r.FailedStep)")
            [void]$it.SubItems.Add("$($r.Attempts)")
            [void]$it.SubItems.Add("$($r.LastAttempt)")
            [void]$lvRetry.Items.Add($it)
        }
    } catch {
        $lblRetryCount.Text = 'Could not read the queue.'
        $lblRetryCount.ForeColor = $script:ColDim
    }
}

function Update-Dashboard {
    param([switch]$Refresh)
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        # Each panel is independent on purpose. One unreadable file must not
        # take the other five down -- that is the difference between a page
        # that is partly useful and a page that is blank.
        Update-EnvironmentPanel -Refresh:$Refresh
        Update-HealthPanel
        Update-LogPanels -Log (Get-DashLogRows)
        Update-RetryPanel
        $lblStatus.Text = "Refreshed $(Get-Date -Format 'HH:mm:ss').  Read-only -- nothing on this page changes anything."
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

$btnRefresh.Add_Click({ Update-Dashboard -Refresh })
#endregion

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Dashboard' | Out-Null

Update-Dashboard

Write-CompletedLog -Action 'View' -Target 'Dashboard' -Result 'Success' -Detail 'Opened the dashboard'

[void]$form.ShowDialog()

# In-process, deliberately. Control returns here with the dashboard gone and the
# console's sessions still live, which is the whole point -- Start-ToolDetached
# would hand the tool a process with no connections in it.
if ($script:OpenAfterClose) {
    $next = Join-Path $PSScriptRoot $script:OpenAfterClose
    if (Test-Path -LiteralPath $next) {
        & $next
    } else {
        $null = Show-ToolDialog(
            "$($script:OpenAfterClose) is not in this folder.", 'Not found')
    }
}
