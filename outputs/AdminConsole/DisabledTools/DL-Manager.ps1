#Requires -Version 5.1
<#
    DL-Manager.ps1
    One front door for distribution group work. Looks the group up, decides
    where it is mastered, and opens the correct editor for that answer.

    Reasoning: ScriptNotes/DL-Manager.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-19
#>
# TOOLKIT-VARIANT: Router
# TOOLKIT-HELP-BEGIN
# Start here for any distribution group.
#
# Look the group up and this window tells you where it actually lives -- in the
# cloud, or still synced from on-premises -- and then opens the right editor for
# it. You no longer have to know which of the two versions of a tool to pick.
#
# WHY IT MATTERS MORE THAN IT USED TO. During the migration, groups move from
# on-premises to cloud a batch at a time. A group that needed the on-premises
# editor last week may need the cloud one today. Nothing about the group's name
# or address changes when that happens, so there is no way to tell by looking.
#
# THE FAILURE THIS PREVENTS IS SILENT. Editing a converted group in the
# on-premises tools does not produce an error. The change is written to Active
# Directory, Active Directory is no longer the authority for that group, and
# the change never reaches Exchange Online. It looks exactly like success.
#
# The banner colour says whether the lookup WORKED, not where the group lives:
#   GREEN  -- found. The line itself says cloud-managed or synced.
#   RED    -- not found in Exchange Online, so there is nothing to route to.
#
# Read-only. This window changes nothing itself; it only opens other tools.
# TOOLKIT-HELP-END
# TOOLKIT-HELP-RUNS: Reads one group from Exchange Online and opens another tool. Changes nothing.

param(
    [string]$SmtpAddress
)

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
if (-not (Test-ConsoleLaunch)) { return }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Routing table ---------------------------------------------------------
<#
    THE PAIRS ARE DECLARED IN THE SCRIPTS, NOT INVENTED HERE.

    Every tool already carries a 'TOOLKIT-VARIANT: Cloud' or
    'TOOLKIT-VARIANT: OnPremAD' marker in its header. Twenty-two files have one
    and, until this script, nothing read them -- they were a comment convention
    that documented an intent the code never enforced.

    This table names the ACTIONS and lets the marker decide which file serves
    each one. Adding a third variant later is then a change to the scripts
    rather than a change to the router.
#>
$script:Actions = @(
    # ONE ACTION NOW, NOT THREE.
    #
    # This routed to three separate editors -- owners/members, send-to/send-as,
    # bulk membership -- which meant three windows, three lookups of the same
    # group, and three chances to have one of them open against the wrong
    # environment. Everything a distribution group has now lives in one form.
    #
    # The routing decision it makes is unchanged and is still the whole point:
    # IsDirSynced picks the side, and the editor is told rather than asked.
    [PSCustomObject]@{
        Key        = 'Edit'
        Text       = 'Edit this group'
        Tip        = 'Owner, membership, delivery restrictions and Send-As, in one window.'
        Cloud      = 'DL-Manager-Editor.ps1'
        OnPrem     = 'DL-Manager-Editor.ps1'
        OnPremArgs = @('-OnPrem')
    }
)
#endregion

#region Lookup ----------------------------------------------------------------
function Get-DLRoutingState {
    <#
        Where is this group mastered?

        ONE PROPERTY DECIDES IT: IsDirSynced, read from Exchange Online. True
        means Active Directory is the authority and edits belong there; false
        means the object is cloud-managed and Exchange Online is the authority.

        IT IS ALSO THE EXACT PROPERTY THE MIGRATION FLIPS. Transferring Source
        of Authority sets isCloudManaged, and IsDirSynced becomes false. So this
        window reports migration progress as a side effect of doing its job, and
        it cannot drift out of step with the migration because it is reading the
        same fact rather than a record of it.

        DELIBERATELY NOT CACHED for the same reason. A cached answer is correct
        until the weekend the group converts, and then it is confidently wrong.
    #>
    param([string]$Smtp)

    $state = [PSCustomObject]@{
        Smtp        = $Smtp
        Found       = $false
        DisplayName = ''
        PrimarySmtp = ''
        GroupType   = ''
        IsDirSynced = $null
        Where       = ''
        Error       = ''
    }

    try {
        $g = Invoke-EXOCommand -Name 'Get-DistributionGroup' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop' }
    } catch {
        # NOT FOUND IN EXCHANGE ONLINE IS NOT "NO SUCH GROUP".
        #
        # It is far more likely to be a group created on-premises that has not
        # synchronized yet. Saying "does not exist" would send somebody looking
        # for a typo in a name that is perfectly correct.
        $state.Error = ("Exchange Online has no distribution group at that address.`r`n`r`n" +
                        "If it was created on-premises in the last hour or so, it may not have " +
                        "synchronized yet -- there is nothing here to route to until it does.`r`n`r`n" +
                        "Underlying error: $($_.Exception.Message)")
        return $state
    }

    $state.Found       = $true
    $state.DisplayName = "$($g.DisplayName)"
    $state.PrimarySmtp = "$($g.PrimarySmtpAddress)"
    $state.GroupType   = "$($g.RecipientTypeDetails)"
    $state.IsDirSynced = [bool]$g.IsDirSynced
    $state.Where       = if ($state.IsDirSynced) { 'OnPrem' } else { 'Cloud' }
    return $state
}
#endregion

#region Launching -------------------------------------------------------------
function Test-ToolNeedsOwnProcess {
    <#
        Does this tool declare TOOLKIT-REQUIRES: OwnProcess?

        The same question the console asks before launching anything, answered
        the same way -- by reading the marker out of the script rather than by
        keeping a list here. A list would be a second copy of a fact that
        already exists in twenty-odd files, and it would go stale the first time
        a tool changed its mind.

        Only the HEADER is read. The string appears in Common.ps1 and in
        AdminConsole.ps1 as part of the detection logic itself, so a whole-file
        search would report those as tools that need their own process.
    #>
    param([Parameter(Mandatory)][string]$ScriptName)

    $path = Join-Path $PSScriptRoot $ScriptName
    if (-not (Test-Path -LiteralPath $path)) { return $false }
    try {
        $head = (Get-Content -LiteralPath $path -TotalCount 40 -ErrorAction Stop) -join "`n"
    } catch { return $false }
    return [bool]([regex]::IsMatch($head, 'TOOLKIT-REQUIRES:\s*OwnProcess', 'IgnoreCase'))
}

function Start-OnPremToolProcess {
    <#
        Opens an on-premises tool in its OWN PowerShell process, connecting to
        Exchange on-premises first.

        This duplicates Start-ToolScriptInNewProcess in AdminConsole.ps1, and
        that is a deliberate, temporary choice rather than an oversight. The
        right end state is one copy in Common.ps1 with both callers using it.
        Making that move means editing two working files, and there is no
        PowerShell parser available in the environment this was written in --
        so the change with the smaller blast radius won. Consolidate it the
        next time Test-Toolkit can be run.

        WHY A SEPARATE PROCESS AT ALL: the on-premises shell and Exchange
        Online both define Get-DistributionGroup and Set-DistributionGroup.
        Loaded together, whichever came last wins and a cloud write silently
        becomes an on-premises one. The process boundary is what keeps the two
        meanings apart.
    #>
    param(
        [Parameter(Mandatory)][string]$ScriptName,
        [string]$Smtp,
        # Extra switches for the child, e.g. -OnPrem. Passed as separate tokens
        # rather than folded into a string, so the child binds them as switches.
        [string[]]$Extra = @()
    )

    $path = Join-Path $PSScriptRoot $ScriptName
    if (-not (Test-Path -LiteralPath $path)) {
        [System.Windows.Forms.MessageBox]::Show(
            "$ScriptName is not in this folder.", 'Not found',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }

    $psExe = $null
    try { $psExe = (Get-Process -Id $PID).Path } catch { }
    if (-not $psExe) { $psExe = 'powershell.exe' }

    # Every path single-quoted with internal quotes doubled, so the whole
    # -Command payload can be wrapped in double quotes without anything inside
    # fighting it. Start-Process does NOT quote for you.
    $esc = { param([string]$s) "'" + ("$s" -replace "'", "''") + "'" }

    $parts = New-Object System.Collections.ArrayList
    [void]$parts.Add((Get-ConsoleResizeCommand))

    $connect = Join-Path $PSScriptRoot 'Connect-ExSE.ps1'
    if (Test-Path -LiteralPath $connect) { [void]$parts.Add("& $(& $esc $connect)") }

    $call = "& $(& $esc $path)"
    if ("$Smtp".Trim()) { $call += " -SmtpAddress $(& $esc $Smtp)" }
    foreach ($x in $Extra) { if ("$x".Trim()) { $call += " $x" } }
    [void]$parts.Add($call)

    # -NoExit so the window survives the tool closing: the output is the record
    # of what happened, and a window that vanishes takes it with it.
    $argLine = '-NoExit -ExecutionPolicy Bypass -Command "' + ($parts -join '; ') + '"'
    try {
        Start-Process -FilePath $psExe -ArgumentList $argLine | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Could not launch $ScriptName`r`n$($_.Exception.Message)", 'Launch failed',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    }
}
#endregion

#region Target ----------------------------------------------------------------
# NO PROMPT BEFORE THE WINDOW. The address box is ON the form.
#
# This used to call Resolve-ToolTarget first, which pops its own modal dialog
# when no address was supplied -- so opening the tool meant answering a
# question in one window and then reading the answer in another. Worse, looking
# up a second group meant closing the whole thing and starting again, which is
# exactly what you do when a ticket lists four lists.
#
# Resolve-ToolTarget is not used at all now. Its job was to decide which
# DIRECTORY an address belongs to, and this window's entire purpose is that
# Exchange Online answers that question instead. Get-TargetSmtp is kept for the
# one useful thing it did: completing a bare alias to a full address.
$script:State = $null
#endregion

#region WinForms UI -----------------------------------------------------------
# SET BY A BUTTON, ACTED ON AFTER ShowDialog RETURNS.
#
# A cloud tool is opened in THIS process so it inherits the live Exchange
# Online session. It cannot be opened while this form is up: a modal form
# disables the other windows on its thread, so the tool would draw and then
# refuse input. Same reason the dashboard defers its session-needing tools.
$script:OpenAfterClose = ''

# Extra switches for the deferred tool, carried separately from its name.
# -Cloud is the only one today; keeping it as an array means a second one
# does not require another script-scope variable.
$script:OpenAfterCloseArgs = @()

# Collected while the buttons are being enabled, printed after. Building the
# detail text first and the buttons second would mean deciding the same
# thing twice; this way the window reports exactly what the enable logic did.
$script:Unavailable = New-Object System.Collections.ArrayList

# THE ACTION BUTTONS ARE TRACKED EXPLICITLY.
#
# The enable/disable logic walked $form.Controls and treated any Button with a
# non-null Tag as an action. Add-ToolHelpButton adds the '?' as a BUTTON and
# tags it 'NoTheme' so the theme skips it -- so the '?' was read as an action,
# produced no tool name, and printed "  : no tool is configured for a cloud
# group" with an empty action name in front of the colon.
#
# It only appeared from the SECOND lookup onwards, because the help button is
# added after the first Update-Verdict has already run. That is why the window
# looked correct the first time and wrong every time after.
$script:ActionButtons = New-Object System.Collections.ArrayList

$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Distribution Group -- where does it live?'
$form.Size          = New-Object System.Drawing.Size(720, 400)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

# --- Address row ----------------------------------------------------------
$lblAddr          = New-Object System.Windows.Forms.Label
$lblAddr.Text     = 'Distribution group (name or address):'
$lblAddr.Location = New-Object System.Drawing.Point(12, 15)
$lblAddr.AutoSize = $true
$form.Controls.Add($lblAddr)

$txtAddr          = New-Object System.Windows.Forms.TextBox
$txtAddr.Location = New-Object System.Drawing.Point(232, 12)
$txtAddr.Size     = New-Object System.Drawing.Size(250, 22)
$txtAddr.Font     = New-Object System.Drawing.Font('Consolas', 9)
$txtAddr.Anchor   = 'Top,Left,Right'
$form.Controls.Add($txtAddr)

$btnLookup          = New-Object System.Windows.Forms.Button
$btnLookup.Text     = 'Look Up'
# LEFT OF THE HELP BUTTON, not under it.
#
# Add-ToolHelpButton places the '?' at ClientSize-40 at run time, and Look Up
# was sitting at x532 in a 644px client -- so the two overlapped and the button
# read as 'Look Up?'. The address box gave up the 40px rather than the window
# growing to hide the collision.
$btnLookup.Location = New-Object System.Drawing.Point(492, 10)
$btnLookup.Size     = New-Object System.Drawing.Size(100, 26)
$btnLookup.Anchor   = 'Top,Right'
$form.Controls.Add($btnLookup)

# ENTER LOOKS UP. Typing an address and pressing Enter is what everyone does,
# and without this it did nothing at all.
$form.AcceptButton = $btnLookup

# --- Verdict banner -------------------------------------------------------
$lblBanner           = New-Object System.Windows.Forms.Label
$lblBanner.Tag       = 'NoTheme'
$lblBanner.Location  = New-Object System.Drawing.Point(12, 48)
$lblBanner.Size      = New-Object System.Drawing.Size(680, 34)
$lblBanner.Anchor    = 'Top,Left,Right'
$lblBanner.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblBanner.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 11, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblBanner)

$txtDetail            = New-Object System.Windows.Forms.TextBox
$txtDetail.Multiline  = $true
$txtDetail.ReadOnly   = $true
$txtDetail.ScrollBars = 'Vertical'
$txtDetail.Location   = New-Object System.Drawing.Point(12, 90)
$txtDetail.Size       = New-Object System.Drawing.Size(680, 196)
$txtDetail.Anchor     = 'Top,Left,Right,Bottom'
$txtDetail.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtDetail)

$lblPick          = New-Object System.Windows.Forms.Label
$lblPick.Location = New-Object System.Drawing.Point(12, 296)
$lblPick.Size     = New-Object System.Drawing.Size(680, 18)
$lblPick.Anchor   = 'Left,Bottom'
$form.Controls.Add($lblPick)

# --- Action buttons -------------------------------------------------------
$toolTip = New-Object System.Windows.Forms.ToolTip
$btnX = 12
foreach ($a in $script:Actions) {
    $b          = New-Object System.Windows.Forms.Button
    $b.Text     = $a.Text
    $b.Location = New-Object System.Drawing.Point($btnX, 318)
    $b.Size     = New-Object System.Drawing.Size(150, 30)
    $b.Anchor   = 'Left,Bottom'

    # Tag carries the ACTION, so the shared handler reads it from the control
    # that was clicked. A handler closing over the loop variable would see
    # whatever it held when the loop ENDED -- every button would open the last
    # tool in the table.
    $b.Tag = $a
    [void]$script:ActionButtons.Add($b)
    $b.Add_Click({
        $act = $this.Tag
        if (-not $act) { return }
        $toolFile = if ($script:State.Where -eq 'Cloud') { "$($act.Cloud)" } else { "$($act.OnPrem)" }
        if (-not $toolFile) { return }

        # WHERE THE GROUP LIVES AND WHERE THE TOOL RUNS ARE TWO DIFFERENT
        # QUESTIONS, and conflating them was wrong.
        #
        # This used to send every on-premises tool into a new process with a
        # full Connect-ExSE in front of it. Only two of them need that:
        # AD-DL_SendToAs and Bulk-DLMembership, which declare
        # TOOLKIT-REQUIRES: OwnProcess because they use Exchange management
        # shell cmdlets that collide with Exchange Online's.
        #
        # AD-GroupEditor does NOT. It works through the ActiveDirectory module,
        # which has no such collision, so it runs perfectly well in this
        # process. Launching it detached bought nothing and cost thirty seconds
        # of Exchange connect before its window appeared -- long enough that
        # pressing the button looked like it had done nothing at all.
        #
        # THE CLOUD SIDE IS ALWAYS IN-PROCESS, WHATEVER THE MARKER SAYS.
        #
        # The Exchange Online session lives in the console's process and cannot
        # be handed to a child, so a cloud tool launched detached would open
        # disconnected. Bulk-DLMembership makes this explicit: it carries
        # TOOLKIT-REQUIRES: OwnProcess for its on-premises identity, and running
        # it with -Cloud is precisely the case where that marker must not apply.
        #
        # The marker still decides for the on-premises side, where it is right:
        # AD-GroupEditor works through the ActiveDirectory module and needs no
        # window of its own, while AD-DL_SendToAs and Bulk-DLMembership need an
        # Exchange management shell that must not share a process with the cloud.
        if ($script:State.Where -ne 'Cloud' -and (Test-ToolNeedsOwnProcess -ScriptName $toolFile)) {
            # Its own process, its own on-premises Exchange session. The form
            # stays open, because nothing about it conflicts with the child.
            $extra = @()
            if ($act.PSObject.Properties['OnPremArgs']) { $extra = @($act.OnPremArgs) }
            Start-OnPremToolProcess -ScriptName $toolFile -Smtp $script:State.PrimarySmtp -Extra $extra
        } else {
            # Deferred: opened by this script after the form closes, in this
            # process. A modal form disables the other windows on its thread, so
            # a tool opened while this one is up would draw and refuse input.
            $script:OpenAfterClose = $toolFile
            $script:OpenAfterCloseArgs = @()
            if ($script:State.Where -eq 'Cloud' -and $act.PSObject.Properties['CloudArgs']) {
                $script:OpenAfterCloseArgs = @($act.CloudArgs)
            }
            $form.Close()
        }
    })
    $form.Controls.Add($b)
    $toolTip.SetToolTip($b, "$($a.Tip)")
    $btnX += 160
}

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(582, 318)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Anchor   = 'Right,Bottom'
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

# --- Lookup -----------------------------------------------------------------
function Update-Verdict {
    <#
        Reads the group and repaints the window. CALLED ON EVERY LOOK UP, not
        once at startup.

        This was straight-line code that ran while the form was being built, so
        the answer was fixed before the window appeared and a second group meant
        a second run of the tool. Wrapping it means the same window answers as
        many groups as the ticket contains -- and, during the migration, lets
        you re-check one you have just converted without reopening anything.
    #>
    $addr = Get-TargetSmtp -Text $txtAddr.Text
    if (-not $addr) {
        $lblBanner.Text      = 'Enter a distribution group above'
        $lblBanner.ForeColor = [System.Drawing.Color]::DimGray
        $txtDetail.Text      = ''
        $lblPick.Text        = ''
        foreach ($c in $script:ActionButtons) { $c.Enabled = $false }
        return
    }

    # RE-ENABLED FIRST, then disabled by the rules below. A button left disabled
    # by the PREVIOUS group would otherwise stay dead for a group that supports
    # it -- the failure gets stickier the more groups you look at.
    foreach ($c in $script:ActionButtons) { $c.Enabled = $true }
    $script:Unavailable.Clear()

    $form.Cursor    = [System.Windows.Forms.Cursors]::WaitCursor
    $btnLookup.Text = 'Looking up...'
    $form.Refresh()
    try {
        $script:State = Get-DLRoutingState -Smtp $addr
    } finally {
        $form.Cursor    = [System.Windows.Forms.Cursors]::Default
        $btnLookup.Text = 'Look Up'
    }

        $sb = New-Object System.Text.StringBuilder

    if (-not $script:State.Found) {
        $lblBanner.Text      = 'Not found in Exchange Online'
        $lblBanner.ForeColor = [System.Drawing.Color]::Firebrick
        [void]$sb.AppendLine($script:State.Error)
        $lblPick.Text = 'Nothing to open.'
        # LOOK UP STAYS ENABLED. This disabled every button but Close, which
        # included the one you need to correct a typo -- a mistyped address
        # left the window with nothing to do but shut it and start again.
        foreach ($c in $script:ActionButtons) { $c.Enabled = $false }
    } else {
        if ($script:State.Where -eq 'Cloud') {
            $lblBanner.Text      = 'CLOUD-MANAGED  --  edit in Exchange Online'
            # THE CONSOLE'S GREEN, not a second one.
            #
            # This was a dark green picked for a white background and then put on
            # the console's dark theme, where it reads as murky rather than as
            # "good". 80,220,130 is the exact colour AdminConsole uses for a
            # connected tenant and a current version, so "green means fine" carries
            # across windows instead of being relearned in each.
            $lblBanner.ForeColor = [System.Drawing.Color]::FromArgb(80, 220, 130)
        } else {
            $lblBanner.Text      = 'SYNCED FROM ON-PREMISES  --  edit at the source'
            # GREEN, THE SAME AS CLOUD. Both are successful lookups.
            #
            # This was amber, and amber was wrong. Colour was carrying WHERE the
            # group lives, which the banner already states in words -- "SYNCED
            # FROM ON-PREMISES" needs no help from the palette. Using amber for a
            # normal, correct result teaches you to ignore amber, and then it is
            # useless for the cases that deserve it.
            #
            # Amber and red are now free for states that actually want
            # attention: not found, blocked, or a group with no owner.
            $lblBanner.ForeColor = [System.Drawing.Color]::FromArgb(80, 220, 130)
        }

        [void]$sb.AppendLine("Group            : $($script:State.DisplayName)")
        [void]$sb.AppendLine("Address          : $($script:State.PrimarySmtp)")
        [void]$sb.AppendLine("Type             : $($script:State.GroupType)")
        [void]$sb.AppendLine("IsDirSynced      : $($script:State.IsDirSynced)")
        [void]$sb.AppendLine('')

        if ($script:State.Where -eq 'Cloud') {
            [void]$sb.AppendLine('Exchange Online is the authority for this group. The buttons below open')
            [void]$sb.AppendLine('the cloud tools, in this window''s session.')
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('Either it was created in the cloud, or its Source of Authority has already')
            [void]$sb.AppendLine('been transferred. Editing it on-premises would write to an object that is')
            [void]$sb.AppendLine('no longer authoritative -- silently, with no error.')
        } else {
            [void]$sb.AppendLine('Active Directory is the authority for this group. The buttons below open')
            [void]$sb.AppendLine('the on-premises tools in their own window, which takes a moment while')
            [void]$sb.AppendLine('that session connects.')
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('Exchange Online will refuse membership and owner changes on this group,')
            [void]$sb.AppendLine('so there is nothing to be gained by trying them there first.')
        }

        $lblPick.Text = 'Choose what you need to change:'

        # Disable any action with no tool for this side, and SAY WHY on the button
        # itself rather than leaving a dead control with no explanation.
        foreach ($c in $script:ActionButtons) {
            $act = $c.Tag
            if (-not $act) { continue }
            $toolFile = if ($script:State.Where -eq 'Cloud') { "$($act.Cloud)" } else { "$($act.OnPrem)" }
            # A DISABLED BUTTON MUST SAY WHY ON THE WINDOW, NOT ONLY IN A TOOLTIP.
            #
            # Two of these greyed out in testing and the reason -- the script was
            # not in the folder -- was reachable only by hovering, which nobody
            # does when they are trying to work out whether the TOOL is broken.
            # A greyed button with no visible cause reads as a fault in this window.
            if (-not $toolFile) {
                $c.Enabled = $false
                $reason = "$($act.Text): no tool is configured for a $($script:State.Where.ToLower()) group."
                [void]$script:Unavailable.Add($reason)
                $toolTip.SetToolTip($c, $reason)
            } elseif (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot $toolFile))) {
                $c.Enabled = $false
                $reason = "$($act.Text): $toolFile is not in the console folder."
                [void]$script:Unavailable.Add($reason)
                $toolTip.SetToolTip($c, $reason)
            }
        }
    }

    if ($script:Unavailable.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('--- Unavailable here ---')
        foreach ($u in $script:Unavailable) { [void]$sb.AppendLine("  $u") }
    }

    $txtDetail.Text = $sb.ToString()


}

$btnLookup.Add_Click({ Update-Verdict })

# An address supplied by the console is looked up immediately -- having typed it
# once in the console's own box, being asked for it again is the thing this
# change exists to remove.
if ("$SmtpAddress".Trim()) {
    $txtAddr.Text = "$SmtpAddress".Trim()
    Update-Verdict
} else {
    Update-Verdict
}

Set-ConsoleTheme -Form $form
Set-FormResizable -Form $form

# BOTTOM-LEFT, not the middle of the screen.
#
# StartPosition was CenterScreen, which drops this window on top of whatever
# you were reading -- usually the console itself, or the ticket. Every other
# tool in the toolkit corners itself, and Set-FormScreenCorner must run AFTER
# Set-FormResizable because that sets MinimumSize and can move the form.
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'DL Manager' | Out-Null

[void]$form.ShowDialog()

# AFTER the dialog, never during. See $script:OpenAfterClose above.
if ($script:OpenAfterClose) {
    $path = Join-Path $PSScriptRoot $script:OpenAfterClose
    if (Test-Path -LiteralPath $path) {
        # SPLATTING DOES NOT ACCEPT A SCOPE-QUALIFIED NAME.
        #
        # This read '@script:OpenAfterCloseArgs' and did not splat at all --
        # PowerShell takes @ + name, and 'script:OpenAfterCloseArgs' is not a
        # plain variable name. No error: the switch was simply never passed, so
        # the cloud tool opened in on-premises mode and then correctly refused
        # to run. The symptom pointed at the child; the fault was here.
        #
        # A local copy is the fix, and it is why splat variables in this toolkit
        # should stay local to the line that uses them.
        $extraArgs = @($script:OpenAfterCloseArgs)
        & $path -SmtpAddress $script:State.PrimarySmtp @extraArgs
    }
}
#endregion
