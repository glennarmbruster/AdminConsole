#Requires -Version 5.1
# TOOLKIT-HELP-RUNS: nothing. It opens the PIM portal in your admin browser profile.
# TOOLKIT-HELP-BEGIN
# Opens Privileged Identity Management so you can activate your admin access.
#
# Your rights are eligible, not standing -- until you activate, you are signed
# in with no admin access, and tools fail in ways that look like faults.
#
# ORDER MATTERS:
#
#     1. Activate FIRST
#     2. Connect SECOND  (ExO / SPO / OnPrem)
#
# A connection made before activating keeps the token it was issued with, and
# that token does not contain access activated afterwards. If you have already
# connected, activate and then reconnect.
#
# HOW IT WORKS:
#
#   ACTIVATE       -- opens the portal so you can turn access on.
#   MARK ACTIVATED -- press after you have, to start a 4-hour countdown.
#
# The countdown then shows here and in the console header, going amber under
# 30 minutes and red once it has run out. That is the "do I need to activate
# again?" answer at a glance.
#
# It is a NOTE, not a live reading. Entra is not consulted -- the tenant will
# not consent to the Graph permission that would make it authoritative, and
# the portal's Active-assignments tab has no address of its own. So the
# countdown knows when you pressed the button, plus four hours, and always
# shows the time it recorded so you can see how old it is.
#
# If you ever activate without pressing the button, the countdown is wrong.
# Press Clear and start again.
#
# WHICH ROW:
#
#   * Groups -- the normal route. Covers Exchange and most day-to-day work.
#
#   * Entra roles -- use this for SHAREPOINT. SharePoint access can be
#     activated through the group, and it does work, but it is noticeably
#     slower to take effect because the group membership has to propagate
#     before SharePoint honours it. Activating the singular Entra role is
#     faster and is the route to prefer when SPO is what you need.
#
# All four open in your admin Edge profile, so you are not signed in as the
# wrong account when you land.
#
# Why there is no live 'what is active' button: active assignments are a TAB
# inside the same blade as the eligible ones, and the portal keeps the selected
# tab in client-side state rather than in the address -- confirmed 2026-08-07.
# There is nothing to link to, so the countdown is the answer instead.
#
# The duration comes from Settings.json as PimDurationHours, defaulting to 4.
#
# Activation itself stays in the portal. Doing it from PowerShell needs
# Microsoft Graph permissions this tenant does not consent to.
# TOOLKIT-HELP-END
<#
    activate.ps1
    Opens the PIM activation and active-assignment blades in the admin
    browser profile.  (Was PimOpen.ps1; renamed to match the console button.)

    Reasoning: ScriptNotes/PimActivation.ps1.notes.md

    Author   : Glenn Armbruster
    Created  : 2026-08-06
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

<#
    These are Azure portal INTERNAL routes. Microsoft does not document them as
    a supported deep-link contract, and they have changed shape before -- the
    older #blade/ form was replaced by #view/.

    So they are settings, not constants. If a link stops landing on the right
    page, open the page by hand, copy the address bar, and paste it into
    Settings.json as PimGroupsUrl / PimRolesUrl. That is a thirty-second fix
    by whoever hits it, rather than a code change.
#>
$script:PimUrlDefaults = @{
    GroupsEligible = 'https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadgroup'
    RolesEligible  = 'https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadmigratedroles'

    # The ACTIVE tab is a tab WITHIN those same blades, and as far as I can
    # tell the portal keeps the selected tab in client-side state rather than
    # in the URL -- so there may be no address for it at all.
    #
    # These therefore default to the same blade. The button still earns its
    # place: right blade, right browser profile, one click, then the Active
    # tab. But it is not a true deep link and the window says so rather than
    # pretending.
    #
    # If your portal DOES give a distinct address when the Active tab is
    # selected, copy it from the address bar into Settings.json as
    # PimGroupsActiveUrl / PimRolesActiveUrl and these buttons become real
    # deep links with no code change.
    GroupsActive   = 'https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadgroup'
    RolesActive    = 'https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadmigratedroles'
}
$script:PimUrlSettingKey = @{
    GroupsEligible = 'PimGroupsUrl'
    RolesEligible  = 'PimRolesUrl'
    GroupsActive   = 'PimGroupsActiveUrl'
    RolesActive    = 'PimRolesActiveUrl'
}

function Get-PimUrl {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('GroupsEligible', 'RolesEligible', 'GroupsActive', 'RolesActive')]
        [string]$Which
    )
    $s = $null
    try { $s = Get-ToolkitSettings } catch { }
    $key = $script:PimUrlSettingKey[$Which]
    if ($s -and $s.PSObject.Properties[$key] -and "$($s.$key)".Trim()) {
        return "$($s.$key)".Trim()
    }
    return $script:PimUrlDefaults[$Which]
}

function Test-PimUrlIsConfigured {
    # True when a setting has been supplied, i.e. the URL is not just my guess.
    param([Parameter(Mandatory)][string]$Which)
    $s = $null
    try { $s = Get-ToolkitSettings } catch { }
    $key = $script:PimUrlSettingKey[$Which]
    return [bool]($s -and $s.PSObject.Properties[$key] -and "$($s.$key)".Trim())
}

#region UI ----------------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'PIM -- activate and check your access'
$form.Size          = New-Object System.Drawing.Size(760, 522)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
$form.MaximizeBox   = $false

$lblOrder           = New-Object System.Windows.Forms.Label
$lblOrder.Tag       = 'NoTheme'
$lblOrder.Text      = 'Activate FIRST, then connect. A session connected before you activate keeps the old token.'
$lblOrder.Location  = New-Object System.Drawing.Point(14, 14)
$lblOrder.Size      = New-Object System.Drawing.Size(620, 18)
$lblOrder.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblOrder.ForeColor = [System.Drawing.Color]::FromArgb(230, 170, 60)
$form.Controls.Add($lblOrder)

# Two columns: ACTIVATE on the left, CHECK on the right. The question "do I
# need to activate again" is answered by the right-hand column, so it gets
# equal billing rather than being buried behind the same button.
# One heading, not two. The second column it labelled is gone -- a header
# over empty space is worse than no header, because it implies something is
# missing rather than that nothing belongs there.
$lblColA          = New-Object System.Windows.Forms.Label
$lblColA.Text     = 'ACTIVATE  --  open the portal, then mark it below'
$lblColA.Location = New-Object System.Drawing.Point(14, 48)
$lblColA.Size     = New-Object System.Drawing.Size(540, 18)
$lblColA.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblColA)

$btnGroupsElig          = New-Object System.Windows.Forms.Button
$btnGroupsElig.Text     = 'Groups'
$btnGroupsElig.Location = New-Object System.Drawing.Point(14, 70)
$btnGroupsElig.Size     = New-Object System.Drawing.Size(160, 38)
$btnGroupsElig.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnGroupsElig)

$lblGroupsNote          = New-Object System.Windows.Forms.Label
$lblGroupsNote.Text     = 'The normal route. Exchange and most day-to-day work.'
$lblGroupsNote.Location = New-Object System.Drawing.Point(186, 80)
$lblGroupsNote.Size     = New-Object System.Drawing.Size(360, 32)
$form.Controls.Add($lblGroupsNote)

$btnRolesElig          = New-Object System.Windows.Forms.Button
$btnRolesElig.Text     = 'Entra roles'
$btnRolesElig.Location = New-Object System.Drawing.Point(14, 120)
$btnRolesElig.Size     = New-Object System.Drawing.Size(160, 38)
$btnRolesElig.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRolesElig)

$lblRolesNote          = New-Object System.Windows.Forms.Label
$lblRolesNote.Text     = 'Use for SharePoint -- faster than waiting for the group to propagate.'
$lblRolesNote.Location = New-Object System.Drawing.Point(186, 124)
$lblRolesNote.Size     = New-Object System.Drawing.Size(360, 36)
$form.Controls.Add($lblRolesNote)

# ONE button, below both routes rather than beside either.
#
# There were two, one per row, which read as "record the groups one" and
# "record the roles one" -- two facts to keep straight when only one is
# tracked. Underneath both, it belongs to whichever route you just used.
$btnActivated          = New-Object System.Windows.Forms.Button
$btnActivated.Text     = 'Activated'
$btnActivated.Location = New-Object System.Drawing.Point(14, 172)
$btnActivated.Size     = New-Object System.Drawing.Size(160, 38)
$btnActivated.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnActivated)

$lblActivatedNote          = New-Object System.Windows.Forms.Label
$lblActivatedNote.Text     = 'Press once you have activated in the portal, whichever route you used. Starts the countdown, and resets it every time.'
$lblActivatedNote.Location = New-Object System.Drawing.Point(186, 176)
$lblActivatedNote.Size     = New-Object System.Drawing.Size(360, 36)
$form.Controls.Add($lblActivatedNote)

$lblCountdown           = New-Object System.Windows.Forms.Label
$lblCountdown.Tag       = 'NoTheme'
$lblCountdown.Location  = New-Object System.Drawing.Point(14, 224)
$lblCountdown.Size      = New-Object System.Drawing.Size(600, 42)
$lblCountdown.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblCountdown)

$btnClear          = New-Object System.Windows.Forms.Button
$btnClear.Text     = 'Clear'
$btnClear.Location = New-Object System.Drawing.Point(624, 228)
$btnClear.Size     = New-Object System.Drawing.Size(100, 26)
$form.Controls.Add($btnClear)

$lblAfter           = New-Object System.Windows.Forms.Label
$lblAfter.Text      = "After activating, come back and use the Connect buttons. If you were already connected,`r`nreconnect -- the existing session cannot see the new access."
$lblAfter.Location  = New-Object System.Drawing.Point(14, 280)
$lblAfter.Size      = New-Object System.Drawing.Size(710, 36)
$form.Controls.Add($lblAfter)

$lblStatus           = New-Object System.Windows.Forms.Label
$lblStatus.Tag       = 'NoTheme'
$lblStatus.Location  = New-Object System.Drawing.Point(14, 326)
$lblStatus.Size      = New-Object System.Drawing.Size(710, 64)
$lblStatus.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblStatus)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(624, 424)
$btnClose.Size     = New-Object System.Drawing.Size(100, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

function Open-Pim {
    param([Parameter(Mandatory)][string]$Which)
    $url = Get-PimUrl -Which $Which
    try {
        # Purpose 'Admin': PIM is the elevated account by definition, and
        # Start-Process would land in whichever profile Edge used last.
        $st = Start-UrlInBrowserProfile -Url $url -Purpose 'Admin'
        $lblStatus.Text = "$st`r`n$url`r`nWhen you have activated, press Activated below to start the countdown."
    } catch {
        $lblStatus.Text = "Could not open: $($_.Exception.Message)`r`n$url"
    }
}

function Update-PimCountdown {
    $line = Get-PimStatusLine
    if (-not $line) {
        $lblCountdown.Text = 'Nothing recorded yet. Activate in the portal, then press Activated.'
        $lblCountdown.ForeColor = [System.Drawing.Color]::DimGray
        return
    }
    # The recorded time is always shown next to the countdown. A bare number
    # would look authoritative; this one is visibly a note, with its age on it.
    $at = Get-PimActivation
    $stamp = if ($at) { "recorded $($at.ToLocalTime().ToString('HH:mm'))" } else { '' }
    $lblCountdown.Text = "$line`r`n$stamp   (assumes $(Get-PimDurationHours)h; not read from Entra)"
    $lblCountdown.ForeColor = Get-PimStatusColour
}

function Mark-PimActivated {
    <#
        One clock, so either button resets the same one.

        Both buttons say Activated and both do the same thing on purpose: the
        two rows are two ROUTES to activation, not two things to track. Whichever
        you used, what matters afterwards is the same single question -- how long
        until it lapses.
    #>
    param([string]$Route = 'PIM')

    $at = Set-PimActivated
    $until = $at.AddHours((Get-PimDurationHours)).ToLocalTime().ToString('HH:mm')
    $lblStatus.Text = "Activated at $($at.ToLocalTime().ToString('HH:mm')). Expires about $until."
    Update-PimCountdown
    try {
        Write-CompletedLog -Action 'PIM Activate' -Target $env:USERNAME -Result 'Success' `
            -Detail "$Route marked activated; $(Get-PimDurationHours)h assumed, expires ~$until" -Environment 'Cloud'
    } catch { }
}

$btnGroupsElig.Add_Click({   Open-Pim -Which 'GroupsEligible' })
$btnRolesElig.Add_Click({    Open-Pim -Which 'RolesEligible' })
$btnActivated.Add_Click({ Mark-PimActivated })

$btnClear.Add_Click({
    Clear-PimActivation
    $lblStatus.Text = 'Countdown cleared.'
    Update-PimCountdown
})

Update-PimCountdown

# A minute is plenty. The label is the only thing that changes and the work is
# arithmetic on two timestamps -- no probing, no I/O beyond one small file.
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 60000
$timer.Add_Tick({ Update-PimCountdown })
$timer.Start()
$form.Add_FormClosed({ $timer.Stop(); $timer.Dispose() })

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'PIM Activate' | Out-Null

[void]$form.ShowDialog()
$form.Dispose()
#endregion
