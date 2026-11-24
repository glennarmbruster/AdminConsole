# DEVIN REFERENCE CODE, PART 2 — eight more working tools

**THIS FILE IS EIGHT SEPARATE FILES CONCATENATED. DO NOT TREAT IT AS ONE.**

Same rules as `DEVIN-REFERENCE-CODE.md`, which you should read first. Each
section is a complete standalone script from a working production toolkit.
Reproduce them under their own names; the name is in every heading.

---

## Why these eight, and not more

This file is **not** a grab-bag of extra examples. More code to read is not the
same as more help — it adds context to sift and raises the chance of copying a
pattern that does not apply here.

Seven were chosen on one principle: **after file 1, five of the eight tools in
the build scope had no exemplar.** This closes that gap exactly.

| Build order | Tool | Exemplar was in |
|---|---|---|
| 0 | Activate (PIM) | **this file** |
| 1 | Mailbox Health | file 1 |
| 2 | Message Trace | file 1 |
| 3 | Unified Lookup | file 1 |
| 4 | Mailbox Permissions | **this file** |
| 5 | Calendar Permissions | **this file** |
| 6 | Out-of-Office Editor | **this file** |
| 7 | Inbox Rules | **this file** |
| 8 | Group Manager *(front door)* | **this file** |
| 8 | Group Manager Editor | file 1 |

`activate.ps1` is listed as tool 0 because it is built with the console in
Phase 2, not with the tools in Phase 3 — **nothing else works until it has been
used.**

The last two are additions rather than gap-fills, and each earns its place:

- **`MailAttributes.ps1`** is the only file that demonstrates constraint **C11**
  in real code — the proxy-address rewrite that destroys X500 addresses if done
  naively. That is the single most damaging mistake available in this toolkit,
  and prose has not been enough to prevent it before.
- **`ExternalForwarding.ps1`** is a complete CSV report, which no file in part 1
  shows end to end.

**Every tool in the first pass now has a working example.** That is the point of
this file. If you find yourself wanting a ninth, ask first — the answer is
usually that the pattern already exists in one of these.

---

## The rules have not changed

- These are the reference for **conventions**, not a codebase to extend.
- **Do not copy a tool and adjust it.** Write each from its spec, consulting
  these for *how* rather than *what*.
- **Do not port multi-directory machinery.** The source tenant had two AD
  domains; this one has one.
- **Do not port anything you cannot explain.**
- **Read the comments.** Where one contradicts your instinct, that is
  information, not noise.

---

## Contents

| # | File | Lines | Reference for |
|---|---|---|---|
| 1 | `activate.ps1` | 326 | BUILT WITH THE CONSOLE, NOT AMONG THE TOOLS |
| 2 | `MailboxPermissions.ps1` | 347 | TOOL 4 in the build order |
| 3 | `CalendarPermissions.ps1` | 246 | TOOL 5 |
| 4 | `OOFEditor.ps1` | 240 | TOOL 6 |
| 5 | `InboxRules.ps1` | 320 | TOOL 7 |
| 6 | `Group-Manager.ps1` | 732 | TOOL 8, THE FRONT DOOR |
| 7 | `MailAttributes.ps1` | 739 | THE PROXY-ADDRESS REFERENCE, and the most dangerous operation in the toolkit |
| 8 | `ExternalForwarding.ps1` | 405 | A CSV REPORT end to end: the shared report-file controls, what "external" means and why it depends on a complete accepted-domain list (C3 -- a truncated list turns ordinary internal mail into a false exfiltration finding), partial-run labelling, and never writing a zero-row CSV |

---
---


<br>

==============================================================================

# ◆ FILE 1 OF 8 — `activate.ps1`

> **326 lines.** BUILT WITH THE CONSOLE, NOT AMONG THE TOOLS. Admin rights at this tenant are eligible, not standing -- until an operator activates, nothing else works and every tool fails in ways that look like faults. Read the ORDER MATTERS section in its help header: activate FIRST, connect SECOND, because a token issued before activation does not contain access activated afterwards.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```powershell
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
```


<br>

==============================================================================

# ◆ FILE 2 OF 8 — `MailboxPermissions.ps1`

> **347 lines.** TOOL 4 in the build order. Full Access and Send As are DIFFERENT grants, constantly confused -- "I can see the shared mailbox but cannot send from it" is Full Access without Send As. Note also how it handles a Microsoft 365 Group, where access comes from MEMBERSHIP and these cmdlets do not apply at all (C12).
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```powershell
#Requires -Version 5.1
<#
    MailboxPermissions.ps1
    Standalone Mailbox Permissions (Full Access / Send As delegation) tool
    -- Layer 2 tool script.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-BEGIN
# Grant, view or remove access to somebody else's mailbox.
#
# Two different things, often confused, both handled here:
#
#   * Full Access -- open and read the mailbox
#   * Send As -- send mail that appears to come from the mailbox
#
# They are granted separately. "I can see the shared mailbox but can't send from
# it" is Full Access without Send As, and is the most common ticket in this area.
#
# Changes are written to the audit log. Access granted here does not expire --
# remove it when it is no longer needed.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

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

# NO PROMPT BEFORE THE WINDOW. The address box is ON the form -- see
# Add-ToolTargetBar in Common.ps1. Resolve-ToolTarget is not used any more:
# its job was to decide which DIRECTORY an address belongs to, and every read
# and write below goes through Invoke-EXOCommand to Exchange Online, which
# does not need to be told.
#
# An empty address is fine. The window opens and asks.
$script:Target = Get-TargetSmtp -Text $SmtpAddress

#region Actions - mailbox (moved unchanged from Mailbox-Group-Inspector.ps1 lines 671-795)
function Get-MailboxState {
    param([string]$Smtp)
    $state = [ordered]@{
        DisplayName          = $null
        PrimarySmtpAddress   = $null
        Alias                = $null
        RecipientTypeDetails = $null
        Hidden               = $null
        LitigationHold       = $null
        ArchiveEnabled       = $null
        ArchiveStatus        = $null
        MailboxSize          = $null
        MailboxItemCount     = $null
        ArchiveSize          = $null
        ArchiveItemCount     = $null
        FullAccess           = @()
        SendAs               = @()
        Errors               = @()
    }
    try {
        $mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $state.DisplayName          = $mbx.DisplayName
        $state.PrimarySmtpAddress   = $mbx.PrimarySmtpAddress
        $state.Alias                = $mbx.Alias
        $state.RecipientTypeDetails = $mbx.RecipientTypeDetails
        $state.Hidden               = $mbx.HiddenFromAddressListsEnabled
        $state.LitigationHold       = $mbx.LitigationHoldEnabled
        $state.ArchiveEnabled = ($mbx.ArchiveGuid -and $mbx.ArchiveGuid -ne [Guid]::Empty)
        $state.ArchiveStatus  = $mbx.ArchiveStatus
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox (core properties)')
    }
    try {
        $stats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $state.MailboxSize      = $stats.TotalItemSize
        $state.MailboxItemCount = $stats.ItemCount
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics (mailbox size)')
    }
    if ($state.ArchiveEnabled) {
        try {
            $astats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; Archive = $true; ErrorAction = 'Stop' }
            $state.ArchiveSize      = $astats.TotalItemSize
            $state.ArchiveItemCount = $astats.ItemCount
        } catch {
            $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics -Archive (archive size)')
        }
    }
    try {
        $fa = Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object {
                $_.User -notlike 'NT AUTHORITY\SELF' -and
                $_.IsInherited -eq $false -and
                $_.Deny -eq $false -and
                ($_.AccessRights -contains 'FullAccess')
            }
        $state.FullAccess = @($fa | ForEach-Object { $_.User.ToString() })
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxPermission (Full Access)')
    }
    try {
        $sa = Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object {
                $_.Trustee -notlike 'NT AUTHORITY\SELF' -and
                $_.AccessControlType -eq 'Allow'
            }
        $state.SendAs = @($sa | ForEach-Object { $_.Trustee.ToString() })
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-RecipientPermission (Send As)')
    }
    return [PSCustomObject]$state
}
function Format-MailboxState {
    param([string]$Heading, $State)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("--- $Heading (Mailbox) ---")
    [void]$sb.AppendLine("  Display Name        : $($State.DisplayName)")
    [void]$sb.AppendLine("  Primary SMTP        : $($State.PrimarySmtpAddress)")
    [void]$sb.AppendLine("  Alias               : $($State.Alias)")
    [void]$sb.AppendLine("  Mailbox Type        : $($State.RecipientTypeDetails)")
    [void]$sb.AppendLine("  Hidden From GAL     : $($State.Hidden)")
    [void]$sb.AppendLine("  Litigation Hold     : $($State.LitigationHold)")
    [void]$sb.AppendLine("  Mailbox Size        : $($State.MailboxSize)  ($($State.MailboxItemCount) items)")
    [void]$sb.AppendLine("  Archive Enabled     : $($State.ArchiveEnabled)  (ArchiveStatus: $($State.ArchiveStatus))")
    if ($State.ArchiveEnabled) {
        [void]$sb.AppendLine("  Archive Size        : $($State.ArchiveSize)  ($($State.ArchiveItemCount) items)")
    }
    [void]$sb.AppendLine("  Full Access grantees:")
    if ($State.FullAccess.Count -eq 0) { [void]$sb.AppendLine("    (none)") }
    else { $State.FullAccess | ForEach-Object { [void]$sb.AppendLine("    $_") } }
    [void]$sb.AppendLine("  Send As grantees    :")
    if ($State.SendAs.Count -eq 0) { [void]$sb.AppendLine("    (none)") }
    else { $State.SendAs | ForEach-Object { [void]$sb.AppendLine("    $_") } }
    if ($State.Errors.Count -gt 0) {
        [void]$sb.AppendLine("  --- Field errors (partial data above is still valid) ---")
        $State.Errors | ForEach-Object { [void]$sb.AppendLine("  $_") }
    }
    return $sb.ToString()
}
function Grant-FullAccessPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Add-MailboxPermission' -Parameters @{
        Identity = $Smtp; User = $Grantee; AccessRights = 'FullAccess'; InheritanceType = 'All'; ErrorAction = 'Stop'
    } | Out-Null
}
function Revoke-FullAccessPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-MailboxPermission' -Parameters @{
        Identity = $Smtp; User = $Grantee; AccessRights = 'FullAccess'; InheritanceType = 'All'; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Grant-SendAsPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Add-RecipientPermission' -Parameters @{
        Identity = $Smtp; Trustee = $Grantee; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Revoke-SendAsPermission {
    param([string]$Smtp, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-RecipientPermission' -Parameters @{
        Identity = $Smtp; Trustee = $Grantee; AccessRights = 'SendAs'; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
#endregion

#region NEW - wrapper window for mailbox delegation (see .DESCRIPTION above)
function Show-MailboxPermissionsForm {
    # NO PARAMETERS. The mailbox can change while the window is open, so it
    # lives in $script:Target rather than in a parameter fixed at build time.
    # $Environment is still read below; it is the script-level parameter, and
    # it only ever labelled the audit line.
    $permForm = New-Object System.Windows.Forms.Form
    $permForm.Text = 'Mailbox Permissions'
    $permForm.ClientSize = New-Object System.Drawing.Size(764, 400)
    $permForm.StartPosition = "CenterScreen"
    $permForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Multiline = $true
    $txtOut.ReadOnly = $true
    $txtOut.ScrollBars = "Vertical"
    $txtOut.Location = New-Object System.Drawing.Point(12, 12)
    $txtOut.Size = New-Object System.Drawing.Size(740, 240)
    $txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
    $permForm.Controls.Add($txtOut)

    $grpMbx = New-Object System.Windows.Forms.GroupBox
    $grpMbx.Text = "Mailbox Permissions"
    $grpMbx.Location = New-Object System.Drawing.Point(12, 262)
    $grpMbx.Size = New-Object System.Drawing.Size(740, 90)
    $permForm.Controls.Add($grpMbx)
    $lblMbxGrantee = New-Object System.Windows.Forms.Label
    $lblMbxGrantee.Text = "Grantee (email or logon):"
    $lblMbxGrantee.Location = New-Object System.Drawing.Point(10, 22)
    $lblMbxGrantee.AutoSize = $true
    $grpMbx.Controls.Add($lblMbxGrantee)
    $txtMbxGrantee = New-Object System.Windows.Forms.TextBox
    $txtMbxGrantee.Location = New-Object System.Drawing.Point(170, 19)
    $txtMbxGrantee.Size = New-Object System.Drawing.Size(260, 22)
    $txtMbxGrantee.Font = New-Object System.Drawing.Font("Consolas", 9)
    $grpMbx.Controls.Add($txtMbxGrantee)
    $btnFAGrant = New-Object System.Windows.Forms.Button
    $btnFAGrant.Text = "Grant Full Access"
    $btnFAGrant.Location = New-Object System.Drawing.Point(10, 54)
    $btnFAGrant.Size = New-Object System.Drawing.Size(140, 28)
    $grpMbx.Controls.Add($btnFAGrant)
    $btnFARevoke = New-Object System.Windows.Forms.Button
    $btnFARevoke.Text = "Revoke Full Access"
    $btnFARevoke.Location = New-Object System.Drawing.Point(156, 54)
    $btnFARevoke.Size = New-Object System.Drawing.Size(140, 28)
    $grpMbx.Controls.Add($btnFARevoke)
    $btnSAGrant = New-Object System.Windows.Forms.Button
    $btnSAGrant.Text = "Grant Send As"
    $btnSAGrant.Location = New-Object System.Drawing.Point(302, 54)
    $btnSAGrant.Size = New-Object System.Drawing.Size(140, 28)
    $grpMbx.Controls.Add($btnSAGrant)
    $btnSARevoke = New-Object System.Windows.Forms.Button
    $btnSARevoke.Text = "Revoke Send As"
    $btnSARevoke.Location = New-Object System.Drawing.Point(448, 54)
    $btnSARevoke.Size = New-Object System.Drawing.Size(140, 28)
    $grpMbx.Controls.Add($btnSARevoke)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh"
    $btnRefresh.Location = New-Object System.Drawing.Point(12, 360)
    $btnRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $permForm.Controls.Add($btnRefresh)
    $btnPermClose = New-Object System.Windows.Forms.Button
    $btnPermClose.Text = "Close"
    $btnPermClose.Location = New-Object System.Drawing.Point(652, 360)
    $btnPermClose.Size = New-Object System.Drawing.Size(100, 28)
    $permForm.Controls.Add($btnPermClose)

    function Update-StateDisplay {
        if (-not "$script:Target".Trim()) {
            $txtOut.Text = 'Enter a mailbox address above and press Look Up.'
            $permForm.Text = 'Mailbox Permissions'
            return
        }
        $permForm.Text = "Mailbox Permissions: $script:Target"
        try {
            $state = Get-MailboxState -Smtp $script:Target
            $txtOut.Text = Format-MailboxState -Heading "CURRENT STATE" -State $state
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mailbox Permissions'
        }
    }

    function Invoke-MailboxEdit {
        param([string]$ActionLabel, [string]$Verb, [scriptblock]$Action)
        $grantee = $txtMbxGrantee.Text.Trim()
        if (-not $grantee) { $txtOut.Text = "Enter a grantee (email or logon) first."; return }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "$ActionLabel for '$grantee' on mailbox '$script:Target'?",
            "Confirm $ActionLabel",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { $txtOut.Text = "Cancelled."; return }
        $grpMbx.Enabled = $false
        try {
            $before = Get-MailboxState -Smtp $script:Target
            $out = New-Object System.Text.StringBuilder
            [void]$out.Append((Format-MailboxState -Heading "BEFORE" -State $before))
            & $Action
            [void]$out.AppendLine("$ActionLabel : SUCCESS ($grantee)")
            $after = Get-MailboxState -Smtp $script:Target
            [void]$out.Append((Format-MailboxState -Heading "AFTER" -State $after))
            $txtOut.Text = $out.ToString()
            # 'Cloud' rather than an empty pair of brackets.
            #
            # $Environment comes from Resolve-ToolTarget and is now NULL for a
            # cloud-only object, which -CloudOk made possible. Left as-is the
            # audit line read "(; grantee: ...)" -- a gap that looks like a bug
            # in the logger rather than the true statement "this object has no
            # on-premises directory".
            $envLabel = if ("$Environment".Trim()) { "$Environment" } else { 'Cloud' }
            Write-CompletedLog -Task "$Verb $script:Target ($envLabel; grantee: $grantee)"
        }
        catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context $ActionLabel
        }
        finally {
            $grpMbx.Enabled = $true
        }
    }

    Update-StateDisplay
    $btnRefresh.Add_Click({ Update-StateDisplay })
    $btnFAGrant.Add_Click({
        Invoke-MailboxEdit -ActionLabel "Grant Full Access" -Verb "Granted FullAccess" -Action {
            Grant-FullAccessPermission -Smtp $script:Target -Grantee $txtMbxGrantee.Text.Trim()
        }
    })
    $btnFARevoke.Add_Click({
        Invoke-MailboxEdit -ActionLabel "Revoke Full Access" -Verb "Revoked FullAccess" -Action {
            Revoke-FullAccessPermission -Smtp $script:Target -Grantee $txtMbxGrantee.Text.Trim()
        }
    })
    $btnSAGrant.Add_Click({
        Invoke-MailboxEdit -ActionLabel "Grant Send As" -Verb "Granted SendAs" -Action {
            Grant-SendAsPermission -Smtp $script:Target -Grantee $txtMbxGrantee.Text.Trim()
        }
    })
    $btnSARevoke.Add_Click({
        Invoke-MailboxEdit -ActionLabel "Revoke Send As" -Verb "Revoked SendAs" -Action {
            Revoke-SendAsPermission -Smtp $script:Target -Grantee $txtMbxGrantee.Text.Trim()
        }
    })
    $btnPermClose.Add_Click({ $permForm.Close() })

    # BEFORE Add-ToolHelpButton: the bar shifts every existing control down and
    # grows the form, and the help button places itself from ClientSize at run
    # time. Before Set-FormScreenCorner too, which positions from the size.
    $bar = Add-ToolTargetBar -Form $permForm -InitialValue $script:Target

    # WIRED HERE, NOT PASSED INTO THE HELPER. Update-StateDisplay is nested in
    # this function, so the handler has to be built in this scope to see it.
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        Update-StateDisplay
    })

    Set-ConsoleTheme -Form $permForm
    Set-FormScreenCorner -Form $permForm -Corner 'BottomLeft'
    # Consistent '?' in the top-right of every tool. Added after the theme so
    # it is not restyled, and BringToFront'd inside the helper because
    # WinForms draws EARLIER-added controls on top.
    Add-ToolHelpButton -Form $permForm -ScriptPath $PSCommandPath -Title 'Mailbox Permissions' | Out-Null

    [void]$permForm.ShowDialog()
}
#endregion

Show-MailboxPermissionsForm
```


<br>

==============================================================================

# ◆ FILE 3 OF 8 — `CalendarPermissions.ps1`

> **246 lines.** TOOL 5. Folder-level permissions. Note that the calendar folder path is LOCALISED and must be resolved rather than hardcoded, and that Default and Anonymous are pseudo-users that always exist and can only be set to None, never removed.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```powershell
#Requires -Version 5.1
<#
    CalendarPermissions.ps1
    Standalone Calendar Permissions tool -- Layer 2 tool script.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-BEGIN
# See and change who can view or edit someone's calendar.
#
# Shows the current permission for each person who has one, and lets you add,
# change or remove access at the usual levels -- from free/busy only through to
# full editing.
#
# The most common ticket this answers is "my assistant can't see my calendar"
# or "someone can see more of my calendar than they should".
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

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

# NO PROMPT BEFORE THE WINDOW. The address box is ON the form -- see
# Add-ToolTargetBar in Common.ps1. An empty address is fine; the window opens
# and asks for one.
$script:Target = Get-TargetSmtp -Text $SmtpAddress

#region Actions - calendar permissions (moved unchanged from Mailbox-Group-Inspector.ps1 lines 1272-1455)
function Get-CalendarFolderIdentity {
    param([string]$Smtp)
    $folders = @(Invoke-EXOCommand -Name 'Get-MailboxFolderStatistics' -Parameters @{ Identity = $Smtp; FolderScope = 'Calendar'; ErrorAction = 'Stop' })
    if ($folders.Count -eq 0) { throw "No Calendar-scope folder found for $Smtp." }
    $primary = $folders | Where-Object { $_.FolderPath -eq '/Calendar' } | Select-Object -First 1
    if (-not $primary) { $primary = $folders | Select-Object -First 1 }
    return "$($Smtp):\$($primary.Name)"
}
function Get-CalendarPermissions {
    param([string]$Smtp)
    try {
        $calId = Get-CalendarFolderIdentity -Smtp $Smtp
        $perms = @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' -Parameters @{ Identity = $calId; ErrorAction = 'Stop' })
        return [PSCustomObject]@{ CalendarIdentity = $calId; Permissions = $perms; Error = $null }
    } catch {
        return [PSCustomObject]@{ CalendarIdentity = $null; Permissions = @(); Error = (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxFolderPermission (Calendar)') }
    }
}
function Set-CalendarPermission {
    param([string]$CalendarIdentity, [string]$Grantee, [string]$Role)
    $existing = @()
    try {
        $existing = @(Invoke-EXOCommand -Name 'Get-MailboxFolderPermission' -Parameters @{ Identity = $CalendarIdentity; User = $Grantee; ErrorAction = 'Stop' })
    } catch {
        $existing = @()
    }
    if ($existing.Count -gt 0) {
        Invoke-EXOCommand -Name 'Set-MailboxFolderPermission' -Parameters @{
            Identity = $CalendarIdentity; User = $Grantee; AccessRights = $Role; Confirm = $false; ErrorAction = 'Stop'
        } | Out-Null
        return 'Updated'
    } else {
        Invoke-EXOCommand -Name 'Add-MailboxFolderPermission' -Parameters @{
            Identity = $CalendarIdentity; User = $Grantee; AccessRights = $Role; Confirm = $false; ErrorAction = 'Stop'
        } | Out-Null
        return 'Added'
    }
}
function Remove-CalendarPermission {
    param([string]$CalendarIdentity, [string]$Grantee)
    Invoke-EXOCommand -Name 'Remove-MailboxFolderPermission' -Parameters @{
        Identity = $CalendarIdentity; User = $Grantee; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Show-CalendarPermissionsForm {
    # NO -Smtp PARAMETER. The address can change while the window is open, so it
    # lives in $script:Target rather than a parameter fixed at build time.
    $calForm = New-Object System.Windows.Forms.Form
    $calForm.Text = "Calendar Permissions: $script:Target"
    $calForm.ClientSize = New-Object System.Drawing.Size(604, 440)
    $calForm.StartPosition = 'Manual'
    $calForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblCalId = New-Object System.Windows.Forms.Label
    $lblCalId.Text = "Calendar folder:"
    $lblCalId.Location = New-Object System.Drawing.Point(12, 12)
    $lblCalId.AutoSize = $true
    $calForm.Controls.Add($lblCalId)
    $grdCal = New-Object System.Windows.Forms.DataGridView
    $grdCal.Location = New-Object System.Drawing.Point(12, 34)
    $grdCal.Size = New-Object System.Drawing.Size(580, 280)
    $grdCal.ReadOnly = $true
    $grdCal.AllowUserToAddRows = $false
    $grdCal.AllowUserToDeleteRows = $false
    $grdCal.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdCal.MultiSelect = $false
    $grdCal.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grdCal.Font = New-Object System.Drawing.Font("Consolas", 9)
    [void]$grdCal.Columns.Add('User', 'User')
    [void]$grdCal.Columns.Add('AccessRights', 'Access Rights')
    foreach ($c in $grdCal.Columns) { $c.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable }
    $calForm.Controls.Add($grdCal)
    $lblGrantee = New-Object System.Windows.Forms.Label
    $lblGrantee.Text = "User (email or logon):"
    $lblGrantee.Location = New-Object System.Drawing.Point(12, 326)
    $lblGrantee.AutoSize = $true
    $calForm.Controls.Add($lblGrantee)
    $txtGrantee = New-Object System.Windows.Forms.TextBox
    $txtGrantee.Location = New-Object System.Drawing.Point(160, 323)
    $txtGrantee.Size = New-Object System.Drawing.Size(220, 20)
    $txtGrantee.Font = New-Object System.Drawing.Font("Consolas", 9)
    $calForm.Controls.Add($txtGrantee)
    $lblRole = New-Object System.Windows.Forms.Label
    $lblRole.Text = "Role:"
    $lblRole.Location = New-Object System.Drawing.Point(390, 326)
    $lblRole.AutoSize = $true
    $calForm.Controls.Add($lblRole)
    $cboRole = New-Object System.Windows.Forms.ComboBox
    $cboRole.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('Reviewer', 'Editor', 'Owner') | ForEach-Object { [void]$cboRole.Items.Add($_) }
    $cboRole.SelectedIndex = 0
    $cboRole.Location = New-Object System.Drawing.Point(430, 323)
    $cboRole.Size = New-Object System.Drawing.Size(120, 20)
    $calForm.Controls.Add($cboRole)
    $lblCalStatus = New-Object System.Windows.Forms.Label
    $lblCalStatus.Text = ""
    $lblCalStatus.Location = New-Object System.Drawing.Point(12, 356)
    $lblCalStatus.Size = New-Object System.Drawing.Size(580, 40)
    $calForm.Controls.Add($lblCalStatus)
    $btnAddPerm = New-Object System.Windows.Forms.Button
    $btnAddPerm.Text = "Add / Update Permission"
    $btnAddPerm.Location = New-Object System.Drawing.Point(12, 400)
    $btnAddPerm.Size = New-Object System.Drawing.Size(170, 28)
    $calForm.Controls.Add($btnAddPerm)
    $btnRemovePerm = New-Object System.Windows.Forms.Button
    $btnRemovePerm.Text = "Remove Selected"
    $btnRemovePerm.Location = New-Object System.Drawing.Point(190, 400)
    $btnRemovePerm.Size = New-Object System.Drawing.Size(150, 28)
    $calForm.Controls.Add($btnRemovePerm)
    $btnCalClose = New-Object System.Windows.Forms.Button
    $btnCalClose.Text = "Close"
    $btnCalClose.Location = New-Object System.Drawing.Point(500, 400)
    $btnCalClose.Size = New-Object System.Drawing.Size(90, 28)
    $calForm.Controls.Add($btnCalClose)
    $script:CalendarIdentityForForm = $null
    function Update-CalGridFromCurrent {
        if (-not "$script:Target".Trim()) {
            $calForm.Text = 'Calendar Permissions'
            return
        }
        $calForm.Text = "Calendar Permissions: $script:Target"
        $grdCal.Rows.Clear()
        $result = Get-CalendarPermissions -Smtp $script:Target
        if ($result.Error) {
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Red
            $lblCalStatus.Text = $result.Error
            $lblCalId.Text = "Calendar folder: (unresolved)"
            return
        }
        $script:CalendarIdentityForForm = $result.CalendarIdentity
        $lblCalId.Text = "Calendar folder: $($result.CalendarIdentity)"
        foreach ($p in $result.Permissions) {
            [void]$grdCal.Rows.Add($p.User.ToString(), ($p.AccessRights -join ', '))
        }
    }
    Update-CalGridFromCurrent
    $btnAddPerm.Add_Click({
        $grantee = $txtGrantee.Text.Trim()
        if (-not $grantee) { $lblCalStatus.Text = "Enter a user first."; return }
        if (-not $script:CalendarIdentityForForm) { $lblCalStatus.Text = "Calendar folder not resolved -- see status above."; return }
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Grant '$($cboRole.SelectedItem)' calendar access to '$grantee' on '$script:Target'?",
            "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnAddPerm.Enabled = $false; $btnRemovePerm.Enabled = $false
        try {
            $action = Set-CalendarPermission -CalendarIdentity $script:CalendarIdentityForForm -Grantee $grantee -Role $cboRole.SelectedItem
            Write-CompletedLog -Task "$action CalendarPermission $script:Target (EXO; user: $grantee; role: $($cboRole.SelectedItem))"
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Black
            $lblCalStatus.Text = "$action permission for $grantee."
            Update-CalGridFromCurrent
        } catch {
            Write-FailedLog -Action 'GrantCalendarPermission' -Target $script:Target -ErrorRecord $_ -Detail "user: $grantee; role: $($cboRole.SelectedItem)"
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Red
            $lblCalStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Calendar permission grant'
        } finally {
            $btnAddPerm.Enabled = $true; $btnRemovePerm.Enabled = $true
        }
    })
    $btnRemovePerm.Add_Click({
        if ($grdCal.SelectedRows.Count -eq 0) { $lblCalStatus.Text = "Select a row first."; return }
        $grantee = $grdCal.SelectedRows[0].Cells['User'].Value
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Remove all calendar permissions for '$grantee' on '$script:Target'?",
            "Confirm Remove", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnAddPerm.Enabled = $false; $btnRemovePerm.Enabled = $false
        try {
            Remove-CalendarPermission -CalendarIdentity $script:CalendarIdentityForForm -Grantee $grantee
            Write-CompletedLog -Task "Removed CalendarPermission $script:Target (EXO; user: $grantee)"
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Black
            $lblCalStatus.Text = "Removed permission for $grantee."
            Update-CalGridFromCurrent
        } catch {
            Write-FailedLog -Action 'RemoveCalendarPermission' -Target $script:Target -ErrorRecord $_ -Detail "user: $grantee"
            $lblCalStatus.ForeColor = [System.Drawing.Color]::Red
            $lblCalStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove-MailboxFolderPermission (Calendar)'
        } finally {
            $btnAddPerm.Enabled = $true; $btnRemovePerm.Enabled = $true
        }
    })
    $btnCalClose.Add_Click({ $calForm.Close() })
    $calForm.StartPosition = 'Manual'
    # BEFORE the theme and the help button: it shifts every existing control
    # down, so anything added after would be left at the old offset.
    $bar = Add-ToolTargetBar -Form $calForm -InitialValue $script:Target
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        Update-CalGridFromCurrent
    })

    Set-ConsoleTheme -Form $calForm
    Set-FormScreenCorner -Form $calForm -Corner 'BottomLeft'
    # Consistent '?' in the top-right of every tool. Added after the theme so
    # it is not restyled, and BringToFront'd inside the helper because
    # WinForms draws EARLIER-added controls on top.
    Add-ToolHelpButton -Form $calForm -ScriptPath $PSCommandPath -Title 'Calendar Permissions' | Out-Null

    [void]$calForm.ShowDialog()
}
#endregion

Show-CalendarPermissionsForm
```


<br>

==============================================================================

# ◆ FILE 4 OF 8 — `OOFEditor.ps1`

> **240 lines.** TOOL 6. Note that there are TWO messages -- internal and external -- usually different, and usually only one of them is wrong. Both are always shown. A Scheduled window with a past end date reads as "on" in some views while sending nothing.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```powershell
#Requires -Version 5.1
<#
    OOFEditor.ps1
    Standalone Out of Office tool -- Layer 2 tool script.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-BEGIN
# Read and set someone's out-of-office message without needing their mailbox
# open.
#
# Shows whether auto-reply is on, the scheduled window if there is one, and both
# messages -- the internal one and the one external senders receive, which are
# often different and only one of which is usually wrong.
#
# For the classic ticket: somebody left, or is off sick, and their auto-reply
# needs setting or clearing on their behalf.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

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

# NO PROMPT BEFORE THE WINDOW. The address box is ON the form -- see
# Add-ToolTargetBar in Common.ps1. An empty address is fine; the window
# opens and asks for one.
$script:Target = Get-TargetSmtp -Text $SmtpAddress

#region Actions - out of office (moved unchanged from Mailbox-Group-Inspector.ps1 lines 1108-1269)
function Get-OutOfOfficeConfig {
    param([string]$Smtp)
    Invoke-EXOCommand -Name 'Get-MailboxAutoReplyConfiguration' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
}
function Set-OutOfOfficeConfig {
    param(
        [string]$Smtp,
        [string]$AutoReplyState,
        [string]$ExternalAudience,
        [string]$InternalMessage,
        [string]$ExternalMessage,
        [Nullable[DateTime]]$StartTime,
        [Nullable[DateTime]]$EndTime
    )
    $params = @{
        Identity         = $Smtp
        AutoReplyState   = $AutoReplyState
        ExternalAudience = $ExternalAudience
        InternalMessage  = $InternalMessage
        ExternalMessage  = $ExternalMessage
        Confirm          = $false
        ErrorAction      = 'Stop'
    }
    if ($AutoReplyState -eq 'Scheduled') {
        if ($StartTime) { $params.StartTime = $StartTime }
        if ($EndTime)   { $params.EndTime   = $EndTime }
    }
    Invoke-EXOCommand -Name 'Set-MailboxAutoReplyConfiguration' -Parameters $params | Out-Null
}
function Show-OutOfOfficeForm {
    # NO -Smtp PARAMETER. The address can change while the window is open, so
    # it lives in $script:Target rather than a parameter fixed at build time.
    $oofForm = New-Object System.Windows.Forms.Form
    $oofForm.Text = "Out of Office: $script:Target"
    $oofForm.ClientSize = New-Object System.Drawing.Size(544, 514)
    $oofForm.StartPosition = 'Manual'
    $oofForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblState = New-Object System.Windows.Forms.Label
    $lblState.Text = "State:"
    $lblState.Location = New-Object System.Drawing.Point(12, 15)
    $lblState.AutoSize = $true
    $oofForm.Controls.Add($lblState)
    $cboState = New-Object System.Windows.Forms.ComboBox
    $cboState.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('Disabled', 'Enabled', 'Scheduled') | ForEach-Object { [void]$cboState.Items.Add($_) }
    $cboState.Location = New-Object System.Drawing.Point(140, 12)
    $cboState.Size = New-Object System.Drawing.Size(150, 20)
    $oofForm.Controls.Add($cboState)
    $lblAudience = New-Object System.Windows.Forms.Label
    $lblAudience.Text = "External audience:"
    $lblAudience.Location = New-Object System.Drawing.Point(300, 15)
    $lblAudience.AutoSize = $true
    $oofForm.Controls.Add($lblAudience)
    $cboAudience = New-Object System.Windows.Forms.ComboBox
    $cboAudience.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('None', 'Known', 'All') | ForEach-Object { [void]$cboAudience.Items.Add($_) }
    $cboAudience.Location = New-Object System.Drawing.Point(420, 12)
    $cboAudience.Size = New-Object System.Drawing.Size(100, 20)
    $oofForm.Controls.Add($cboAudience)
    $lblStart = New-Object System.Windows.Forms.Label
    $lblStart.Text = "Start (Scheduled only, e.g. 2026-07-10 08:00):"
    $lblStart.Location = New-Object System.Drawing.Point(12, 45)
    $lblStart.AutoSize = $true
    $oofForm.Controls.Add($lblStart)
    $txtStart = New-Object System.Windows.Forms.TextBox
    $txtStart.Location = New-Object System.Drawing.Point(280, 42)
    $txtStart.Size = New-Object System.Drawing.Size(240, 20)
    $txtStart.Font = New-Object System.Drawing.Font("Consolas", 9)
    $oofForm.Controls.Add($txtStart)
    $lblEnd = New-Object System.Windows.Forms.Label
    $lblEnd.Text = "End (Scheduled only):"
    $lblEnd.Location = New-Object System.Drawing.Point(12, 70)
    $lblEnd.AutoSize = $true
    $oofForm.Controls.Add($lblEnd)
    $txtEnd = New-Object System.Windows.Forms.TextBox
    $txtEnd.Location = New-Object System.Drawing.Point(280, 67)
    $txtEnd.Size = New-Object System.Drawing.Size(240, 20)
    $txtEnd.Font = New-Object System.Drawing.Font("Consolas", 9)
    $oofForm.Controls.Add($txtEnd)
    $lblInt = New-Object System.Windows.Forms.Label
    $lblInt.Text = "Internal message:"
    $lblInt.Location = New-Object System.Drawing.Point(12, 98)
    $lblInt.AutoSize = $true
    $oofForm.Controls.Add($lblInt)
    $txtInt = New-Object System.Windows.Forms.TextBox
    $txtInt.Multiline = $true
    $txtInt.ScrollBars = "Vertical"
    $txtInt.Location = New-Object System.Drawing.Point(12, 118)
    $txtInt.Size = New-Object System.Drawing.Size(520, 140)
    $txtInt.Font = New-Object System.Drawing.Font("Consolas", 9)
    $oofForm.Controls.Add($txtInt)
    $lblExt = New-Object System.Windows.Forms.Label
    $lblExt.Text = "External message:"
    $lblExt.Location = New-Object System.Drawing.Point(12, 264)
    $lblExt.AutoSize = $true
    $oofForm.Controls.Add($lblExt)
    $txtExt = New-Object System.Windows.Forms.TextBox
    $txtExt.Multiline = $true
    $txtExt.ScrollBars = "Vertical"
    $txtExt.Location = New-Object System.Drawing.Point(12, 284)
    $txtExt.Size = New-Object System.Drawing.Size(520, 140)
    $txtExt.Font = New-Object System.Drawing.Font("Consolas", 9)
    $oofForm.Controls.Add($txtExt)
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = ""
    $lblStatus.Location = New-Object System.Drawing.Point(12, 430)
    $lblStatus.Size = New-Object System.Drawing.Size(520, 40)
    $oofForm.Controls.Add($lblStatus)
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Save"
    $btnSave.Location = New-Object System.Drawing.Point(12, 474)
    $btnSave.Size = New-Object System.Drawing.Size(100, 28)
    $oofForm.Controls.Add($btnSave)
    $btnOofClose = New-Object System.Windows.Forms.Button
    $btnOofClose.Text = "Close"
    $btnOofClose.Location = New-Object System.Drawing.Point(120, 474)
    $btnOofClose.Size = New-Object System.Drawing.Size(100, 28)
    $oofForm.Controls.Add($btnOofClose)
    function Update-OofFormFromCurrent {
        # OPENED WITH NO ADDRESS IS A NORMAL STATE NOW -- the window appears
        # before anything has been typed, so the refresh must not run empty.
        if (-not "$script:Target".Trim()) {
            $oofForm.Text = 'Out of Office'
            return
        }
        $oofForm.Text = 'Out of Office: ' + $script:Target
        try {
            $cfg = Get-OutOfOfficeConfig -Smtp $script:Target
            $cboState.SelectedItem    = $cfg.AutoReplyState.ToString()
            $cboAudience.SelectedItem = $cfg.ExternalAudience.ToString()
            $txtStart.Text = if ($cfg.StartTime) { $cfg.StartTime.ToString() } else { '' }
            $txtEnd.Text   = if ($cfg.EndTime)   { $cfg.EndTime.ToString() }   else { '' }
            $txtInt.Text   = $cfg.InternalMessage
            $txtExt.Text   = $cfg.ExternalMessage
            $lblStatus.ForeColor = [System.Drawing.Color]::Black
            $lblStatus.Text = "Loaded current settings."
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxAutoReplyConfiguration'
        }
    }
    Update-OofFormFromCurrent
    $btnSave.Add_Click({
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Save Out of Office settings for '$script:Target'? State: $($cboState.SelectedItem)",
            "Confirm Save", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnSave.Enabled = $false
        try {
            $st = $null; $et = $null
            if ($cboState.SelectedItem -eq 'Scheduled') {
                if ($txtStart.Text.Trim()) { $st = [DateTime]::Parse($txtStart.Text.Trim()) }
                if ($txtEnd.Text.Trim())   { $et = [DateTime]::Parse($txtEnd.Text.Trim()) }
            }
            Set-OutOfOfficeConfig -Smtp $script:Target -AutoReplyState $cboState.SelectedItem -ExternalAudience $cboAudience.SelectedItem `
                -InternalMessage $txtInt.Text -ExternalMessage $txtExt.Text -StartTime $st -EndTime $et
            Write-CompletedLog -Task "Set OutOfOffice $script:Target (EXO; state: $($cboState.SelectedItem))"
            $lblStatus.ForeColor = [System.Drawing.Color]::Black
            $lblStatus.Text = "Saved successfully."
            Update-OofFormFromCurrent
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Set-MailboxAutoReplyConfiguration'
        } finally {
            $btnSave.Enabled = $true
        }
    })
    $btnOofClose.Add_Click({ $oofForm.Close() })
    $oofForm.StartPosition = 'Manual'
    Set-ConsoleTheme -Form $oofForm
    # Consistent '?' in the top-right of every tool. Added after the theme so
    # it is not restyled, and BringToFront'd inside the helper because
    # WinForms draws EARLIER-added controls on top.
    # BEFORE Add-ToolHelpButton: the bar shifts every existing control down
    # and grows the form, and the help button places itself from ClientSize at
    # run time.
    $bar = Add-ToolTargetBar -Form $oofForm -Label 'Mailbox or address:' -InitialValue $script:Target
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        Update-OofFormFromCurrent
    })

    Add-ToolHelpButton -Form $oofForm -ScriptPath $PSCommandPath -Title 'Out of Office' | Out-Null

    # AFTER the bar, not before. Add-ToolTargetBar grows the form by the
    # height of the address row, so a corner pinned first leaves the window
    # hanging that far past the edge of the working area -- which on the
    # bottom-left corner means underneath the taskbar.
    Set-FormScreenCorner -Form $oofForm -Corner 'BottomLeft'

    [void]$oofForm.ShowDialog()
}
#endregion

Show-OutOfOfficeForm
```


<br>

==============================================================================

# ◆ FILE 5 OF 8 — `InboxRules.ps1`

> **320 lines.** TOOL 7. Deleting a rule cannot be undone; disabling is reversible, so disable is the default and delete is the deliberate choice. Rule names are frequently blank and are not unique -- bind carefully. A rule forwarding externally is a security finding, not a nuisance.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```powershell
#Requires -Version 5.1
<#
    InboxRules.ps1
    Standalone Inbox Rules tool -- Layer 2 tool script.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-BEGIN
# View, disable, enable and delete the inbox rules on someone's mailbox.
#
# The usual reasons to look: mail vanishing into a folder nobody remembers
# creating, mail being forwarded somewhere it should not be, or a rule left
# behind by a compromise.
#
# Deleting a rule cannot be undone -- disabling is reversible, so prefer that
# while you are still working out whether a rule is wanted.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

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

# NO PROMPT BEFORE THE WINDOW. The address box is ON the form -- see
# Add-ToolTargetBar in Common.ps1. Resolve-ToolTarget is not used any more:
# its job was to decide which DIRECTORY an address belongs to, and every read
# and write below goes through Invoke-EXOCommand to Exchange Online, which
# does not need to be told.
#
# An empty address is fine. The window opens and asks.
$script:Target = Get-TargetSmtp -Text $SmtpAddress

#region Actions - inbox rules (moved unchanged from Mailbox-Group-Inspector.ps1 lines 807-982; Get-InboxRulesForMailbox is in Common.ps1)
function Enable-TargetInboxRule {
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Enable-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Disable-TargetInboxRule {
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Disable-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Remove-TargetInboxRule {
    param([string]$Smtp, [string]$RuleIdentity)
    Invoke-EXOCommand -Name 'Remove-InboxRule' -Parameters @{
        Identity = $RuleIdentity; Confirm = $false; ErrorAction = 'Stop'
    } | Out-Null
}
function Update-RulesGrid {
    # STILL TAKES -Smtp. This one is a script-scope function that reaches up
    # into the form's controls, so it is not the place to introduce a second
    # source of truth for the address -- the caller passes $script:Target.
    param([string]$Smtp)
    $grdRules.Rows.Clear()
    if (-not "$Smtp".Trim()) {
        $txtOut.Text = 'Enter a mailbox address above and press Look Up.'
        $script:CurrentRules = @()
        return
    }
    $result = Get-InboxRulesForMailbox -Smtp $Smtp
    if ($result.Error) {
        $txtOut.Text = $result.Error
        $script:CurrentRules = @()
        return
    }
    $script:CurrentRules = $result.Rules
    foreach ($r in $result.Rules) {
        $desc = if ($r.Description) { ($r.Description -replace '\s+', ' ').Trim() } else { '' }
        [void]$grdRules.Rows.Add($r.Name, $desc, $r.Enabled, $r.Identity.ToString())
    }
}
function Show-RuleDetailForm {
    param($Rule, [string]$Smtp)
    $ruleForm = New-Object System.Windows.Forms.Form
    $ruleForm.Text = "Inbox Rule: $($Rule.Name)"
    $ruleForm.ClientSize = New-Object System.Drawing.Size(544, 384)
    $ruleForm.StartPosition = "CenterParent"
    $ruleForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Text = "Rule Name : $($Rule.Name)"
    $lblName.Location = New-Object System.Drawing.Point(12, 12)
    $lblName.AutoSize = $true
    $ruleForm.Controls.Add($lblName)
    $lblMailbox = New-Object System.Windows.Forms.Label
    $lblMailbox.Text = "Mailbox   : $Smtp"
    $lblMailbox.Location = New-Object System.Drawing.Point(12, 32)
    $lblMailbox.AutoSize = $true
    $ruleForm.Controls.Add($lblMailbox)
    $lblEnabled = New-Object System.Windows.Forms.Label
    $lblEnabled.Text = "Enabled   : $($Rule.Enabled)"
    $lblEnabled.Location = New-Object System.Drawing.Point(12, 52)
    $lblEnabled.AutoSize = $true
    $lblEnabled.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $ruleForm.Controls.Add($lblEnabled)
    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = "Full Description:"
    $lblDesc.Location = New-Object System.Drawing.Point(12, 78)
    $lblDesc.AutoSize = $true
    $ruleForm.Controls.Add($lblDesc)
    $txtDesc = New-Object System.Windows.Forms.TextBox
    $txtDesc.Multiline = $true
    $txtDesc.ReadOnly = $true
    $txtDesc.ScrollBars = "Vertical"
    $txtDesc.Location = New-Object System.Drawing.Point(12, 98)
    $txtDesc.Size = New-Object System.Drawing.Size(520, 200)
    $txtDesc.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtDesc.Text = $Rule.Description
    $ruleForm.Controls.Add($txtDesc)
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = ""
    $lblStatus.Location = New-Object System.Drawing.Point(12, 304)
    $lblStatus.Size = New-Object System.Drawing.Size(520, 34)
    $ruleForm.Controls.Add($lblStatus)
    $btnEnable = New-Object System.Windows.Forms.Button
    $btnEnable.Text = "Enable"
    $btnEnable.Location = New-Object System.Drawing.Point(12, 344)
    $btnEnable.Size = New-Object System.Drawing.Size(100, 28)
    $ruleForm.Controls.Add($btnEnable)
    $btnDisable = New-Object System.Windows.Forms.Button
    $btnDisable.Text = "Disable"
    $btnDisable.Location = New-Object System.Drawing.Point(120, 344)
    $btnDisable.Size = New-Object System.Drawing.Size(100, 28)
    $ruleForm.Controls.Add($btnDisable)
    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Text = "Delete"
    $btnDelete.Location = New-Object System.Drawing.Point(228, 344)
    $btnDelete.Size = New-Object System.Drawing.Size(100, 28)
    $ruleForm.Controls.Add($btnDelete)
    $btnRuleClose = New-Object System.Windows.Forms.Button
    $btnRuleClose.Text = "Close"
    $btnRuleClose.Location = New-Object System.Drawing.Point(432, 344)
    $btnRuleClose.Size = New-Object System.Drawing.Size(100, 28)
    $ruleForm.Controls.Add($btnRuleClose)
    $ruleIdentity = $Rule.Identity.ToString()
    $btnEnable.Add_Click({
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Enable Inbox rule '$($Rule.Name)' on mailbox '$Smtp'?",
            "Confirm Enable", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnEnable.Enabled = $false; $btnDisable.Enabled = $false; $btnDelete.Enabled = $false
        try {
            Enable-TargetInboxRule -Smtp $Smtp -RuleIdentity $ruleIdentity
            $lblStatus.ForeColor = [System.Drawing.Color]::Black
            $lblStatus.Text = "Enabled successfully."
            $lblEnabled.Text = "Enabled   : True"
            Write-CompletedLog -Task "Enabled InboxRule '$($Rule.Name)' $Smtp (EXO)"
            Update-RulesGrid -Smtp $Smtp
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Enable-InboxRule'
        } finally {
            $btnEnable.Enabled = $true; $btnDisable.Enabled = $true; $btnDelete.Enabled = $true
        }
    })
    $btnDisable.Add_Click({
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Disable Inbox rule '$($Rule.Name)' on mailbox '$Smtp'?",
            "Confirm Disable", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnEnable.Enabled = $false; $btnDisable.Enabled = $false; $btnDelete.Enabled = $false
        try {
            Disable-TargetInboxRule -Smtp $Smtp -RuleIdentity $ruleIdentity
            $lblStatus.ForeColor = [System.Drawing.Color]::Black
            $lblStatus.Text = "Disabled successfully."
            $lblEnabled.Text = "Enabled   : False"
            Write-CompletedLog -Task "Disabled InboxRule '$($Rule.Name)' $Smtp (EXO)"
            Update-RulesGrid -Smtp $Smtp
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Disable-InboxRule'
        } finally {
            $btnEnable.Enabled = $true; $btnDisable.Enabled = $true; $btnDelete.Enabled = $true
        }
    })
    $btnDelete.Add_Click({
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Permanently delete Inbox rule '$($Rule.Name)' on mailbox '$Smtp'? This cannot be undone.",
            "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnEnable.Enabled = $false; $btnDisable.Enabled = $false; $btnDelete.Enabled = $false
        try {
            Remove-TargetInboxRule -Smtp $Smtp -RuleIdentity $ruleIdentity
            Write-CompletedLog -Task "Deleted InboxRule '$($Rule.Name)' $Smtp (EXO)"
            Update-RulesGrid -Smtp $Smtp
            [System.Windows.Forms.MessageBox]::Show("Rule deleted.", "Deleted", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            $ruleForm.Close()
        } catch {
            $lblStatus.ForeColor = [System.Drawing.Color]::Red
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove-InboxRule'
            $btnEnable.Enabled = $true; $btnDisable.Enabled = $true; $btnDelete.Enabled = $true
        }
    })
    $btnRuleClose.Add_Click({ $ruleForm.Close() })
    [void]$ruleForm.ShowDialog()
}
#endregion

#region NEW - wrapper window for the grid (see .DESCRIPTION above)
function Show-InboxRulesForm {
    # NO -Smtp PARAMETER. The mailbox can change while the window is open, so
    # it lives in $script:Target rather than a parameter fixed at build time.
    $rulesForm = New-Object System.Windows.Forms.Form
    $rulesForm.Text = 'Inbox Rules'
    $rulesForm.ClientSize = New-Object System.Drawing.Size(764, 468)
    $rulesForm.StartPosition = "CenterScreen"
    $rulesForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

    $lblRules = New-Object System.Windows.Forms.Label
    $lblRules.Text = "Inbox Rules (double-click a row to open):"
    $lblRules.Location = New-Object System.Drawing.Point(12, 12)
    $lblRules.AutoSize = $true
    $rulesForm.Controls.Add($lblRules)

    $grdRules = New-Object System.Windows.Forms.DataGridView
    $grdRules.Location = New-Object System.Drawing.Point(12, 32)
    $grdRules.Size = New-Object System.Drawing.Size(740, 340)
    $grdRules.ReadOnly = $true
    $grdRules.AllowUserToAddRows = $false
    $grdRules.AllowUserToDeleteRows = $false
    $grdRules.AllowUserToResizeRows = $false
    $grdRules.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdRules.MultiSelect = $false
    $grdRules.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grdRules.Font = New-Object System.Drawing.Font("Consolas", 9)
    [void]$grdRules.Columns.Add('Name', 'Name')
    [void]$grdRules.Columns.Add('Description', 'Description')
    [void]$grdRules.Columns.Add('Enabled', 'Enabled')
    [void]$grdRules.Columns.Add('Identity', 'Identity')
    $grdRules.Columns['Name'].FillWeight        = 25
    $grdRules.Columns['Description'].FillWeight = 45
    $grdRules.Columns['Enabled'].FillWeight     = 10
    $grdRules.Columns['Identity'].Visible       = $false
    foreach ($col in $grdRules.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable }
    $rulesForm.Controls.Add($grdRules)

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Multiline = $true
    $txtOut.ReadOnly = $true
    $txtOut.ScrollBars = "Vertical"
    $txtOut.Location = New-Object System.Drawing.Point(12, 380)
    $txtOut.Size = New-Object System.Drawing.Size(740, 44)
    $txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
    $rulesForm.Controls.Add($txtOut)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh"
    $btnRefresh.Location = New-Object System.Drawing.Point(12, 428)
    $btnRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $rulesForm.Controls.Add($btnRefresh)
    $btnRulesClose = New-Object System.Windows.Forms.Button
    $btnRulesClose.Text = "Close"
    $btnRulesClose.Location = New-Object System.Drawing.Point(652, 428)
    $btnRulesClose.Size = New-Object System.Drawing.Size(100, 28)
    $rulesForm.Controls.Add($btnRulesClose)

    $script:CurrentRules = @()

    function Update-RulesFromCurrent {
        $rulesForm.Text = if ("$script:Target".Trim()) { "Inbox Rules: $script:Target" } else { 'Inbox Rules' }
        Update-RulesGrid -Smtp $script:Target
    }

    Update-RulesFromCurrent
    $btnRefresh.Add_Click({ Update-RulesFromCurrent })
    $grdRules.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $identity = $grdRules.Rows[$e.RowIndex].Cells['Identity'].Value
        $rule = $script:CurrentRules | Where-Object { $_.Identity.ToString() -eq $identity } | Select-Object -First 1
        if (-not $rule) {
            $txtOut.Text = "Could not match that row back to a rule -- click Refresh and try again."
            return
        }
        Show-RuleDetailForm -Rule $rule -Smtp $script:Target
        Update-RulesFromCurrent
    })
    $btnRulesClose.Add_Click({ $rulesForm.Close() })

    # BEFORE Add-ToolHelpButton: the bar shifts every existing control down and
    # grows the form, and the help button places itself from ClientSize at run
    # time. Before Set-FormScreenCorner too, which positions from the size.
    $bar = Add-ToolTargetBar -Form $rulesForm -InitialValue $script:Target

    # WIRED HERE, NOT PASSED INTO THE HELPER. Update-RulesFromCurrent is nested
    # in this function, so the handler has to be built in this scope to see it.
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        Update-RulesFromCurrent
    })

    Set-ConsoleTheme -Form $rulesForm
    Set-FormScreenCorner -Form $rulesForm -Corner 'BottomLeft'
    # Consistent '?' in the top-right of every tool. Added after the theme so
    # it is not restyled, and BringToFront'd inside the helper because
    # WinForms draws EARLIER-added controls on top.
    Add-ToolHelpButton -Form $rulesForm -ScriptPath $PSCommandPath -Title 'Inbox Rules' | Out-Null

    [void]$rulesForm.ShowDialog()
}
#endregion

Show-InboxRulesForm
```


<br>

==============================================================================

# ◆ FILE 6 OF 8 — `Group-Manager.ps1`

> **732 lines.** TOOL 8, THE FRONT DOOR. Pairs with Group-Manager-Editor.ps1 in file 1, which is the editor it routes to. This one classifies: WHAT the group is and WHERE it lives, then opens the right editor. A Microsoft 365 Group is a different animal managed by a separate cmdlet family (*-UnifiedGroup) that the distribution-list editor cannot speak.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```powershell
#Requires -Version 5.1
<#
    Group-Manager.ps1
    One front door for ANY group. Looks it up, works out what kind it is and
    where it is mastered, and opens the correct editor for both answers.

    RENAMED FROM Group-Manager.ps1 ON 2026-08-21, and the old name was accurate
    rather than merely dated: the tool routed on ONE axis -- which side of the
    hybrid owns the group -- while silently assuming every group was a
    distribution list. Point it at a Microsoft 365 Group and it failed with
    Exchange's own words, "The current operation is not supported on
    GroupMailbox", because Get-DistributionGroup was the first question it
    asked. The type is a second axis, and now it is asked first.

    Reasoning: ScriptNotes/Group-Manager.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-21
#>
# TOOLKIT-VARIANT: Router
# TOOLKIT-HELP-BEGIN
# Start here for any group -- distribution list, mail-enabled security group, or
# Microsoft 365 Group.
#
# Look the group up and this window tells you two things: WHAT it is, and WHERE
# it lives. Then it opens the right editor for that combination, so you do not
# have to know which of three tools to pick.
#
#   Distribution list, synced     Group Manager Editor, against on-premises
#   Distribution list, cloud      Group Manager Editor, against Exchange Online
#   Microsoft 365 Group           Group Editor -- members and owners
#
# A MICROSOFT 365 GROUP IS A DIFFERENT ANIMAL, and that is why it gets its own
# destination. Access to one comes from MEMBERSHIP, not from mailbox permissions,
# and it is managed with a completely separate family of cmdlets
# (*-UnifiedGroup) that the distribution-list editor cannot speak.
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
        Cloud      = 'Group-Manager-Editor.ps1'
        OnPrem     = 'Group-Manager-Editor.ps1'
        OnPremArgs = @('-OnPrem')

        # A THIRD DESTINATION, FOR A THIRD KIND OF GROUP.
        #
        # Group-Manager-Editor is built entirely on *-DistributionGroup, and a
        # Microsoft 365 Group refuses those with "The current operation is not
        # supported on GroupMailbox". GroupEditor is the tool that speaks
        # *-UnifiedGroup and *-UnifiedGroupLinks.
        #
        # THIS IS WHY THE TOOL WAS RENAMED. It routed on ONE axis -- which side
        # of the hybrid owns the group -- and silently assumed every group was a
        # distribution list. The type is a second axis, and 'Group-Manager' was an
        # honest name for a tool that only understood one of them.
        Unified    = 'GroupEditor.ps1'
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
        Kind        = ''     # 'Distribution' or 'Unified'
        Error       = ''
    }

    # WHAT KIND OF GROUP IS THIS, ASKED BEFORE WHERE IT LIVES.
    #
    # Get-DistributionGroup used to be the first and only question, and it is
    # the wrong one to open with: a Microsoft 365 Group answers it with "The
    # current operation is not supported on GroupMailbox" -- an error that reads
    # like a broken tool rather than a category mistake. Three real groups were
    # reported as unroutable that way.
    #
    # Get-Recipient answers for EVERY group type and names which one it found,
    # so the type decides the path instead of the first cmdlet's tolerance.
    # AND THE SAME LOOKUP RESOLVES A BARE NAME.
    #
    # Get-TargetSmtp turns 'qateam' into 'qateam@<primary domain>'. A Microsoft
    # 365 Group whose only address is @tenant.onmicrosoft.com does not exist at
    # that guess, so typing the full address worked and typing the name did not.
    #
    # Exchange's -Identity accepts an alias, so the bare word IS a valid
    # identity. Try the completed address first (it is right nearly always),
    # then the bare name, and adopt whichever answers.
    $kind    = 'Distribution'
    $lookup  = "$Smtp"
    $bare    = if ("$Smtp" -match '^([^@]+)@') { $Matches[1] } else { '' }

    foreach ($id in @($Smtp, $bare)) {
        if (-not "$id".Trim()) { continue }
        try {
            $r = @(Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{
                Identity = $id; ResultSize = 1; ErrorAction = 'Stop' })
            if ($r.Count -gt 0) {
                if ("$($r[0].RecipientTypeDetails)" -match '(?i)groupmailbox') { $kind = 'Unified' }
                # THE IDENTITY THAT ANSWERED is what the rest of this function
                # asks with. Carrying the guess forward would find the group and
                # then fail the very next call.
                $lookup = $id
                break
            }
        } catch {
            # Try the next identity. If neither answers, the path below reports
            # the error the operator has always seen for a missing group.
        }
    }
    $state.Kind = $kind
    $Smtp       = $lookup

    if ($kind -eq 'Unified') {
        # A MICROSOFT 365 GROUP IS ALWAYS CLOUD-MANAGED FOR THESE PURPOSES.
        #
        # There is no IsDirSynced question to ask: Get-UnifiedGroup is a
        # cloud-only cmdlet, membership and ownership are held in the cloud, and
        # GroupEditor -- the tool this routes to -- is TOOLKIT-VARIANT: Cloud.
        # Routing one on-premises would open a tool that cannot see it.
        try {
            $g = Invoke-EXOCommand -Name 'Get-UnifiedGroup' -Parameters @{
                Identity = $Smtp; ErrorAction = 'Stop' }
        } catch {
            $state.Error = ("Exchange Online has no Microsoft 365 Group at that address.`r`n`r`n" +
                            "Underlying error: $($_.Exception.Message)")
            return $state
        }
        $state.Found       = $true
        $state.DisplayName = "$($g.DisplayName)"
        $state.PrimarySmtp = "$($g.PrimarySmtpAddress)"
        $state.GroupType   = 'Microsoft 365 Group'
        $state.IsDirSynced = $false
        $state.Where       = 'Cloud'
        return $state
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
        # THREE-WAY, NOT TWO-WAY. The Kind of group is asked BEFORE the side of
        # the hybrid, because a Microsoft 365 Group has no on-premises side to
        # choose and its editor is a different tool entirely. Getting this order
        # wrong is what produced "not supported on GroupMailbox".
        $toolFile =
            if     ($script:State.Kind  -eq 'Unified') { "$($act.Unified)" }
            elseif ($script:State.Where -eq 'Cloud')   { "$($act.Cloud)" }
            else                                       { "$($act.OnPrem)" }
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
        # A MICROSOFT 365 GROUP IS NEVER LAUNCHED DETACHED. Get-UnifiedGroup is
        # cloud-only, so GroupEditor needs the Exchange Online session that lives
        # in THIS process -- a child would open disconnected.
        if ($script:State.Kind -ne 'Unified' -and
            $script:State.Where -ne 'Cloud' -and (Test-ToolNeedsOwnProcess -ScriptName $toolFile)) {
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
        if ($script:State.Kind -eq 'Unified') {
            # A THIRD VERDICT, NAMED RATHER THAN FOLDED INTO 'CLOUD-MANAGED'.
            #
            # Both are cloud-managed and both open in this process, so the
            # routing is identical -- but the TOOL is not, and the operator
            # should know they are about to land somewhere different. Saying
            # "Microsoft 365 Group" also answers the question this tool used to
            # fail on without explanation.
            $lblBanner.Text      = 'MICROSOFT 365 GROUP  --  members and owners, in Group Editor'
            $lblBanner.ForeColor = [System.Drawing.Color]::FromArgb(80, 220, 130)
        } elseif ($script:State.Where -eq 'Cloud') {
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
            $toolFile =
                if     ($script:State.Kind  -eq 'Unified') { "$($act.Unified)" }
                elseif ($script:State.Where -eq 'Cloud')   { "$($act.Cloud)" }
                else                                       { "$($act.OnPrem)" }
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
```


<br>

==============================================================================

# ◆ FILE 7 OF 8 — `MailAttributes.ps1`

> **739 lines.** THE PROXY-ADDRESS REFERENCE, and the most dangerous operation in the toolkit. Constraint C11 lives here: a bare EmailAddresses assignment REPLACES the whole collection, destroying X500 addresses and silently breaking replies to historical mail. Read how the existing collection is carried across before writing anything that touches addresses.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```powershell
#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# ---------------------------------------------------------------------------
# CLOUD VERSION, AND NOW THE ONLY VERSION.
#
# AD-MailAttributes.ps1 was the on-premises twin. It has been PARKED to
# DisabledTools\, so there is no longer an AD- variant beside this file and
# nothing here needs a domain controller.
#
# The header used to say the AD- variant was "beside it", which stopped being
# true the moment the file moved. A comment that names a sibling is a
# reference like any other -- it goes stale silently, and the next person
# reads it as current.
# -------------------------------------------------------------------------
<#
    MailAttributes.ps1
    Standalone Mail Attributes tool -- Layer 2 tool script.

    Cloud-only build. Reads and writes the mailbox object in Exchange Online
    via Get-User / Get-Mailbox and Set-User / Set-Mailbox. There is no AD
    account behind it and nothing here touches one.

    Author   : Glenn Armbruster
    Revised  : 2026-08-03
#>
# TOOLKIT-HELP-BEGIN
# View and edit the mail-related attributes on an account -- addresses, display
# name, alias, and how the person appears in the address book.
#
# Mostly used for renames: someone marries, changes name, or was created with a
# typo. Adding the new address while keeping the old one is the safe move,
# because mail sent to the old address keeps arriving.
#
# Changing the PRIMARY address changes what recipients see and what shows in the
# directory. The old address should normally be kept as a secondary rather than
# deleted.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress
)

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

# NO PROMPT BEFORE THE WINDOW. The address box is ON the form -- see
# Add-ToolTargetBar in Common.ps1. An empty address is fine; the window
# opens and asks for one.
$script:Target = Get-TargetSmtp -Text $SmtpAddress

#region Actions - mail attributes ---------------------------------------------
<#
    Cloud attribute set.

    The on-premises tool read raw AD attributes. Exchange Online exposes the
    same facts under different names, and a few not at all, so the list is the
    intersection rather than a rename of the old one:

        AD                    Cloud                      Read from
        --------------------  -------------------------  --------------
        displayName           DisplayName                Get-User
        givenName             FirstName                  Get-User
        initials              Initials                   Get-User
        sn                    LastName                   Get-User
        company               Company                    Get-User
        name                  Name                       Get-User
        userPrincipalName     UserPrincipalName          Get-User
        mail                  WindowsEmailAddress        Get-User
        mailNickname          Alias                      Get-Mailbox
        msExchArchiveName     ArchiveName                Get-Mailbox
        proxyAddresses        EmailAddresses             Get-Mailbox

    Dropped, deliberately:

    - targetAddress. It is the on-premises pointer to the cloud mailbox. In a
      cloud-only tenant the mailbox IS the destination, so there is nothing for
      it to point at. Get-MailUser -ExternalEmailAddress is the equivalent for
      mail users, which is a different object and a different tool.
    - Rename of the directory object. Set-User -Name renames the mailbox, not
      an AD RDN -- a different operation with different consequences, so it is
      shown read-only here rather than quietly repurposing the old edit path.
    - mail / WindowsEmailAddress editing. Set-Mailbox -WindowsEmailAddress has
      no effect while the mailbox is subject to an email address policy, which
      is the normal state. An edit that silently does nothing is worse than no
      edit button, so it is read-only.
#>
$script:MailAttribsDisplayOrder = @(
    'Company', 'DisplayName', 'Name', 'FirstName', 'Initials', 'LastName',
    'WindowsEmailAddress', 'Alias', 'ArchiveName', 'PrimarySmtpAddress',
    'EmailAddresses', 'RecipientTypeDetails', 'UserPrincipalName'
)
$script:MailAttribsDisplayLabels = @{
    'Company'              = 'Office'
    'FirstName'            = 'FirstName'
    'LastName'             = 'LastName'
    'WindowsEmailAddress'  = 'mail (WindowsEmailAddress)'
    'Alias'                = 'Alias (mailNickname)'
    'ArchiveName'          = 'ArchiveName (msExchArchiveName)'
    'EmailAddresses'       = 'EmailAddresses (proxyAddresses)'
    'PrimarySmtpAddress'   = 'Primary SMTP'
    'RecipientTypeDetails' = 'Mailbox Type'
}
$script:MailAttribsEditableAttributes = @(
    'DisplayName', 'FirstName', 'Initials', 'LastName', 'Alias',
    'ArchiveName', 'PrimarySmtpAddress'
)
function Get-MailAttribsUser {
    <#
        One record assembled from two cmdlets.

        Get-User carries the person fields, Get-Mailbox the mail fields, and
        neither has all of them. Both are keyed on the same address, so a
        mailbox that resolves in one and not the other is a real fault worth
        surfacing rather than hiding behind a half-populated grid.
    #>
    param([Parameter(Mandatory)] [string]$Smtp)

    $u = Invoke-EXOCommand -Name 'Get-User' -Parameters @{
        Identity = $Smtp; ErrorAction = 'Stop' }
    $mb = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
        Identity = $Smtp; ErrorAction = 'Stop' }

    return [PSCustomObject]@{
        Identity             = "$($mb.Identity)"
        Company              = "$($u.Company)"
        DisplayName          = "$($u.DisplayName)"
        Name                 = "$($u.Name)"
        FirstName            = "$($u.FirstName)"
        Initials             = "$($u.Initials)"
        LastName             = "$($u.LastName)"
        WindowsEmailAddress  = "$($u.WindowsEmailAddress)"
        UserPrincipalName    = "$($u.UserPrincipalName)"
        Alias                = "$($mb.Alias)"
        ArchiveName          = (@($mb.ArchiveName) -join '; ')
        ArchiveEnabled       = ("$($mb.ArchiveStatus)" -ne 'None' -and "$($mb.ArchiveGuid)" -ne '00000000-0000-0000-0000-000000000000')
        RecipientTypeDetails = "$($mb.RecipientTypeDetails)"
        EmailAddresses       = @($mb.EmailAddresses | ForEach-Object { "$_" })
    }
}
function Get-PrimarySmtpFromProxyAddresses {
    param($ProxyAddresses)
    foreach ($addr in @($ProxyAddresses)) {
        if ("$addr" -cmatch '^SMTP:(.+)$') { return $Matches[1] }
    }
    return ''
}
function Set-PrimarySmtpInProxyAddresses {
    param(
        $CurrentProxyAddresses,
        [Parameter(Mandatory)] [string]$NewPrimary,
        [bool]$KeepOldPrimaryAsAlias = $true
    )
    $kept = New-Object System.Collections.Generic.List[string]
    $oldPrimary = $null
    foreach ($addr in @($CurrentProxyAddresses)) {
        $s = "$addr"
        if ($s -cmatch '^SMTP:(.+)$') {
            $oldPrimary = $Matches[1]
            continue
        }
        if ($s -match '^smtp:(.+)$' -and $Matches[1] -ieq $NewPrimary) {
            continue
        }
        [void]$kept.Add($s)
    }
    if ($KeepOldPrimaryAsAlias -and $oldPrimary -and $oldPrimary -ine $NewPrimary) {
        [void]$kept.Add("smtp:$oldPrimary")
    }
    $result = New-Object System.Collections.Generic.List[string]
    [void]$result.Add("SMTP:$NewPrimary")
    foreach ($k in $kept) { [void]$result.Add($k) }
    return $result.ToArray()
}
function Add-SecondarySmtpToProxyAddresses {
    param(
        $CurrentProxyAddresses,
        [Parameter(Mandatory)] [string]$NewSecondary
    )
    foreach ($addr in @($CurrentProxyAddresses)) {
        if ("$addr" -match '^smtp:(.+)$' -and $Matches[1] -ieq $NewSecondary) {
            throw "'$NewSecondary' already exists in EmailAddresses (as $addr)."
        }
    }
    $result = New-Object System.Collections.Generic.List[string]
    foreach ($addr in @($CurrentProxyAddresses)) { [void]$result.Add("$addr") }
    [void]$result.Add("smtp:$NewSecondary")
    return $result.ToArray()
}
function Format-AttributeValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) {
        $items = New-Object System.Collections.Generic.List[string]
        foreach ($item in $Value) { [void]$items.Add("$item") }
        return ($items -join '; ')
    }
    return "$Value"
}
function Format-MailAttribsTable {
    param($UserObject)
    $table = New-Object System.Data.DataTable
    [void]$table.Columns.Add('Attribute', [string])
    [void]$table.Columns.Add('Value', [string])
    [void]$table.Columns.Add('_Editable', [string])
    [void]$table.Columns.Add('_AttributeKey', [string])
    foreach ($attrName in $script:MailAttribsDisplayOrder) {
        $rawValue = if ($attrName -eq 'PrimarySmtpAddress') {
            Get-PrimarySmtpFromProxyAddresses -ProxyAddresses $UserObject.EmailAddresses
        } else {
            $UserObject.$attrName
        }
        $label = if ($script:MailAttribsDisplayLabels.ContainsKey($attrName)) { $script:MailAttribsDisplayLabels[$attrName] } else { $attrName }
        $row = $table.NewRow()
        $row['Attribute']     = $label
        $row['Value']         = Format-AttributeValue -Value $rawValue
        $row['_Editable']     = if ($script:MailAttribsEditableAttributes -contains $attrName) { 'Y' } else { 'N' }
        $row['_AttributeKey'] = $attrName
        [void]$table.Rows.Add($row)
    }
    return ,$table
}
function Set-MailAttribsAttribute {
    <#
        Set-User for the person fields, Set-Mailbox for the mail fields.

        Clearing works differently from on-premises: there is no -Clear, so a
        blank is written as $null. Alias and PrimarySmtpAddress refuse a blank
        outright -- an aliasless or addressless mailbox is broken, not empty.

        PrimarySmtpAddress is NOT a Set-Mailbox parameter in Exchange Online
        (Microsoft documents it as on-premises only). The primary is changed by
        rewriting the whole EmailAddresses collection with exactly one
        uppercase SMTP: entry, which is what Set-PrimarySmtpInProxyAddresses
        already builds.
    #>
    param(
        [Parameter(Mandatory)] [string]$Identity,
        [Parameter(Mandatory)] [string]$AttributeName,
        [string]$NewValue,
        $CurrentProxyAddresses,
        [bool]$KeepOldPrimaryAsAlias = $true
    )
    $isBlank = ("$NewValue".Trim() -eq '')
    $value   = if ($isBlank) { $null } else { $NewValue.Trim() }

    switch ($AttributeName) {
        'DisplayName' {
            Invoke-EXOCommand -Name 'Set-User' -Parameters @{
                Identity = $Identity; DisplayName = $value; ErrorAction = 'Stop' } | Out-Null
        }
        'FirstName' {
            Invoke-EXOCommand -Name 'Set-User' -Parameters @{
                Identity = $Identity; FirstName = $value; ErrorAction = 'Stop' } | Out-Null
        }
        'Initials' {
            Invoke-EXOCommand -Name 'Set-User' -Parameters @{
                Identity = $Identity; Initials = $value; ErrorAction = 'Stop' } | Out-Null
        }
        'LastName' {
            Invoke-EXOCommand -Name 'Set-User' -Parameters @{
                Identity = $Identity; LastName = $value; ErrorAction = 'Stop' } | Out-Null
        }
        'Alias' {
            if ($isBlank) { throw "Alias cannot be blank -- a mailbox without one is not addressable." }
            Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
                Identity = $Identity; Alias = $value; ErrorAction = 'Stop' } | Out-Null
        }
        'ArchiveName' {
            Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
                Identity = $Identity; ArchiveName = $value; ErrorAction = 'Stop' } | Out-Null
        }
        'PrimarySmtpAddress' {
            if ($isBlank) { throw "PrimarySmtpAddress cannot be blank." }
            $newProxyAddresses = Set-PrimarySmtpInProxyAddresses -CurrentProxyAddresses $CurrentProxyAddresses -NewPrimary $value -KeepOldPrimaryAsAlias $KeepOldPrimaryAsAlias
            Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
                Identity = $Identity; EmailAddresses = $newProxyAddresses; ErrorAction = 'Stop' } | Out-Null
        }
        default { throw "Attribute '$AttributeName' is not handled by Set-MailAttribsAttribute." }
    }
}
function Add-MailAttribsSecondaryProxyAddress {
    param(
        [Parameter(Mandatory)] [string]$Identity,
        $CurrentProxyAddresses,
        [Parameter(Mandatory)] [string]$NewSecondary
    )
    $newProxyAddresses = Add-SecondarySmtpToProxyAddresses -CurrentProxyAddresses $CurrentProxyAddresses -NewSecondary $NewSecondary
    Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
        Identity = $Identity; EmailAddresses = $newProxyAddresses; ErrorAction = 'Stop' } | Out-Null
}
function Show-MailAttribsEditDialog {
    param(
        [System.Windows.Forms.Form]$Owner,
        [string]$AttributeName,
        [string]$CurrentValue
    )
    $warningNote = switch ($AttributeName) {
        'Alias' { 'The Alias is the mailbox nickname. Changing it can regenerate the primary address if this mailbox is still subject to an email address policy -- check the Primary SMTP row afterwards. Cannot be left blank.' }
        'ArchiveName' { 'This only names an archive that already exists; it does not create one. If this mailbox has no archive the change is accepted and has no visible effect. Outlook desktop ignores this name and shows "Online Archive - <primary address>" regardless -- it is Outlook on the web that displays it.' }
        'PrimarySmtpAddress' { 'This rewrites the whole EmailAddresses collection: sets this value as the SMTP: (primary) entry. If a different primary already exists, you will be prompted next to Keep it as a secondary alias or Remove it entirely. Set-Mailbox -PrimarySmtpAddress does not exist in Exchange Online, so this is the supported route. Cannot be left blank.' }
        default { '' }
    }
    $dlgHeight = if ($warningNote -ne '') { 330 } else { 300 }
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Edit Attribute -- $AttributeName"
    $dlg.Size = New-Object System.Drawing.Size(520, $dlgHeight)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblAttr = New-Object System.Windows.Forms.Label
    $lblAttr.Text = "Attribute: $AttributeName"
    $lblAttr.Location = New-Object System.Drawing.Point(12, 12)
    $lblAttr.AutoSize = $true
    $lblAttr.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $dlg.Controls.Add($lblAttr)
    $nextY = 42
    if ($warningNote -ne '') {
        $lblWarnNote = New-Object System.Windows.Forms.Label
        $lblWarnNote.Text = $warningNote
        $lblWarnNote.Location = New-Object System.Drawing.Point(12, $nextY)
        $lblWarnNote.Size = New-Object System.Drawing.Size(480, 40)
        $lblWarnNote.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Italic)
        $lblWarnNote.ForeColor = [System.Drawing.Color]::DarkOrange
        $dlg.Controls.Add($lblWarnNote)
        $nextY += 46
    }
    $lblCurrent = New-Object System.Windows.Forms.Label
    $lblCurrent.Text = "Current Value:"
    $lblCurrent.Location = New-Object System.Drawing.Point(12, $nextY)
    $lblCurrent.AutoSize = $true
    $dlg.Controls.Add($lblCurrent)
    $nextY += 18
    $txtCurrent = New-Object System.Windows.Forms.TextBox
    $txtCurrent.Location = New-Object System.Drawing.Point(12, $nextY)
    $txtCurrent.Size = New-Object System.Drawing.Size(480, 40)
    $txtCurrent.Multiline = $true
    $txtCurrent.ReadOnly = $true
    $txtCurrent.ScrollBars = "Vertical"
    $txtCurrent.BackColor = [System.Drawing.Color]::WhiteSmoke
    $txtCurrent.Text = $CurrentValue
    $txtCurrent.Font = New-Object System.Drawing.Font("Consolas", 9)
    $dlg.Controls.Add($txtCurrent)
    $nextY += 48
    $lblNew = New-Object System.Windows.Forms.Label
    $lblNew.Text = "New Value (blank only allowed if this field is already blank):"
    $lblNew.Location = New-Object System.Drawing.Point(12, $nextY)
    $lblNew.AutoSize = $true
    $dlg.Controls.Add($lblNew)
    $nextY += 18
    $txtNew = New-Object System.Windows.Forms.TextBox
    $txtNew.Location = New-Object System.Drawing.Point(12, $nextY)
    $txtNew.Size = New-Object System.Drawing.Size(480, 60)
    $txtNew.Multiline = $true
    $txtNew.ScrollBars = "Vertical"
    $txtNew.Text = $CurrentValue
    $txtNew.Font = New-Object System.Drawing.Font("Consolas", 9)
    $dlg.Controls.Add($txtNew)
    $nextY += 68
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(320, $nextY)
    $btnCancel.Size = New-Object System.Drawing.Size(84, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(408, $nextY)
    $btnOk.Size = New-Object System.Drawing.Size(84, 28)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton = $btnOk
    $result = $dlg.ShowDialog($Owner)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return $txtNew.Text
}
function Show-KeepOrRemoveOldPrimaryDialog {
    param(
        [System.Windows.Forms.Form]$Owner,
        [string]$OldPrimary,
        [string]$NewPrimary
    )
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Previous Primary SMTP Address"
    $dlg.ClientSize = New-Object System.Drawing.Size(444, 174)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $dlg.Tag = 'Cancel'
    $lblMsg = New-Object System.Windows.Forms.Label
    $lblMsg.Text = "Current primary SMTP address:`n$OldPrimary`n`nIt is being replaced by:`n$NewPrimary`n`nWhat should I do with the current PrimarySmtpAddress?"
    $lblMsg.Location = New-Object System.Drawing.Point(12, 12)
    $lblMsg.Size = New-Object System.Drawing.Size(430, 100)
    $dlg.Controls.Add($lblMsg)
    $btnKeep = New-Object System.Windows.Forms.Button
    $btnKeep.Text = "Keep as Alias"
    $btnKeep.Location = New-Object System.Drawing.Point(20, 130)
    $btnKeep.Size = New-Object System.Drawing.Size(130, 32)
    $btnKeep.Add_Click({ $dlg.Tag = 'Keep'; $dlg.Close() })
    $dlg.Controls.Add($btnKeep)
    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text = "Remove"
    $btnRemove.Location = New-Object System.Drawing.Point(165, 130)
    $btnRemove.Size = New-Object System.Drawing.Size(110, 32)
    $btnRemove.Add_Click({ $dlg.Tag = 'Remove'; $dlg.Close() })
    $dlg.Controls.Add($btnRemove)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(290, 130)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 32)
    $btnCancel.Add_Click({ $dlg.Tag = 'Cancel'; $dlg.Close() })
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel
    [void]$dlg.ShowDialog($Owner)
    return "$($dlg.Tag)"
}
function Show-MailAttribsForm {
    # NO -Smtp PARAMETER. The address can change while the window is open, so
    # it lives in $script:Target rather than a parameter fixed at build time.
    $mailForm = New-Object System.Windows.Forms.Form
    $mailForm.Text = "Mail Attributes: $script:Target"
    $mailForm.ClientSize = New-Object System.Drawing.Size(664, 656)
    $mailForm.StartPosition = 'Manual'
    $mailForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Text = "Reads the mailbox from Exchange Online (Get-User + Get-Mailbox) -- cloud only, no directory account is touched. Double-click a row, or select it and click 'Edit Attribute' below, to edit it. 'EmailAddresses' is read-only here -- edit 'Primary SMTP' to change the primary address (you'll be prompted to Keep the old primary as an alias or Remove it), or use 'Add Secondary' below to add another address. Name, UserPrincipalName, mail, Office and Mailbox Type are read-only/context-only. targetAddress has no cloud equivalent and is not shown."
    $lblNote.Location = New-Object System.Drawing.Point(12, 12)
    $lblNote.Size = New-Object System.Drawing.Size(640, 94)
    $mailForm.Controls.Add($lblNote)
    $grdMail = New-Object System.Windows.Forms.DataGridView
    $grdMail.Location = New-Object System.Drawing.Point(12, 110)
    $grdMail.Size = New-Object System.Drawing.Size(640, 304)
    $grdMail.ReadOnly = $true
    $grdMail.AllowUserToAddRows = $false
    $grdMail.AllowUserToDeleteRows = $false
    $grdMail.AllowUserToResizeRows = $false
    $grdMail.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdMail.MultiSelect = $false
    $grdMail.AutoGenerateColumns = $false
    $grdMail.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
    $grdMail.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
    $grdMail.RowHeadersVisible = $false
    $grdMail.BackgroundColor = [System.Drawing.Color]::White
    $grdMail.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
    $grdMail.Font = New-Object System.Drawing.Font("Consolas", 9)
    $grdMail.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue
    $colMailAttribute = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailAttribute.Name = 'Attribute'
    $colMailAttribute.HeaderText = 'Attribute'
    $colMailAttribute.DataPropertyName = 'Attribute'
    $colMailAttribute.Width = 160
    $colMailAttribute.ReadOnly = $true
    [void]$grdMail.Columns.Add($colMailAttribute)
    $colMailValue = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailValue.Name = 'Value'
    $colMailValue.HeaderText = 'Value'
    $colMailValue.DataPropertyName = 'Value'
    $colMailValue.ReadOnly = $true
    $colMailValue.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    [void]$grdMail.Columns.Add($colMailValue)
    $colMailEditableFlag = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailEditableFlag.Name = '_Editable'
    $colMailEditableFlag.DataPropertyName = '_Editable'
    $colMailEditableFlag.Visible = $false
    [void]$grdMail.Columns.Add($colMailEditableFlag)
    $colMailAttributeKey = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMailAttributeKey.Name = '_AttributeKey'
    $colMailAttributeKey.DataPropertyName = '_AttributeKey'
    $colMailAttributeKey.Visible = $false
    [void]$grdMail.Columns.Add($colMailAttributeKey)
    $mailForm.Controls.Add($grdMail)
    $grdMail.DataSource = New-Object System.Windows.Forms.BindingSource
    $mailBindingSource = $grdMail.DataSource
    $lblAddSmtp = New-Object System.Windows.Forms.Label
    $lblAddSmtp.Text = "Add secondary SMTP address (appended as smtp: -- does not replace the primary or any other address):"
    $lblAddSmtp.Location = New-Object System.Drawing.Point(12, 422)
    $lblAddSmtp.Size = New-Object System.Drawing.Size(640, 16)
    $mailForm.Controls.Add($lblAddSmtp)
    $txtAddSmtp = New-Object System.Windows.Forms.TextBox
    $txtAddSmtp.Location = New-Object System.Drawing.Point(12, 440)
    $txtAddSmtp.Size = New-Object System.Drawing.Size(430, 22)
    $txtAddSmtp.Font = New-Object System.Drawing.Font("Consolas", 9)
    $mailForm.Controls.Add($txtAddSmtp)
    $btnAddSmtp = New-Object System.Windows.Forms.Button
    $btnAddSmtp.Text = "Add Secondary"
    $btnAddSmtp.Location = New-Object System.Drawing.Point(452, 439)
    $btnAddSmtp.Size = New-Object System.Drawing.Size(110, 26)
    $mailForm.Controls.Add($btnAddSmtp)
    $lblMailLog = New-Object System.Windows.Forms.Label
    $lblMailLog.Text = "Log:"
    $lblMailLog.Location = New-Object System.Drawing.Point(12, 474)
    $lblMailLog.AutoSize = $true
    $mailForm.Controls.Add($lblMailLog)
    $txtMailLog = New-Object System.Windows.Forms.TextBox
    $txtMailLog.Multiline = $true
    $txtMailLog.ReadOnly = $true
    $txtMailLog.ScrollBars = "Vertical"
    $txtMailLog.Location = New-Object System.Drawing.Point(12, 494)
    $txtMailLog.Size = New-Object System.Drawing.Size(640, 110)
    $txtMailLog.Font = New-Object System.Drawing.Font("Consolas", 9)
    $txtMailLog.BackColor = [System.Drawing.Color]::WhiteSmoke
    $mailForm.Controls.Add($txtMailLog)
    $btnMailRefresh = New-Object System.Windows.Forms.Button
    $btnMailRefresh.Text = "Refresh"
    $btnMailRefresh.Location = New-Object System.Drawing.Point(12, 616)
    $btnMailRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $mailForm.Controls.Add($btnMailRefresh)
    $btnMailEdit = New-Object System.Windows.Forms.Button
    $btnMailEdit.Text = "Edit Attribute"
    $btnMailEdit.Location = New-Object System.Drawing.Point(122, 616)
    $btnMailEdit.Size = New-Object System.Drawing.Size(140, 28)
    $btnMailEdit.Enabled = $false
    $btnMailEdit.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
    $mailForm.Controls.Add($btnMailEdit)
    $btnMailClose = New-Object System.Windows.Forms.Button
    $btnMailClose.Text = "Close"
    $btnMailClose.Location = New-Object System.Drawing.Point(552, 616)
    $btnMailClose.Size = New-Object System.Drawing.Size(100, 28)
    $mailForm.Controls.Add($btnMailClose)
    $script:MailAttribsCurrentUser = $null
    function Update-MailAttribsGrid {
        param($UserObject)
        $table = Format-MailAttribsTable -UserObject $UserObject
        $mailBindingSource.DataSource = $null
        $mailBindingSource.DataSource = $table
        $grdMail.Refresh()
    }
    function Update-MailAttribsFromCurrent {
        # OPENED WITH NO ADDRESS IS A NORMAL STATE NOW -- the window appears
        # before anything has been typed, so the refresh must not run empty.
        if (-not "$script:Target".Trim()) {
            $mailForm.Text = 'Mail Attributes'
            return
        }
        $mailForm.Text = 'Mail Attributes: ' + $script:Target
        $btnMailRefresh.Enabled = $false
        $btnAddSmtp.Enabled = $false
        $txtMailLog.Text = "Loading mail attributes..."
        $mailForm.Refresh()
        try {
            $script:MailAttribsCurrentUser = Get-MailAttribsUser -Smtp $script:Target
            Update-MailAttribsGrid -UserObject $script:MailAttribsCurrentUser
            $txtMailLog.Text = "Loaded mail attributes for '$($script:MailAttribsCurrentUser.UserPrincipalName)' ($($script:MailAttribsCurrentUser.RecipientTypeDetails))."
        } catch {
            $script:MailAttribsCurrentUser = $null
            $mailBindingSource.DataSource = $null
            $txtMailLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mail Attribs'
        } finally {
            $btnMailRefresh.Enabled = $true
            $btnAddSmtp.Enabled = $true
        }
    }
    Update-MailAttribsFromCurrent
    $btnMailRefresh.Add_Click({ Update-MailAttribsFromCurrent })
    function Invoke-MailAttribsEditSelectedRow {
        if ($grdMail.SelectedRows.Count -eq 0 -or $null -eq $script:MailAttribsCurrentUser) { return }
        $row = $grdMail.SelectedRows[0]
        $attrLabel = "$($row.Cells['Attribute'].Value)"
        $attrName  = "$($row.Cells['_AttributeKey'].Value)"
        if ($script:MailAttribsEditableAttributes -notcontains $attrName) { return }
        $btnMailRefresh.Enabled = $false
        $btnMailEdit.Enabled = $false
        $btnAddSmtp.Enabled = $false
        try {
            $oldValue = "$($row.Cells['Value'].Value)"
            $newValue = Show-MailAttribsEditDialog -Owner $mailForm -AttributeName $attrLabel -CurrentValue $oldValue
            if ($null -eq $newValue) {
                $txtMailLog.Text = "Edit cancelled."
                return
            }
            if ($script:MailAttribsNeverBlank -contains $attrName -and $newValue.Trim() -eq '') {
                [void][System.Windows.Forms.MessageBox]::Show("'$attrLabel' cannot be blank.", 'Invalid Input', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            $oldIsBlank = ("$oldValue".Trim() -eq '')
            $newIsBlank = $newValue.Trim() -eq ''
            if ($newIsBlank -and -not $oldIsBlank) {
                [void][System.Windows.Forms.MessageBox]::Show("'$attrLabel' currently has a value and cannot be saved blank. Enter a replacement value instead.", 'Invalid Input', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            $keepOldPrimaryAsAlias = $true
            if ($attrName -eq 'PrimarySmtpAddress') {
                $oldPrimaryTrim = $oldValue.Trim()
                $newPrimaryTrim = $newValue.Trim()
                if ($oldPrimaryTrim -ne '' -and -not ($oldPrimaryTrim -ieq $newPrimaryTrim)) {
                    $choice = Show-KeepOrRemoveOldPrimaryDialog -Owner $mailForm -OldPrimary $oldPrimaryTrim -NewPrimary $newPrimaryTrim
                    if ($choice -ne 'Keep' -and $choice -ne 'Remove') {
                        $txtMailLog.Text = "Edit cancelled."
                        return
                    }
                    $keepOldPrimaryAsAlias = ($choice -eq 'Keep')
                }
            }
            $displayNew = if ($newValue.Trim() -eq '') { '(cleared)' } else { $newValue.Replace("`r`n", '; ') }
            $confirmExtra = switch ($attrName) {
                'PrimarySmtpAddress' {
                    if ($keepOldPrimaryAsAlias) { "`n`nThis rewrites the EmailAddresses collection: sets this as SMTP: (primary) and demotes the previous primary to smtp: (secondary alias)." }
                    else { "`n`nThis rewrites the EmailAddresses collection: sets this as SMTP: (primary) and REMOVES the previous primary address entirely." }
                }
                default { '' }
            }
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "Change '$attrLabel' for $script:Target ($($script:MailAttribsCurrentUser.UserPrincipalName)):`n`nFrom: $oldValue`nTo:   $displayNew$confirmExtra",
                'Confirm Attribute Edit',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                $txtMailLog.Text = "Cancelled."
                return
            }
            Set-MailAttribsAttribute -Identity $script:MailAttribsCurrentUser.Identity `
                -AttributeName $attrName -NewValue $newValue -CurrentProxyAddresses $script:MailAttribsCurrentUser.EmailAddresses `
                -KeepOldPrimaryAsAlias $keepOldPrimaryAsAlias
            $refreshed = Get-MailAttribsUser -Smtp $script:Target
            $script:MailAttribsCurrentUser = $refreshed
            $afterValue = if ($attrName -eq 'PrimarySmtpAddress') {
                Get-PrimarySmtpFromProxyAddresses -ProxyAddresses $refreshed.EmailAddresses
            } else {
                Format-AttributeValue -Value ($refreshed.$attrName)
            }
            Update-MailAttribsGrid -UserObject $refreshed
            $log = New-Object System.Text.StringBuilder
            [void]$log.AppendLine("--- BEFORE ---")
            [void]$log.AppendLine("$attrLabel = $oldValue")
            [void]$log.AppendLine("--- AFTER ---")
            [void]$log.AppendLine("$attrLabel = $afterValue")
            $txtMailLog.Text = $log.ToString().TrimEnd()
            $logVerb = if ($attrName -eq 'PrimarySmtpAddress') {
                "Edited Attribute $attrName ($(if($keepOldPrimaryAsAlias){'kept old primary as alias'}else{'removed old primary'}))"
            } else {
                "Edited Attribute $attrName"
            }
            Write-CompletedLog -Task "$logVerb $($script:MailAttribsCurrentUser.UserPrincipalName) (Mail Attribs; $script:Target)"
        } catch {
            $txtMailLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mail Attribs edit'
        } finally {
            $btnMailRefresh.Enabled = $true
            $btnAddSmtp.Enabled = $true
            $hasRow = ($grdMail.SelectedRows.Count -gt 0 -and $null -ne $script:MailAttribsCurrentUser)
            $isEditable = $false
            if ($hasRow) { $isEditable = ("$($grdMail.SelectedRows[0].Cells['_Editable'].Value)" -eq 'Y') }
            $btnMailEdit.Enabled = ($hasRow -and $isEditable)
        }
    }
    $grdMail.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $grdMail.ClearSelection()
        $grdMail.Rows[$e.RowIndex].Selected = $true
        Invoke-MailAttribsEditSelectedRow
    })
    $grdMail.Add_SelectionChanged({
        $hasRow = ($grdMail.SelectedRows.Count -gt 0 -and $null -ne $script:MailAttribsCurrentUser)
        $isEditable = $false
        if ($hasRow) { $isEditable = ("$($grdMail.SelectedRows[0].Cells['_Editable'].Value)" -eq 'Y') }
        $btnMailEdit.Enabled = ($hasRow -and $isEditable)
    })
    $btnMailEdit.Add_Click({ Invoke-MailAttribsEditSelectedRow })
    $btnAddSmtp.Add_Click({
        if ($null -eq $script:MailAttribsCurrentUser) {
            [void][System.Windows.Forms.MessageBox]::Show("Load mail attributes first (click Refresh).", 'No User Loaded', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $btnMailRefresh.Enabled = $false
        $btnAddSmtp.Enabled = $false
        try {
            $newAddr = $txtAddSmtp.Text.Trim()
            if ($newAddr -eq '') {
                [void][System.Windows.Forms.MessageBox]::Show("Enter an address to add.", 'Missing Input', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            if ($newAddr -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
                [void][System.Windows.Forms.MessageBox]::Show("'$newAddr' does not look like a valid email address (expected user@domain.tld).", 'Invalid Input', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            $beforeValue = Format-AttributeValue -Value $script:MailAttribsCurrentUser.EmailAddresses
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "Add 'smtp:$newAddr' as a secondary address for $script:Target ($($script:MailAttribsCurrentUser.UserPrincipalName))?`n`nThis does not change the primary address or remove any existing address.",
                'Confirm Add Secondary SMTP Address',
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) {
                $txtMailLog.Text = "Cancelled."
                return
            }
            Add-MailAttribsSecondaryProxyAddress -Identity $script:MailAttribsCurrentUser.Identity `
                -CurrentProxyAddresses $script:MailAttribsCurrentUser.EmailAddresses -NewSecondary $newAddr
            $refreshed = Get-MailAttribsUser -Smtp $script:Target
            $script:MailAttribsCurrentUser = $refreshed
            $afterValue = Format-AttributeValue -Value $refreshed.EmailAddresses
            Update-MailAttribsGrid -UserObject $refreshed
            $log = New-Object System.Text.StringBuilder
            [void]$log.AppendLine("--- BEFORE (EmailAddresses) ---")
            [void]$log.AppendLine($beforeValue)
            [void]$log.AppendLine("--- AFTER ---")
            [void]$log.AppendLine($afterValue)
            $txtMailLog.Text = $log.ToString().TrimEnd()
            Write-CompletedLog -Task "Added Secondary SMTP $($script:MailAttribsCurrentUser.UserPrincipalName) (Mail Attribs; $script:Target; smtp:$newAddr)"
            $txtAddSmtp.Text = ''
        } catch {
            $txtMailLog.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Add Secondary SMTP'
        } finally {
            $btnMailRefresh.Enabled = $true
            $btnAddSmtp.Enabled = $true
        }
    })
    $btnMailClose.Add_Click({ $mailForm.Close() })
    $mailForm.StartPosition = 'Manual'
    Set-ConsoleTheme -Form $mailForm
    # BEFORE Add-ToolHelpButton: the bar shifts every existing control down
    # and grows the form, and the help button places itself from ClientSize at
    # run time.
    $bar = Add-ToolTargetBar -Form $mailForm -Label 'Mailbox or address:' -InitialValue $script:Target
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        Update-MailAttribsFromCurrent
    })

    Add-ToolHelpButton -Form $mailForm -ScriptPath $PSCommandPath -Title 'Mail Attributes (Rename)' | Out-Null
    Set-FormScreenCorner -Form $mailForm -Corner 'BottomLeft'
    [void]$mailForm.ShowDialog()
}
#endregion

Show-MailAttribsForm
```


<br>

==============================================================================

# ◆ FILE 8 OF 8 — `ExternalForwarding.ps1`

> **405 lines.** A CSV REPORT end to end: the shared report-file controls, what "external" means and why it depends on a complete accepted-domain list (C3 -- a truncated list turns ordinary internal mail into a false exfiltration finding), partial-run labelling, and never writing a zero-row CSV.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```powershell
#Requires -Version 5.1
<#
    ExternalForwarding.ps1
    Finds mail being forwarded outside the tenant, from both mailbox settings and user inbox rules, and exports to CSV. Author : Glenn Armbruster

    Reasoning: ScriptNotes/ExternalForwarding.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-07-31
#>
# TOOLKIT-HELP-BEGIN
# Finds mail leaving the organisation, from BOTH places it can be set up:
#
#   * mailbox forwarding, configured by an admin or by the user
#   * inbox rules, which users create themselves and which are the ones nobody
#     thinks to check
#
# The classic sign of a compromised mailbox is a quiet rule forwarding
# everything to an outside address. This is the report that finds it.
#
# Writes a CSV. Read-only.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

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
$ReportDir = Get-ReportPath

$AllowlistedDomains = @()
#endregion

#region Data -----------------------------------------------------------------
function Test-ExternalDomain {
    param([string]$Domain, [string[]]$Accepted)
    if (-not $Domain) { return $false }
    return (-not ($Accepted -contains $Domain.ToLower()))
}

function Get-MailboxForwarding {
    param([string[]]$Accepted, [scriptblock]$Progress)

    $cmd = if (Get-Command 'Get-EXOMailbox' -ErrorAction SilentlyContinue) {
               'Get-EXOMailbox'
           } else { 'Get-Mailbox' }

    $props = @('PrimarySmtpAddress', 'DisplayName', 'RecipientTypeDetails',
               'ForwardingAddress', 'ForwardingSmtpAddress',
               'DeliverToMailboxAndForward')

    & $Progress 'Reading mailbox-level forwarding...' 5
    $all = @(Invoke-EXOBulkQuery -Cmdlet $cmd -Properties $props -Progress $Progress)

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($m in $all) {
        foreach ($pair in @(
            @{ Val = "$($m.ForwardingSmtpAddress)"; Kind = 'Mailbox ForwardingSmtpAddress' },
            @{ Val = "$($m.ForwardingAddress)";     Kind = 'Mailbox ForwardingAddress' }
        )) {
            if (-not $pair.Val) { continue }
            $dom = Get-SmtpDomain $pair.Val
            if (-not $dom) { continue }          # unresolvable / legacy DN
            if (-not (Test-ExternalDomain -Domain $dom -Accepted $Accepted)) { continue }

            $rows.Add([PSCustomObject][ordered]@{
                PrimarySmtpAddress = "$($m.PrimarySmtpAddress)"
                DisplayName        = "$($m.DisplayName)"
                RecipientType      = "$($m.RecipientTypeDetails)"
                Source             = $pair.Kind
                RuleName           = ''
                RuleEnabled        = ''
                Action             = 'Forward'
                Destination        = $pair.Val
                DestinationDomain  = $dom
                KeepsCopy          = [bool]$m.DeliverToMailboxAndForward
                Allowlisted        = ($AllowlistedDomains -contains $dom)
            })
        }
    }
    return $rows
}

function Get-InboxRuleForwarding {
    param(
        [object[]]$Mailboxes,
        [string[]]$Accepted,
        [scriptblock]$Progress
    )
    $rows   = New-Object System.Collections.Generic.List[object]
    $failed = @()
    $n = 0
    $total = @($Mailboxes).Count

    foreach ($mb in $Mailboxes) {
        $n++
        if ($n % 10 -eq 0 -or $n -eq $total) {
            $pct = 20 + [int](75.0 * $n / [Math]::Max(1, $total))
            & $Progress "Inbox rules: $n of $total -- $($rows.Count) external found..." $pct
        }
        $smtp = "$($mb.PrimarySmtpAddress)"
        if (-not $smtp) { continue }

        try {
            $rules = @(Invoke-EXOCommand -Name 'Get-InboxRule' `
                        -Parameters @{ Mailbox = $smtp; ErrorAction = 'Stop' })
        } catch {
            $failed += $smtp
            continue
        }

        foreach ($r in $rules) {
            foreach ($act in @(
                @{ Prop = 'ForwardTo';             Label = 'Forward' },
                @{ Prop = 'ForwardAsAttachmentTo'; Label = 'Forward as attachment' },
                @{ Prop = 'RedirectTo';            Label = 'Redirect (no copy kept)' }
            )) {
                $targets = @($r.($act.Prop))
                foreach ($t in $targets) {
                    if (-not $t) { continue }
                    $dom = Get-SmtpDomain "$t"
                    if (-not $dom) { continue }
                    if (-not (Test-ExternalDomain -Domain $dom -Accepted $Accepted)) { continue }

                    $rows.Add([PSCustomObject][ordered]@{
                        PrimarySmtpAddress = $smtp
                        DisplayName        = "$($mb.DisplayName)"
                        RecipientType      = "$($mb.RecipientTypeDetails)"
                        Source             = 'Inbox rule'
                        RuleName           = "$($r.Name)"
                        RuleEnabled        = [bool]$r.Enabled
                        Action             = $act.Label
                        Destination        = "$t"
                        DestinationDomain  = $dom
                        KeepsCopy          = ($act.Prop -ne 'RedirectTo')
                        Allowlisted        = ($AllowlistedDomains -contains $dom)
                    })
                }
            }
        }
    }
    return [PSCustomObject]@{ Rows = $rows; Failed = $failed }
}

function Get-ForwardingReport {
    param([scriptblock]$Progress, [switch]$ScanInboxRules)

    & $Progress 'Reading accepted domains...' 2
    $accepted = Get-AcceptedDomainList
    if ($accepted.Count -eq 0) {
        throw "Could not read accepted domains. Without them there is no way to tell external from internal, so the report would be meaningless."
    }

    $rows   = New-Object System.Collections.Generic.List[object]
    $failed = @()

    foreach ($r in (Get-MailboxForwarding -Accepted $accepted -Progress $Progress)) {
        $rows.Add($r)
    }

    $scanned = 0
    if ($ScanInboxRules) {
        & $Progress 'Listing mailboxes for inbox-rule scan...' 15
        $cmd = if (Get-Command 'Get-EXOMailbox' -ErrorAction SilentlyContinue) {
                   'Get-EXOMailbox'
               } else { 'Get-Mailbox' }
        $mbs = @(Invoke-EXOBulkQuery -Cmdlet $cmd `
                    -Properties @('PrimarySmtpAddress', 'DisplayName', 'RecipientTypeDetails') `
                    -Progress $Progress)
        $scanned = $mbs.Count
        $res = Get-InboxRuleForwarding -Mailboxes $mbs -Accepted $accepted -Progress $Progress
        foreach ($r in $res.Rows) { $rows.Add($r) }
        $failed = $res.Failed
    }

    $sorted = @($rows | Sort-Object `
        @{ Expression = { $_.Allowlisted } },
        @{ Expression = { $_.KeepsCopy } },
        @{ Expression = { $_.DestinationDomain } },
        @{ Expression = { $_.PrimarySmtpAddress } })

    return [PSCustomObject]@{
        Rows            = $sorted
        AcceptedDomains = $accepted
        ScannedRules    = $ScanInboxRules.IsPresent
        MailboxesScanned = $scanned
        Failed          = $failed
    }
}

function Format-ReportSummary {
    param($Result)
    $rows = @($Result.Rows)
    $sb = New-Object System.Text.StringBuilder

    $live = @($rows | Where-Object { -not $_.Allowlisted })
    $noCopy = @($live | Where-Object { -not $_.KeepsCopy })
    $ruleBased = @($live | Where-Object { $_.Source -eq 'Inbox rule' })

    [void]$sb.AppendLine("External forwarding found : $($rows.Count)")
    [void]$sb.AppendLine("  Not allowlisted         : $($live.Count)")
    [void]$sb.AppendLine("  No copy kept in mailbox : $($noCopy.Count)   <-- highest priority")
    [void]$sb.AppendLine("  From inbox rules        : $($ruleBased.Count)")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("Distinct mailboxes affected: $(@($live | Select-Object -ExpandProperty PrimarySmtpAddress -Unique).Count)")
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine('By source:')
    $live | Group-Object Source | Sort-Object Count -Descending | ForEach-Object {
        [void]$sb.AppendLine(("  {0,-32}{1,6}" -f $_.Name, $_.Count))
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Top destination domains:')
    $live | Group-Object DestinationDomain | Sort-Object Count -Descending |
        Select-Object -First 15 | ForEach-Object {
            [void]$sb.AppendLine(("  {0,-32}{1,6}" -f $_.Name, $_.Count))
        }

    [void]$sb.AppendLine('')
    if ($Result.ScannedRules) {
        [void]$sb.AppendLine("Inbox rules scanned across $($Result.MailboxesScanned) mailbox(es).")
    } else {
        [void]$sb.AppendLine('*** Inbox rules were NOT scanned. ***')
        [void]$sb.AppendLine('  Mailbox-level forwarding only. Rule-based forwarding is the more')
        [void]$sb.AppendLine('  common exfiltration route -- tick the box and re-run for full coverage.')
    }
    [void]$sb.AppendLine("Accepted domains treated as internal: $($Result.AcceptedDomains.Count)")

    if ($Result.Failed -and $Result.Failed.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("*** INCOMPLETE -- rules could not be read for $($Result.Failed.Count) mailbox(es): ***")
        foreach ($f in ($Result.Failed | Select-Object -First 20)) { [void]$sb.AppendLine("  $f") }
        if ($Result.Failed.Count -gt 20) {
            [void]$sb.AppendLine("  ... and $($Result.Failed.Count - 20) more")
        }
    }
    return $sb.ToString()
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Report -- External Forwarding'
$form.Size            = New-Object System.Drawing.Size(760, 676)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Finds mail forwarded outside the tenant, from mailbox settings and from user inbox rules.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblInfo)

$lblTenant           = New-Object System.Windows.Forms.Label
$lblTenant.Tag       = 'NoTheme'
$lblTenant.Location  = New-Object System.Drawing.Point(12, 32)
$lblTenant.Size      = New-Object System.Drawing.Size(720, 18)
$lblTenant.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblTenant)

$chkRules          = New-Object System.Windows.Forms.CheckBox
$chkRules.Text     = 'Also scan inbox rules (slow -- one call per mailbox, allow 30-60 min)'
$chkRules.Location = New-Object System.Drawing.Point(12, 56)
$chkRules.Size     = New-Object System.Drawing.Size(500, 20)
$form.Controls.Add($chkRules)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run Report'
$btnRun.Location = New-Object System.Drawing.Point(12, 82)
$btnRun.Size     = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnRun)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Reports Folder'
$btnOpen.Location = New-Object System.Drawing.Point(140, 82)
$btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnOpen)


$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(622, 82)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 156)
$lblStatus.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 176 -Width 720 -Height 12

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Location   = New-Object System.Drawing.Point(12, 194)
$txtOut.Size       = New-Object System.Drawing.Size(720, 420)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'ExternalForwarding' `
                -X 12 -Y 118 -Width 720 -StatusLabel $lblStatus


$btnOpen.Add_Click({ Open-ReportFolder -Path $ReportDir })

$btnRun.Add_Click({
    if ($chkRules.Checked) {
        $ans = [System.Windows.Forms.MessageBox]::Show(
            "Scanning inbox rules queries every mailbox individually and can take 30-60 minutes on a tenant this size.`n`nThe window will stay responsive and show progress. Continue?",
            'Long-running scan', 'YesNo', 'Question')
        if ($ans -ne 'Yes') { return }
    }

    $btnRun.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = ''
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar `
                        -Activity 'External forwarding report'
        $report = Get-ForwardingReport -Progress $progress -ScanInboxRules:$chkRules.Checked

        Complete-ProgressReporter -Activity 'External forwarding report' -Bar $pbar

        if (@($report.Rows).Count -eq 0) {
            $lblStatus.Text = 'No external forwarding found -- nothing written.'
            $txtOut.Text = (Format-ReportSummary -Result $report) +
                           [Environment]::NewLine +
                           'No CSV was written because there were no rows.'
            return
        }

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir "ExternalForwarding_$stamp.csv"
        $report.Rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $files.SelectPath($csv)
        $lblStatus.Text  = "Done. $(@($report.Rows).Count) row(s) written."

        $txtOut.Text = (Format-ReportSummary -Result $report) +
                       [Environment]::NewLine +
                       'Saved to:' + [Environment]::NewLine + "  $csv"

        $note = @()
        if (-not $report.ScannedRules) { $note += 'mailbox-level only' }
        if ($report.Failed.Count)      { $note += "$($report.Failed.Count) unreadable" }
        # 'PARTIAL' is the keyword the log is scanned for.
        $suffix = if ($note.Count) { " -- PARTIAL ($($note -join ', '))" } else { '' }
        # PARTIAL is the keyword the log is scanned for; Result is the same
        # fact in a column a filter can see.
        if ($suffix) { $logResult = 'Partial' } else { $logResult = 'Success' }
        Write-CompletedLog -Task "External forwarding report -- $(@($report.Rows).Count) findings$suffix" -Result $logResult
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'External forwarding report'
    } finally {
        Complete-ProgressReporter -Activity 'External forwarding report'
        $form.Cursor    = $prev
        $btnRun.Enabled = $true
    }
})

$d = Get-ConnectedTenantDomain
if ($d) {
    $lblTenant.Text      = "Tenant: $d"
    $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(0, 128, 0)
} else {
    $lblTenant.Text      = 'Tenant: not connected -- connect to Exchange Online first.'
    $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(200, 0, 0)
}

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled, and BringToFront'd inside the helper because WinForms
# draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'External Forwarding' | Out-Null

[void]$form.ShowDialog()
#endregion
```


<br>

==============================================================================

# END OF REFERENCE CODE, PART 2

Eight files. Part 1 is `DEVIN-REFERENCE-CODE.md`; the instructions that
govern both are in `DEVIN-PACK.md`.
