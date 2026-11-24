#Requires -Version 5.1
# TOOLKIT-VARIANT: Dual
#
# CLOUD BY DEFAULT, ON-PREMISES ON A SWITCH. The window opens against
# Exchange Online and does everything it has always done. The Editing switch
# in the top right points it at the on-premises object instead, which is the
# only place a synced mailbox's addresses, alias, display name, hide-from-GAL
# and directory fields can actually be changed.
#
# It DEGRADES rather than refuses: a tenant with no on-premises Exchange, or
# a window where on-premises is not connected, gets the switch greyed out
# with the reason on it, and the cloud side works exactly as before. Nobody
# loses the tool because a server is down.
#
# Nothing is called unqualified. Every cloud call goes through
# Invoke-EXOCommand and every on-premises call through Invoke-OnPremCommand,
# so which Exchange answers is decided here and not by import order.
<#
    MailboxEditor.ps1
    One mailbox, every tab. The tool you open when you have an address and a
    question, instead of having to know the verb first.

    Author   : Glenn Armbruster
    Created  : 2026-09-18

    WHY THIS EXISTS. Fourteen single-purpose tools -- permissions, calendar
    permissions, delegates, forwarding, hide/unhide, convert type, archive,
    hold, retention, automap, health, stats, out of office, recover deleted --
    came to 5,580 lines, of which a large share was the same scaffolding
    repeated: resolve the address, build a form, a status line, an output box,
    error handling, an audit line. The logic was never the bulk of it.

    The cost was not the line count. It was that the operator had to decide
    'is this a permissions question or a mail-flow question' BEFORE they could
    open anything. You have an address; you should be able to look at the
    mailbox and decide afterwards.

    THREE RULES THIS TOOL IS BUILT ON, and each one is a decision that could
    have gone the other way:

    1. PER-SECTION APPLY, NEVER A GLOBAL SAVE. Every block has its own Apply
       that does exactly that one thing and writes its own audit line. A tab
       with ten fields and one Save button produces 'I changed six things, hit
       Save, two failed, which two?' -- and leaves the operator unable to
       answer it. A half-finished edit here costs nothing.

    2. THE LOGIC IS IN Common.ps1. This file is a UI shell. Unified Lookup
       edits some of the same things and calls the same functions, so there is
       one implementation and one audit shape rather than two of each.

    3. TABS LOAD WHEN OPENED. Loading everything up front is twenty-odd round
       trips and half a minute before anything appears. Overview loads on open;
       every other tab loads the first time it is clicked, and caches.

    AND THE THING THAT MAKES IT SAFE ON A HYBRID MAILBOX. Several fields are
    mastered in Active Directory for a synced mailbox. Exchange Online accepts
    the write, reports success, and the next sync puts the old value back --
    silently, within about half an hour. Get-MailboxWriteTarget in Common.ps1
    declares where every field lives, and this tool refuses or warns rather
    than letting an operator watch a change succeed and then quietly undo
    itself.
#>
# TOOLKIT-HELP-RUNS: Reads one mailbox across several tabs, in Exchange Online or -- on the Editing switch -- against the on-premises object. Each section's Apply changes only that section, and the risky ones ask first.
# TOOLKIT-HELP-BEGIN
# Everything about one mailbox, in tabs. Type an address, press Load.
#
#   Overview     Who it is, what type, size and quota, archive size and
#                whether auto-expanding is on, last logon, licence,
#                and the directory fields (title, department, manager). Also
#                the three things people change most: hide from the address
#                book, convert the mailbox type, and enable the archive.
#   Health       A read-only verdict. Protocols, account state, hybrid GUIDs,
#                and any inbox rule that sends mail out of the tenant.
#   Addresses    Every SMTP address, which is primary, and changing it.
#   Access       Full Access, Send As, Send on Behalf -- who has them, add and
#                remove -- plus the two 'keep a copy in Sent Items' settings.
#   Automap Fix  A mailbox that will not appear in someone's Outlook. Tick the
#                person, and their Full Access is removed, held ten seconds,
#                granted again with automapping on (or off, to keep access
#                but stop it opening), and read back -- with what to tell
#                them afterwards. Only direct Full Access is listed; access
#                through a group never automaps.
#   Folder       Calendar, Inbox, Contacts, Tasks and Notes permissions, and
#   Permissions  the calendar delegates.
#   Mail Flow    Forwarding, and who is allowed to send to this mailbox.
#   Retention    Litigation hold, retention hold, and the retention policy.
#   Automatic    The full out-of-office editor -- state, schedule, audience,
#   Replies      and both messages with formatting, templates and a preview.
#   Inbox Rules  List them, read one, turn one off, delete one.
#   Recovery     The dumpster: what is recoverable and how much of it.
# THE ARCHIVE HAS TWO SWITCHES ON Overview, and both are one-way. Enable
# archive creates the online archive. Enable auto-expanding turns on automatic
# extra storage, and appears only when there IS an archive and auto-expanding
# is known to be off -- the reference requires the archive first, and states
# that auto-expanding cannot be disabled once on. If its state could not be
# read the button stays out: a one-way change is not made on a guess. There is
# also a TENANT-WIDE auto-expanding switch (Set-OrganizationConfig), which this
# window deliberately does not touch -- same irreversibility, every mailbox.
#
# PROTOCOL ACCESS, on the Health tab, is editable. Outlook on the web, MAPI,
# EWS, ActiveSync, POP and IMAP are drop-downs with their own Apply, written
# with Set-CASMailbox. Only what you changed is sent, and the confirmation
# spells out the direction that bites: turning MAPI or OWA off stops that
# person working, and turning POP or IMAP on re-opens a legacy protocol that
# was almost certainly off on purpose. A protocol whose state could not be
# read shows blank and is left alone rather than being written as off.
#
# DESCRIPTION AND NOTES, on Overview, are two different fields and are
# labelled the way ADUC labels them. Description is the one-line box on the
# General tab (LDAP 'description'), the field most AD reports read; it is
# written straight to Active Directory because no Exchange cmdlet can write it
# on-premises, and it is HIDDEN rather than greyed whenever there is no
# directory object behind the mailbox -- a cloud-only tenant, a cloud-native
# mailbox, or cloud mode on a synced one. Notes is the big box on the
# Telephones tab, written with Set-User, which works on-premises and in
# Exchange Online both. Apply directory fields writes whichever of them
# changed; the directory write goes first, so a directory failure stops the
# apply rather than half-completing it.
#
#   Recover      Individual deleted items, including after Deleted Items has
#   Items        been emptied -- the same job as the EAC's recovery screen.
#                Searches Deleted Items, Recoverable Items and Purges together
#                and says which of the three each item is in; DiscoveryHolds is
#                there as its own pick.
#                'Purge in' is how long before the service deletes the item
#                for good, worked out from the mailbox's deleted-item
#                retention; it reads 'on hold' when a hold stops that clock.
#                From and To are shown when the service returns them.
#   Quarantine   Quarantined mail for this mailbox, with release.
#   Statistics   Sizes, item counts and folder breakdown.
#
# EACH BLOCK HAS ITS OWN APPLY. There is no global Save. Whatever you press
# does that one thing and logs it; nothing else on the tab is touched.
#
# THE EDITING SWITCH, TOP RIGHT. Cloud is the default and is what you want
# nearly always. Set it to On-premises when the mailbox is synced from Active
# Directory and you need to change something the cloud will not let you:
# addresses, alias, hide-from-address-list, title, department, office or
# manager. In that mode the window is pointed at the on-premises object, the
# banner turns amber and says so, and the tabs that only exist in Exchange
# Online -- Access, Folder Permissions, Mail Flow, Retention, Automatic
# Replies, Inbox Rules, Recovery, Recover Items, Quarantine -- are greyed out.
#
# CONVERT IS THE ONE CONTROL THAT IGNORES THAT SWITCH, and it is deliberate.
# Mailbox type does not sync cloud-to-on-premises, so converting in Exchange
# Online alone is undone by the next directory sync. Convert writes BOTH
# directories whichever mode you are in -- on-premises first -- and the
# confirmation tells you exactly what it is about to do on each side. If
# on-premises cannot be reached it says only the cloud will change, and that a
# sync will probably reverse it, and lets you decide.
#
# AN ON-PREMISES CHANGE IS NOT INSTANT. Exchange Online will not show it until
# the next directory sync. Check back in about half an hour before deciding it
# did not work.
#
# If the switch is greyed out, this window has no on-premises Exchange
# connection (or the tenant has none); the reason is on the status line.
#
# SYNCED MAILBOXES. If the mailbox came from Active Directory, some fields
# cannot be changed here -- Exchange Online would accept the change and the
# next directory sync would put it back, usually within half an hour. Those
# fields are greyed with the reason, and a few that vary by hybrid
# configuration are marked as needing a re-check after a sync cycle. That is
# deliberate: being told it might not stick is better than watching it revert.
#
# Nothing is changed until you press an Apply.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
if (-not (Test-ConsoleLaunch)) { return }

#region State -------------------------------------------------------------------
# ONE PLACE FOR 'WHICH MAILBOX AND WHAT DO WE KNOW ABOUT IT'. Every tab reads
# $script:Mbx; nothing else holds a copy, so a reload cannot leave one tab
# showing the previous mailbox.
$script:Mbx     = $null    # Resolve-MailboxTarget result
$script:Loaded  = @{}      # tab key -> $true once its data has been fetched
$script:Busy    = $false
$script:Closing = $false   # set by FormClosing; see the tab handler at the foot of the file
# Set by any tab read that CATCHES a failure instead of throwing. Without it a
# tab whose data could not be read is cached as loaded and stays empty.
$script:TabReadFailed = $false

# WHICH DIRECTORY THIS WINDOW IS WRITING TO. 'Cloud' or 'OnPrem'. Every gate,
# every Apply and the identity banner read it, so there is one answer and it
# is on screen.
$script:Mode       = 'Cloud'
$script:OnPrem     = $null   # Resolve-OnPremMailboxTarget result
$script:OnPremWhy  = ''      # why the switch is unavailable, if it is
$script:OvDescDn   = ''      # DN the Description box is bound to, '' when hidden
$script:OvDescMulti = $false # description carries several values -- read-only
$script:OvDirWas   = $null   # the directory values the boxes were filled from
$script:OvDirCloud = $false  # $true when Apply should write to Exchange Online
$script:CasWas     = $null   # the protocol states the combos were filled from
$script:IdWas      = $null   # display name / alias the boxes were filled from
$script:ArchAuto   = 'unknown'  # auto-expanding archive: on | off | unknown
$script:Licence    = $null   # Get-UserLicenseState result, read on Overview
$script:Health     = $null   # the Health tab's snapshot, for the verdict
#endregion

#region Form shell --------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Mailbox Editor'
$form.ClientSize    = New-Object System.Drawing.Size(1120, 760)
$form.StartPosition = 'Manual'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$y = 12
$lblWho          = New-Object System.Windows.Forms.Label
$lblWho.Text     = 'Mailbox:'
$lblWho.Location = New-Object System.Drawing.Point(12, ($y + 3))
$lblWho.Size     = New-Object System.Drawing.Size(60, 18)
$form.Controls.Add($lblWho)

$txtMbx          = New-Object System.Windows.Forms.TextBox
$txtMbx.Location = New-Object System.Drawing.Point(74, $y)
$txtMbx.Size     = New-Object System.Drawing.Size(340, 24)
$txtMbx.Font     = New-Object System.Drawing.Font('Consolas', 9)
$txtMbx.Text     = "$(Get-TargetSmtp -Text $SmtpAddress)"
$form.Controls.Add($txtMbx)

$btnLoad          = New-Object System.Windows.Forms.Button
$btnLoad.Text     = 'Load'
$btnLoad.Location = New-Object System.Drawing.Point(424, ($y - 2))
$btnLoad.Size     = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnLoad)

# THE EDITING SWITCH, FROM THE SHARED BUILDER RATHER THAN TYPED OUT HERE.
#
# This window had its own copy -- same coordinates, same two entries, same
# order -- written before New-EditorModeSwitch existed and left behind when
# the other three editors moved onto it. It was the fourth copy of a control
# that decides WHICH DIRECTORY A WRITE LANDS IN, which is the exact thing
# that builder's docstring exists to stop.
#
# It cost something real on 2026-09-19: a cloud-only tenant has no second
# directory to choose, so the builder replaces the whole control with the
# words 'Cloud Tenant' -- and this window, not calling it, went on offering
# a choice between one real destination and one that does not exist here.
$cmbMode = New-EditorModeSwitch -Form $form -X 718 -Y $y -LabelX 660

# CLOSE MOVED LEFT. Add-ToolHelpButton places the '?' in the first clear spot
# it can find; with Close at 1018 the only gap left was 12px wide, so the two
# ended up on top of each other. This leaves the corner free for it.
$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(952, ($y - 2))
$btnClose.Size     = New-Object System.Drawing.Size(90, 28)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)
$y += 34

# THE IDENTITY BANNER. Always visible, whatever tab is open, because every
# Apply on every tab acts on this mailbox and the operator should never have
# to go back to a tab to check which one.
$lblIdent           = New-Object System.Windows.Forms.Label
$lblIdent.Tag       = 'NoTheme'
$lblIdent.Location  = New-Object System.Drawing.Point(12, $y)
$lblIdent.Size      = New-Object System.Drawing.Size(1096, 18)
$lblIdent.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblIdent.ForeColor = (Get-ToolStatusColour -Kind Muted)
$lblIdent.Text      = 'No mailbox loaded.'
$form.Controls.Add($lblIdent)
$y += 20

# THE SYNC BANNER, separate from the identity line and coloured, because it
# changes what the tool will let you do. A synced mailbox is the case where an
# apparently successful write reverts half an hour later.
$lblSync           = New-Object System.Windows.Forms.Label
$lblSync.Tag       = 'NoTheme'
$lblSync.Location  = New-Object System.Drawing.Point(12, $y)
$lblSync.Size      = New-Object System.Drawing.Size(1096, 18)
$lblSync.ForeColor = (Get-ToolStatusColour -Kind Muted)
$lblSync.Text      = ''
$form.Controls.Add($lblSync)
$y += 22

$tabs          = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(12, $y)
$tabs.Size     = New-Object System.Drawing.Size(1096, 640)
$form.Controls.Add($tabs)

$lblStatus           = New-Object System.Windows.Forms.Label
$lblStatus.Location  = New-Object System.Drawing.Point(12, ($y + 646))
$lblStatus.Size      = New-Object System.Drawing.Size(1096, 18)
$lblStatus.ForeColor = (Get-ToolStatusColour -Kind Muted)
$lblStatus.Text      = 'Enter a mailbox and press Load.'
$form.Controls.Add($lblStatus)

$form.ClientSize = New-Object System.Drawing.Size(1120, ($lblStatus.Bottom + 12))


# ---- register with the shared editor runtime ----------------------------------
# Everything that used to be nine local functions now lives in Common.ps1 and
# is driven by this one call. The hooks below are the only things about this
# window that differ from the Room, Group and Person editors.
Initialize-TabbedEditor -Name 'Mailbox' -Noun 'mailbox' -ObjectName 'Mbx' `
    -Form $form -Tabs $tabs -StatusLabel $lblStatus -IdentityLabel $lblIdent -SyncLabel $lblSync -ModeCombo $cmbMode `
    -CloudOnlyTabs @('Access', 'Automap', 'Folders', 'MailFlow', 'Retention', 'Replies', 'Rules', 'Recovery', 'Restore', 'Quarantine', 'Stats', 'Health') `
    -CloudGridNames @('gridAccess', 'gridAutomap', 'gridFolder', 'gridStats', 'gridRules', 'gridRestore', 'gridQuar') `
    -TabLoaders @{
        Overview  = { Update-OverviewTab }
        Addresses = { Update-AddressesTab }
        Access    = { Update-AccessTab }
        Automap   = { Update-AutomapTab }
        Folders   = { Update-FoldersTab }
        MailFlow  = { Update-MailFlowTab }
        Retention = { Update-RetentionTab }
        Replies   = { Update-RepliesTab }
        Recovery  = { Update-RecoveryTab }
        Rules     = { Update-RulesTab }
        Health    = { Update-HealthTab }
    } `
    -Identity {
        param($m)
        $sync = $(if ($m.IsSynced) {
                if ($script:OnPremWhy) { "SYNCED FROM ACTIVE DIRECTORY -- addresses, display name, alias and hide-from-GAL are mastered on-premises. On-premises editing is unavailable here: $($script:OnPremWhy)" }
                else { 'SYNCED FROM ACTIVE DIRECTORY -- addresses, display name, alias and hide-from-GAL are mastered on-premises. Set Editing (top right) to On-premises to change them.' }
            } else { 'Cloud-only mailbox -- every field on every tab is editable here.' })
        @{ Text = "$($m.Display)   <$($m.Smtp)>   $($m.Type)"; Colour = 'Good'
           Sync = $sync; SyncColour = $(if ($m.IsSynced) { 'Warn' } else { 'Muted' }) }
    }

$tabOverview  = New-EditorTab -Key 'Overview'  -Text 'Overview'
# SECOND, because the question after 'who is this' is 'is anything wrong with
# it'. Health carries only what no other tab shows.
$tabHealth    = New-EditorTab -Key 'Health'    -Text 'Health'
$tabAddresses = New-EditorTab -Key 'Addresses' -Text 'Email Address'
# THE TWO PERMISSION TABS SIT TOGETHER. Full Access is mailbox-wide, Folder
# Perm is per-folder, and "who can see this mailbox" is usually answered by
# looking at both -- so Automap Fix no longer sits between them.
#
# KEYS ARE NOT CAPTIONS. Every -Key below is unchanged: the Key becomes the
# TabPage's Name, and Invoke-EditorTabLoad dispatches the lazy load on that.
# Renaming a caption is cosmetic; renaming a Key would silently stop a tab
# ever loading.
$tabAccess    = New-EditorTab -Key 'Access'    -Text 'Full Access'
$tabFolders   = New-EditorTab -Key 'Folders'   -Text 'Folder Perm'
$tabAutomap   = New-EditorTab -Key 'Automap'   -Text 'Automap Fix'
$tabMailFlow  = New-EditorTab -Key 'MailFlow'  -Text 'Mail Flow'
$tabRetention = New-EditorTab -Key 'Retention' -Text 'Retention'
$tabReplies   = New-EditorTab -Key 'Replies'   -Text 'Auto Replies'
$tabRules     = New-EditorTab -Key 'Rules'     -Text 'Inbox Rules'
$tabRecovery  = New-EditorTab -Key 'Recovery'  -Text 'Recovery'
$tabRestore   = New-EditorTab -Key 'Restore'   -Text 'Recover Items'
$tabQuar      = New-EditorTab -Key 'Quarantine' -Text 'Quarantine'
# LAST, AND IT LOADS ON A BUTTON. Per-folder statistics are the slowest read in
# this window by a wide margin, and nothing else should ever wait behind them.
$tabStats     = New-EditorTab -Key 'Stats'     -Text 'MBX Stats'
#endregion

#region Shell helpers -----------------------------------------------------------





function Test-EditorFieldWritable {
    <#
        The hybrid gate, in front of every write.

        Directory  refuse, and say where the field actually lives
        Warn       proceed, but tell the operator it may not survive a sync
        Ok         proceed silently

        Returns $true when the caller may go ahead.
    #>
    param([Parameter(Mandatory)][string]$Field)

    # ON-PREMISES MODE INVERTS THE GATE. The fields the cloud side refuses are
    # exactly the ones this mode exists to write, and the cloud-mastered ones
    # are the ones it cannot touch -- litigation hold, folder permissions and
    # the rest are properties of the Exchange Online mailbox and have no
    # on-premises equivalent to write to.
    if ($script:Mode -eq 'OnPrem') {
        if (-not $script:OnPrem -or -not $script:OnPrem.Ok) {
            Set-EditorStatus -Text 'No on-premises object is loaded.' -Bad
            return $false
        }
        # ONLY 'Directory' IS ALLOWED THROUGH. Letting 'Unverified' pass was
        # wrong and dangerous: every Unverified field -- forwarding, delivery
        # restrictions, send-on-behalf -- is handled by a cloud handler that
        # calls Invoke-EXOCommand directly. Allowing them here meant the gate
        # said yes, the write went to Exchange Online, and the audit line
        # stamped it OnPrem. Unverified means we do not know; the answer to
        # not knowing is not to write.
        $o = Get-MailboxFieldOrigin -Field $Field
        if ($o -ne 'Directory') {
            $m = "'$Field' is not a field on-premises masters, so this mode cannot change it. Switch Editing back to Cloud."
            $null = Show-ToolDialog($m, 'Not an on-premises field', 'OK', 'Warning')
            Set-EditorStatus -Text $m -Bad
            return $false
        }
        return $true
    }

    $t = Get-MailboxWriteTarget -Field $Field -IsSynced $script:Mbx.IsSynced
    if ($t.State -eq 'Directory') {
        # THE REFUSAL NOW POINTS AT THE ANSWER. Saying only "this is mastered
        # in Active Directory" left the operator with a true statement and
        # nowhere to go; the switch that fixes it is in this window.
        $extra = $(if ($script:OnPremWhy) { "`r`n`r`nOn-premises editing is unavailable in this window: $($script:OnPremWhy)" }
                   else { "`r`n`r`nSet Editing (top right) to On-premises to change it here." })
        $null = Show-ToolDialog(($t.Reason + $extra), 'Mastered in Active Directory', 'OK', 'Warning')
        Set-EditorStatus -Text $t.Reason -Bad
        return $false
    }
    if ($t.State -eq 'Warn') {
        $a = Show-ToolDialog(($t.Reason + "`r`n`r`nGo ahead?"), 'May not persist', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return $false }
    }
    return $true
}

function Invoke-EditorSetMailbox {
    <#
        ONE Set, ROUTED TO WHICHEVER DIRECTORY THE WINDOW IS POINTED AT.

        Every address and visibility handler calls this instead of naming a
        cmdlet. Written the other way -- an if/else inside each handler -- the
        four handlers would each have their own idea of which cmdlet on-prem
        wants, and the one that got it wrong would write to the cloud object
        while the banner said on-premises.

        The Identity is supplied HERE, from the object that mode actually
        loaded. A caller passing its own would be passing the cloud address to
        an on-premises cmdlet, which is exactly the class of mistake this
        window is meant to stop.
    #>
    param([Parameter(Mandatory)][hashtable]$Parameters)
    $p = @{} + $Parameters
    $p['ErrorAction'] = 'Stop'
    if ($script:Mode -eq 'OnPrem') {
        if (-not $script:OnPrem.SetCmdlet) {
            throw 'This on-premises object has no mailbox, so there is nothing to write to.'
        }
        $p['Identity'] = $script:OnPrem.Smtp
        Invoke-OnPremCommand -Name $script:OnPrem.SetCmdlet -Parameters $p | Out-Null
    } else {
        $p['Identity'] = $script:Mbx.Smtp
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters $p | Out-Null
    }
}

# STEP MESSAGES BEFORE EVERY BLOCKING CALL -- see the Apply handlers.
#
# Every Exchange and directory call in this file runs on the UI thread, and
# none of them has a timeout. When one blocks, the window greys out and the
# operator is told nothing at all: 'it just spins and goes to not responding'.
# Set-EditorStatus already ends with Refresh(), so the message paints before
# the call starts rather than queueing behind it -- which makes the status
# line the only thing that says WHICH call is stuck. There is no wrapper for
# this; Set-EditorStatus is the wrapper.

function Invoke-EditorReload {
    # Re-read whichever object was just written, and repaint the banners.
    if ($script:Mode -eq 'OnPrem') {
        $script:OnPrem = Resolve-OnPremMailboxTarget -Identity $script:OnPrem.Smtp
    } else {
        $script:Mbx = Resolve-MailboxTarget -Smtp $script:Mbx.Smtp
    }
    Update-EditorIdentity
}

function Get-EditorSyncSuffix {
    # Appended to the status line after an on-premises write. The operator who
    # is told only "done" goes and looks in Exchange Online, sees the old
    # value, and raises a bug against this tool.
    if ($script:Mode -eq 'OnPrem') { return " $($script:OnPremSyncNotice)" }
    return ''
}

function Test-EditorCloudOnly {
    <#
        In front of every Apply that has no field gate of its own.

        Five handlers -- remove access, fix automapping, copy permissions,
        limits and tip, single item recovery -- act on things with no field in
        the origin map, so Test-EditorFieldWritable never sees them. Without
        this they stayed live in on-premises mode and wrote to Exchange Online
        while the banner said otherwise, and the audit line then named the
        wrong directory for a real change.

        Returns $true when the caller may go ahead.
    #>
    param([string]$What = 'This')
    if ($script:Mode -ne 'Cloud') {
        Set-EditorStatus -Text "$What is an Exchange Online operation. Set Editing back to Cloud." -Warn
        return $false
    }
    return $true
}

# THE GRID AND DIALOG HELPERS MOVED TO Common.ps1 ON 2026-09-18.
# Add-GridPickColumn, Get-GridPicked, Set-GridPicked, New-DialogCheck,
# New-DialogList, ConvertFrom-DialogList and Show-ResultListDialog are shared
# with the Room and Group editors now. Three copies of a tick-box column is
# three places for 'selected' to come to mean three different things.


# THE TAB LAYOUT HELPERS MOVED TO Common.ps1 ON 2026-09-18 alongside the grid
# and dialog parts, for the same reason: the Room and Group editors lay their
# tabs out the same way, and three copies of 'what a label looks like' is how
# three windows end up looking like three products.

#endregion

#region Overview tab ------------------------------------------------------------
$ovY = 8
# DISPLAY NAME AND ALIAS ARE EDITABLE, AND THE APPLY SITS ON THE HEADING ROW.
#
# Both were read-only boxes with no Apply in EITHER mode -- not a hybrid gate,
# just a write nobody built, so the only way to rename a mailbox was to leave
# this window. Set-Mailbox declares -DisplayName and -Alias, both Applicable
# to Exchange Online as well as on-premises.
#
# THEY ARE STILL DIRECTORY-MASTERED ON A SYNCED MAILBOX, and the existing
# write-target engine already knew that -- $script:MailboxFieldOrigin lists
# both as 'Directory'. So this asks Get-MailboxWriteTarget rather than
# inventing a second opinion: a synced mailbox is refused in cloud mode with
# the engine's own reason, and a cloud-only one is simply Ok.
#
# THE APPLY IS ON THE HEADING ROW because this tab is within a dozen pixels
# of the bottom of the tab area and a button row would not fit.
[void](New-TabLabel $tabOverview 'IDENTITY' 8 $ovY -Bold)
$btnIdApply = New-TabButton $tabOverview 'Apply name / alias' 900 ($ovY - 6) 150
$ovIdNote = New-TabLabel $tabOverview '' 170 $ovY 720 -Muted
# 26, NOT 22. The Apply button is 28 tall and starts four pixels above this
# row, so the default 22 left it four pixels inside the Display name row.
# Four pixels is what this costs; the tab has them.
$ovY += 26
[void](New-TabLabel $tabOverview 'Display name:' 8 $ovY)
# New-TabValue, NOT A HAND-ROLLED TextBox. It returns a read-only TextBox of
# the height every other value on this tab uses; rolling our own at 22px put
# it one pixel into the row below and the layout checker caught it.
# Set-EditorIdentityEditable flips ReadOnly.
$ovName = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Mailbox type:' 560 $ovY 135)
$ovType = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
[void](New-TabLabel $tabOverview 'Primary address:' 8 $ovY)
$ovSmtp = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Alias:' 560 $ovY 135)
$ovAlias = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
[void](New-TabLabel $tabOverview 'User ID (UPN):' 8 $ovY)
$ovUpn = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Object GUID:' 560 $ovY 135)
$ovGuid = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
# 'Hidden from GAL' IS NOT SHOWN HERE ANY MORE. The editable row below the
# identity block is the one place it appears; a read-only copy of the same fact
# eight rows above it was the same fact twice on one tab.
[void](New-TabLabel $tabOverview 'Source:' 8 $ovY)
$ovSource = New-TabValue $tabOverview 170 $ovY 360
$ovY += 24

# LICENCE IS ON THE FIRST TAB EVEN THOUGH IT IS READ-ONLY HERE. It decides
# whether mail keeps arriving, and it is the first thing asked about a mailbox
# that has stopped working -- so it belongs where you land, not three tabs in.
# Group-based, so it reports the GROUP that drives the assignment rather than
# the SKU, because the group is the thing anyone can act on.
[void](New-TabLabel $tabOverview 'Licensed:' 8 $ovY)
$ovLicence = New-TabValue $tabOverview 170 $ovY 360
$ovLicNote = New-TabLabel $tabOverview '' 560 $ovY 460 -Muted
$ovLicNote.Text = '-- licence GROUP membership; change it in the directory.'
$ovY += 28

# HIDE-FROM-GAL LIVES ON THE FIRST TAB IN ALL THREE EDITORS.
#
# It was on Settings here and on Overview in the Room Editor, which is exactly
# the kind of drift that makes three windows feel like three products. It is a
# first-tab question -- 'can people find this' belongs beside 'what is this'
# -- so the Room Editor's placement won and this moved to match.
#
# The Settings tab was then removed entirely, so this is the only place it
# appears -- see the note where that tab used to be.
[void](New-TabLabel $tabOverview 'Hide from the address book:' 8 ($ovY + 3) 180)
$cmbOvHide               = New-Object System.Windows.Forms.ComboBox
$cmbOvHide.DropDownStyle = 'DropDownList'
$cmbOvHide.Location      = New-Object System.Drawing.Point(196, $ovY)
$cmbOvHide.Size          = New-Object System.Drawing.Size(110, 24)
[void]$cmbOvHide.Items.Add('True'); [void]$cmbOvHide.Items.Add('False')
$tabOverview.Controls.Add($cmbOvHide)
$btnOvHide = New-TabButton $tabOverview 'Apply' 320 ($ovY - 2) 100
$ovHideNote = New-TabLabel $tabOverview '' 430 ($ovY + 3) 600 -Muted
$ovY += 30

# MAILBOX TYPE AND ARCHIVE JOIN IT, for the same reason hide did: these are
# the three things somebody changes about a mailbox without needing to know
# anything else about it, and making people find a Settings tab for them was
# the friction this window exists to remove.
[void](New-TabLabel $tabOverview 'Mailbox type:' 8 ($ovY + 3) 180)
$cmbOvType               = New-Object System.Windows.Forms.ComboBox
$cmbOvType.DropDownStyle = 'DropDownList'
$cmbOvType.Location      = New-Object System.Drawing.Point(196, $ovY)
$cmbOvType.Size          = New-Object System.Drawing.Size(140, 24)
foreach ($t in @('Regular', 'Shared', 'Room', 'Equipment')) { [void]$cmbOvType.Items.Add($t) }
$tabOverview.Controls.Add($cmbOvType)
$btnOvType = New-TabButton $tabOverview 'Convert' 350 ($ovY - 2) 100
$ovTypeNote = New-TabLabel $tabOverview '' 460 ($ovY + 3) 580 -Muted
$ovY += 30

# NO EXPLANATION ON THE ARCHIVE BUTTON. There is nothing to decide: either the
# mailbox has an archive or it does not, and if it does the button is greyed
# and says so. A paragraph about what an archive is would be read once and
# then be in the way for ever.
[void](New-TabLabel $tabOverview 'Online archive:' 8 ($ovY + 3) 180)
$btnOvArchive = New-TabButton $tabOverview 'Enable archive' 196 ($ovY - 2) 140
# AUTO-EXPANDING IS A SECOND, SEPARATE, ONE-WAY SWITCH.
#
# Enable-Mailbox -AutoExpandingArchive, cloud only. The reference states both
# of the things that shape this control, so neither is a guess:
#
#   'The user's archive mailbox has to be enabled before auto-expanding
#    archiving can be enabled.'   -- so it is offered only when there IS an
#                                    archive, never as a way to create one.
#   'After you enable auto-expanding archiving for the user's mailbox, it
#    can't be disabled.'          -- so there is an Enable and no Disable, and
#                                    the confirmation says so out loud.
#
# THE NOTE GAVE UP 190 PIXELS to make room. Nothing else on this tab had any
# to spare.
$ovArchNote = New-TabLabel $tabOverview '' 350 ($ovY + 3) 500 -Muted
$btnOvArchAuto = New-TabButton $tabOverview 'Enable auto-expanding' 862 ($ovY - 2) 175
$ovY += 32

[void](New-TabLabel $tabOverview 'DIRECTORY' 8 $ovY -Bold)
$ovDirNote = New-TabLabel $tabOverview '' 175 $ovY 860 -Muted
$ovY += 22
[void](New-TabLabel $tabOverview 'Title:' 8 $ovY)
$ovTitle = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Department:' 560 $ovY 135)
$ovDept = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
[void](New-TabLabel $tabOverview 'Manager:' 8 $ovY)
$ovMgr = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Office:' 560 $ovY 135)
$ovOffice = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24

# TWO FREE-TEXT FIELDS THAT ARE NOT THE SAME FIELD.
#
#   Description   the one-line box at the top of ADUC's General tab, LDAP
#                 'description'. The field most AD reports read. There is NO
#                 Exchange path to it on-premises -- Set-Mailbox, Set-User and
#                 Set-Place declare no Description parameter at all -- so it
#                 goes straight at the directory attribute. ACTIVE DIRECTORY
#                 ONLY, and HIDDEN rather than greyed when there is no
#                 directory object behind the mailbox.
#   Notes         the big box on ADUC's Telephones tab. Written with
#                 Set-User -Notes, which is Applicable on-premises AND in
#                 Exchange Online, so this one survives a cloud-only tenant.
#
# Merging them into one control would silently overwrite whichever the
# operator did not mean, so they stay apart and are labelled as they are in
# the tool people compare against.
$lblOvDescCap = New-TabLabel $tabOverview 'Description:' 8 $ovY 150
$ovDesc = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Notes:' 560 $ovY 135)
$ovUserNotes = New-TabValue $tabOverview 700 $ovY 320
$ovY += 28

# HIDDEN IN CLOUD MODE. Title, department, office and manager are directory
# properties; in the cloud they are read-only here on purpose, because writing
# them to a synced object is the silent-revert trap this whole window is built
# around. In on-premises mode the same four boxes become editable and this
# appears under them.
$btnDirApply = New-TabButton $tabOverview 'Apply Directory Fields' 170 $ovY 170
$btnDirApply.Visible = $false
$ovY += 32

[void](New-TabLabel $tabOverview 'SIZE AND ACTIVITY' 8 $ovY -Bold)
$ovY += 22
[void](New-TabLabel $tabOverview 'Mailbox size:' 8 $ovY)
$ovSize = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Items:' 560 $ovY 135)
$ovItems = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
[void](New-TabLabel $tabOverview 'Quota:' 8 $ovY)
$ovQuota = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Archive:' 560 $ovY 135)
$ovArchive = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
# THE ARCHIVE'S OWN SIZE, on its own line. The row above says whether an
# archive EXISTS, which answers a different question from 'is it filling up'.
# Asked for 2026-09-19: the mailbox had a size and a quota and the archive had
# neither, so the one place people look to decide whether to enable
# auto-expanding could not tell them how close it was to needing it.
[void](New-TabLabel $tabOverview 'Archive size:' 8 $ovY)
$ovArchSize = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Auto-expanding:' 560 $ovY 135)
$ovArchAuto = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
# REFRESH SITS ON THE LOGON ROW rather than on a row of its own. That row was
# bought to pay for the Description and Notes fields above: this tab was
# within six pixels of the bottom of the tab area before they went in, and a
# button with a whole row to itself was the cheapest thing to give up.
[void](New-TabLabel $tabOverview 'Last logon:' 8 $ovY)
$ovLogon = New-TabValue $tabOverview 170 $ovY 360
$ovLogonNote = New-TabLabel $tabOverview '' 560 $ovY 350 -Muted
$btnOvRefresh = New-TabButton $tabOverview 'Refresh size' 924 ($ovY - 2) 130
$ovY += 34

# THE SAME BOX EVERY EDITOR HAS, in the same place. This used to be a notes
# box; it now leads with findings, and the notes follow. The Health tab is
# still the full check -- this is the glance, from facts the tab already read.
$ovNotes = Add-EditorVerdictBox -Tab $tabOverview -Y $ovY -Height 74

function Get-EditorDirectoryFields {
    <#
        Title, department, manager, office.

        THESE ARE NOT EXCHANGE FIELDS. For a cloud-only object they come from
        Entra; for a synced one they are mastered in Active Directory. Read
        from the recipient's user object where the toolkit already has a way
        to do it, and reported as blank rather than guessed at when it does
        not -- a wrong title in a ticket is worse than an absent one.
    #>
    param([Parameter(Mandatory)]$Mbx)
    $r = [PSCustomObject]@{ Title = ''; Department = ''; Manager = ''; Office = ''; Notes = ''; Source = ''; Note = '' }
    # Get-User is the Exchange view of the directory object and works for both
    # cloud and synced mailboxes without needing AD or Graph to be connected.
    try {
        $u = Invoke-EXOCommand -Name 'Get-User' -Parameters @{ Identity = $Mbx.Smtp; ErrorAction = 'Stop' }
        $r.Title      = "$($u.Title)"
        $r.Department = "$($u.Department)"
        $r.Office     = "$($u.Office)"
        $r.Manager    = "$($u.Manager)"
        $r.Notes      = "$($u.Notes)"
        $r.Source     = 'Get-User (Exchange Online)'
    } catch {
        # EXPECTED, NOT A FAULT, on a mailbox with no user object behind it.
        # A shared, room or equipment mailbox has no title or manager to read,
        # and reporting that as a problem trains people to ignore this box.
        if ("$($Mbx.Type)" -match '(?i)Shared|Room|Equipment') {
            $r.Source = 'no user object -- a ' + $Mbx.Type + ' has no directory fields'
        } else {
            $r.Note = "Directory fields could not be read: $($_.Exception.Message)"
        }
    }
    return $r
}

function Update-OverviewDescription {
    <#
        Show, hide and fill the Description box.

        HIDDEN RATHER THAN GREYED when there is nothing behind it. A cloud-only
        tenant has no Active Directory at all, and a greyed box invites the
        question 'why can't I edit this' about a field that does not exist.
        The test is whether the distinguished name belongs to a CONFIGURED
        directory, which also correctly hides it for a cloud-native mailbox
        and in cloud mode, where the only DN available is Exchange Online's.

        MORE THAN ONE VALUE IS SHOWN READ-ONLY. description is multi-valued in
        the schema, and writing a single value back would discard the rest --
        the same line the Attributes editor takes.
    #>
    param([string]$Dn)
    $dir = $null
    if ("$Dn".Trim()) { try { $dir = Get-DirectoryFromDn -Dn "$Dn" } catch { $dir = $null } }
    $show = [bool]$dir
    foreach ($c in @($lblOvDescCap, $ovDesc)) { $c.Visible = $show }
    $script:OvDescDn = $(if ($show) { "$Dn" } else { '' })
    $script:OvDescMulti = $false
    if (-not $show) { $ovDesc.Text = ''; return }

    $d = Get-AdDescription -Dn "$Dn"
    $ovDesc.Text = $d.Text
    $script:OvDescMulti = ($d.Count -gt 1)
    if (-not $d.Ok) {
        $ovDesc.Text = '(could not be read)'
        $script:OvDescDn = ''
    }
}

function Set-EditorIdentityEditable {
    # Display name and alias share one Apply and one rule, so they share one
    # place that decides whether they are writable and says why not.
    param([bool]$On, [string]$Reason = '')
    foreach ($t in @($ovName, $ovAlias)) {
        $t.ReadOnly    = -not $On
        $t.BorderStyle = $(if ($On) { 'FixedSingle' } else { 'None' })
        $t.BackColor   = $(if ($On) { [System.Drawing.Color]::White } else { [System.Drawing.Color]::WhiteSmoke })
    }
    $btnIdApply.Enabled = $On
    $ovIdNote.Text      = $Reason
    # NOT RED: the reason is always "this field is mastered elsewhere", which
    # is a routing fact, not a failure.
    $ovIdNote.ForeColor = $(if ($On) { (Get-ToolStatusColour -Kind Muted) } else { (Get-ToolStatusColour -Kind Warn) })
}

function Set-EditorDirectoryEditable {
    # The directory boxes are the same controls in both modes; only their
    # ReadOnly flag and the Apply button move. Two sets of controls would mean
    # two places for the values to disagree.
    #
    # NOTES TRAVELS WITH THEM. Set-User -Notes is Applicable on-premises and in
    # Exchange Online both, so it is writable wherever the other four are.
    # DESCRIPTION DOES NOT -- it is an Active Directory attribute with its own
    # visibility rule, handled where it is populated.
    param([bool]$On, [bool]$Cloud = $false)
    foreach ($t in @($ovTitle, $ovDept, $ovMgr, $ovOffice, $ovUserNotes)) {
        $t.ReadOnly    = -not $On
        $t.BorderStyle = $(if ($On) { 'FixedSingle' } else { 'None' })
        $t.BackColor   = $(if ($On) { [System.Drawing.Color]::White } else { [System.Drawing.Color]::WhiteSmoke })
    }
    $btnDirApply.Visible = $On
    $script:OvDirCloud = $Cloud
    # THE CAPTION IS FIXED. It briefly relabelled itself 'Apply directory
    # fields (cloud)' to show where the write was going; the caption was too
    # long for the button and the information was already on the note beside
    # it, which says in words whether this is Exchange Online or on-premises.
    # Two places saying the same thing, one of them badly, is not redundancy.
}

function Update-OverviewOnPrem {
    <#
        The same tab, read from the on-premises object.

        SIZE AND ACTIVITY IS DELIBERATELY BLANKED rather than left showing the
        cloud numbers. A mailbox size beside an on-premises banner reads as the
        on-premises mailbox's size, and for a migrated mailbox there is no such
        thing -- the stub holds no mail.
    #>
    $o = $script:OnPrem
    if (-not $o -or -not $o.Ok) {
        $ovNotes.Text = $(if ($o) { $o.Reason } else { 'No on-premises object loaded.' })
        return
    }
    $u = $o.User
    $ovName.Text   = "$($o.Display)"
    $ovType.Text   = "$($o.Kind)$(if ($o.Raw.RecipientTypeDetails) { " ($($o.Raw.RecipientTypeDetails))" })"
    $ovSmtp.Text   = "$($o.Smtp)"
    $ovAlias.Text  = "$($o.Alias)"
    # ON-PREMISES MODE WRITES ON-PREMISES, so both are editable whenever there
    # is an object to write to, using whichever Set-* cmdlet resolved for it.
    $script:IdWas = [PSCustomObject]@{ DisplayName = "$($o.Display)"; Alias = "$($o.Alias)"; Cloud = $false }
    Set-EditorIdentityEditable -On ([bool]$o.SetCmdlet) -Reason $(if ($o.SetCmdlet) { "Written on-premises with $($o.SetCmdlet)." } else { 'No on-premises object to write to.' })
    $ovUpn.Text    = "$(if ($u) { $u.UserPrincipalName } else { '' })"
    $ovGuid.Text   = "$($o.Raw.Guid)"
    $ovSource.Text = 'Active Directory (on-premises)'
    Update-OverviewHide -Hidden ([bool]$o.Hidden)

    $ovTitle.Text  = "$(if ($u) { $u.Title } else { '' })"
    $ovDept.Text   = "$(if ($u) { $u.Department } else { '' })"
    $ovMgr.Text    = "$(if ($u) { $u.Manager } else { '' })"
    $ovOffice.Text = "$(if ($u) { $u.Office } else { '' })"
    $ovUserNotes.Text = "$(if ($u) { $u.Notes } else { '' })"
    # THE BASELINE IS WHATEVER FILLED THE BOXES, not whatever the Apply
    # handler can reach later. Comparing a typed value against a re-read from
    # somewhere else is how an edit gets discarded as 'unchanged'.
    $script:OvDirWas = [PSCustomObject]@{
        Title = "$(if ($u) { $u.Title } else { '' })"; Department = "$(if ($u) { $u.Department } else { '' })"
        Office = "$(if ($u) { $u.Office } else { '' })"; Manager = "$(if ($u) { $u.Manager } else { '' })"
        Notes = "$(if ($u) { $u.Notes } else { '' })"; Identity = "$(if ($u) { $u.Identity } else { '' })"
    }
    Update-OverviewDescription -Dn "$($o.Raw.DistinguishedName)"
    $ovDirNote.Text = '-- editable here; these are the Active Directory values.'
    $ovDirNote.ForeColor = (Get-ToolStatusColour -Kind Muted)

    $ovLicence.Text = '(switch to Cloud to read the licence)'
    foreach ($t in @($ovSize, $ovItems, $ovQuota, $ovLogon)) { $t.Text = '(Exchange Online only)' }
    $ovArchive.Text = '(Exchange Online only)'
    $ovArchSize.Text = '(Exchange Online only)'
    $ovArchAuto.Text = '(Exchange Online only)'
    $ovLogonNote.Text = 'Size, item count and last logon live on the cloud mailbox. Switch Editing to Cloud to see them.'

    $ovNotes.Text = "$($o.Reason)`r`n$($script:OnPremSyncNotice)"
    Set-EditorDirectoryEditable -On ($null -ne $u)
    if (-not $u) {
        $ovDirNote.Text = '-- no on-premises user object was returned, so the directory fields cannot be edited.'
        $ovDirNote.ForeColor = (Get-ToolStatusColour -Kind Bad)
    }
}

function Update-OverviewTypeArchive {
    <#
        The type picker and the archive button.

        THE ARCHIVE IS CLOUD-ONLY and is greyed in on-premises mode.

        THE TYPE PICKER IS NOT, and this is the one place in the window where
        the Cloud / On-premises switch does not decide anything. Mailbox type
        has to be set on BOTH sides in a hybrid -- see the long note on the
        Convert button -- so the button does both whichever mode is selected,
        and greying it in on-premises mode would only hide the half of the job
        that mode is for.
    #>
    if ($script:Mode -eq 'OnPrem') {
        $btnOvArchive.Enabled = $false
        $btnOvArchAuto.Enabled = $false
        $ovArchNote.Text = 'The archive is an Exchange Online feature. Set Editing back to Cloud.'
        $ovArchNote.ForeColor = (Get-ToolStatusColour -Kind Warn)
        # AND THEN FALL THROUGH, so the picker is filled and enabled below.
    }

    $mb = $script:Mbx.Raw
    $from = "$($mb.RecipientTypeDetails)"
    $show = @{ UserMailbox = 'Regular'; SharedMailbox = 'Shared'
               RoomMailbox = 'Room'; EquipmentMailbox = 'Equipment' }[$from]
    if ($show) {
        $i = $cmbOvType.Items.IndexOf($show)
        if ($i -ge 0) { $cmbOvType.SelectedIndex = $i }
    } else {
        # A type this picker does not carry -- SchedulingMailbox, TeamMailbox
        # and the like. Left alone, the box keeps the PREVIOUS mailbox's
        # selection while the note beside it says something else.
        $cmbOvType.SelectedIndex = -1
    }
    $cmbOvType.Enabled = $true; $btnOvType.Enabled = $true
    # KEPT SHORT ON PURPOSE. The label is 580px; the earlier wording ran past
    # it and was silently clipped. The reason lives in the confirm dialog,
    # where there is room for it and where it is actually needed.
    $ovTypeNote.Text = "Currently $from.  Convert writes to both directories."
    $ovTypeNote.ForeColor = (Get-ToolStatusColour -Kind Muted)

    # THE SAME GUID TEST AS THE SETTINGS TAB HAD. A mailbox with no archive
    # returns the all-zero GUID, which compares correctly -- but one where the
    # property is ABSENT returns $null, and '$null -ne [guid]::Empty' is TRUE,
    # which reported an archive on every mailbox that did not have one.
    $arcGuid = "$($mb.ArchiveGuid)"
    $hasArchive = ("$($mb.ArchiveStatus)" -eq 'Active') -or
                  ($arcGuid -and $arcGuid -ne '00000000-0000-0000-0000-000000000000')
    # THE MODE TEST IS REPEATED HERE ON PURPOSE. The branch above no longer
    # returns early, so without this line the archive button would be switched
    # back on a few lines after being deliberately switched off.
    if ($script:Mode -ne 'OnPrem') {
        $btnOvArchive.Enabled = -not $hasArchive
        $ovArchNote.Text = $(if ($hasArchive) { 'Archive already enabled.' } else { 'No archive. Enabling one is not reversible from this window.' })
        $ovArchNote.ForeColor = (Get-ToolStatusColour -Kind Muted)
    }
}

$btnOvType.Add_Click({
    <#
        THE ONE BUTTON THAT IGNORES THE CLOUD / ON-PREMISES SWITCH.

        Mailbox type is NOT a synced attribute in the direction people assume.
        msExchRecipientTypeDetails flows on-premises -> cloud. It never flows
        back. So converting in Exchange Online alone leaves the on-premises
        remote mailbox still saying RemoteUserMailbox, and the next directory
        sync can push that stale value up and UNDO the conversion.

        Microsoft says this outright in the Set-RemoteMailbox reference, under
        -Type: "In hybrid environments, changing the mailbox type of a
        migrated mailbox needs to be done on both sides: Set-Mailbox in
        Exchange Online and Set-RemoteMailbox in on-premises Exchange." The
        same page names the symptom when you do not -- directory sync
        converting shared mailboxes in Exchange Online back into user
        mailboxes.

        So this button does both, whichever mode the window is in. Every other
        button on every tab still obeys the switch; this one cannot, because
        a half-converted mailbox is a worse outcome than either whole one.

        ON-PREMISES GOES FIRST, DELIBERATELY. Sync carries on-premises upward,
        so setting it there first means a sync cycle landing mid-operation
        pushes the value we WANT. Cloud-first leaves a window in which sync
        reverts the change we just made.
    #>
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $to   = "$($cmbOvType.SelectedItem)"
        if (-not $to) { Set-EditorStatus -Text 'Pick a mailbox type first.' -Bad; return }
        $from = "$($script:Mbx.Raw.RecipientTypeDetails)"
        $typeIs = @{ Regular = 'UserMailbox'; Shared = 'SharedMailbox'
                     Room = 'RoomMailbox'; Equipment = 'EquipmentMailbox' }
        if ($from -ieq "$($typeIs[$to])") { Set-EditorStatus -Text "Already a $to mailbox."; return }

        # ---- work out what the on-premises side is, before asking anything --
        $ready   = Test-OnPremMailboxReady
        $onPrem  = $null
        $opKind  = 'none'
        $opWhy   = ''
        if ($ready.Ok) {
            try { $onPrem = Resolve-OnPremMailboxTarget -Identity $script:Mbx.Smtp } catch { $onPrem = $null }
            if     (-not $onPrem -or -not $onPrem.Ok) { $opKind = 'missing'; $opWhy = 'No on-premises object was found for this address.' }
            elseif ($onPrem.Kind -eq 'RemoteMailbox') { $opKind = 'remote' }
            elseif ($onPrem.Kind -eq 'Mailbox')       { $opKind = 'local' }
            else                                      { $opKind = 'user'; $opWhy = 'The on-premises object is a user with no mailbox.' }
        } else {
            $opKind = 'unavailable'; $opWhy = $ready.Reason
        }

        if ($opKind -eq 'user') {
            Show-ToolDialog($opWhy + "`r`n`r`nThere is nothing on-premises to convert, and converting only in the cloud would be undone by the next directory sync.",
                            'Cannot convert', 'OK', 'Warning') | Out-Null
            return
        }

        # ---- one dialog, saying exactly what will be written and where -----
        $plan = switch ($opKind) {
            'remote'      { "BOTH SIDES will be changed:`r`n" +
                            "   on-premises   Set-RemoteMailbox -Type $to`r`n" +
                            "   Exchange Online   Set-Mailbox -Type $to" }
            'local'       { "This mailbox has not been migrated -- it still lives on-premises, so there is nothing in the cloud to convert.`r`n" +
                            "   on-premises   Set-Mailbox -Type $to" }
            'missing'     { "ONLY THE CLOUD WILL BE CHANGED. $opWhy`r`n`r`n" +
                            "If this address IS synced from on-premises and the on-premises object simply could not be read, the next directory sync will UNDO this conversion." }
            'unavailable' { "ONLY THE CLOUD WILL BE CHANGED. $opWhy`r`n`r`n" +
                            "Mailbox type does not sync cloud-to-on-premises. If this mailbox is synced, the next directory sync will UNDO this conversion. Connect on-premises Exchange and do it there as well." }
            default       { "Exchange Online only." }
        }
        $extra = ''
        if ($to -eq 'Shared') {
            $extra = "`r`n`r`nAFTER THIS: recurring meetings this person organises can no longer be transferred -- the transfer cmdlet refuses a shared mailbox. Do that first if it is needed."
        }
        if ($from -imatch 'Shared' -and $to -eq 'Regular') {
            $extra = "`r`n`r`nA regular mailbox needs a licence. Assign one, or mail will stop after the grace period."
        }
        $a = Show-ToolDialog(("Convert $($script:Mbx.Display) from $from to $to?`r`n`r`n" + $plan + $extra),
                             'Convert mailbox type', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        # ---- write, on-premises first ---------------------------------------
        $did = @()
        if ($opKind -eq 'remote' -or $opKind -eq 'local') {
            $cmd = $(if ($opKind -eq 'remote') { 'Set-RemoteMailbox' } else { 'Set-Mailbox' })
            try {
                Invoke-OnPremCommand -Name $cmd -Parameters @{
                    Identity = $onPrem.Smtp; Type = $to; ErrorAction = 'Stop' } | Out-Null
                $did += "on-premises ($cmd)"
                Write-EditorAudit -Action 'MailboxTypeConvertOnPrem' -Environment 'OnPrem' -Detail "$cmd -Type $to (was $from)."
            } catch {
                # NOTHING IS ATTEMPTED IN THE CLOUD AFTER THIS FAILS. A cloud-only
                # conversion is the exact state this whole button exists to avoid.
                Write-EditorAudit -Action 'MailboxTypeConvertOnPrem' -Result 'Failed' -Environment 'OnPrem' -Detail "$cmd -Type $to failed: $($_.Exception.Message)"
                Show-ToolDialog(("The on-premises change failed, so nothing was changed in the cloud either:`r`n`r`n$($_.Exception.Message)`r`n`r`n" +
                                 "The mailbox is still $from on both sides."),
                                'Convert failed', 'OK', 'Error') | Out-Null
                Set-EditorStatus -Text 'Convert failed on-premises; nothing was changed.' -Bad
                return
            }
        }

        if ($opKind -ne 'local') {
            try {
                Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
                    Identity = $script:Mbx.Smtp; Type = $to; ErrorAction = 'Stop' } | Out-Null
                $did += 'Exchange Online (Set-Mailbox)'
                Write-EditorAudit -Action 'MailboxTypeConvertCloud' -Environment 'Cloud' -Detail "Set-Mailbox -Type $to (was $from)."
            } catch {
                # THE HALF-CONVERTED STATE IS NAMED, NOT SWALLOWED. On-premises
                # is already changed; saying so is the difference between a
                # retry and a mystery.
                Write-EditorAudit -Action 'MailboxTypeConvertCloud' -Result 'Failed' -Environment 'Cloud' -Detail "Set-Mailbox -Type $to failed: $($_.Exception.Message)"
                $done = $(if ($did) { "On-premises IS ALREADY CONVERTED to $to." } else { 'Nothing was changed.' })
                Show-ToolDialog(("The cloud change failed:`r`n`r`n$($_.Exception.Message)`r`n`r`n$done`r`n`r`n" +
                                 "Run Set-Mailbox -Identity $($script:Mbx.Smtp) -Type $to in Exchange Online, or press Convert again once the cause is fixed."),
                                'Convert partly done', 'OK', 'Error') | Out-Null
                Set-EditorStatus -Text 'Converted on-premises; the cloud change failed.' -Bad
                $script:Mbx = Resolve-MailboxTarget -Smtp $script:Mbx.Smtp
                Update-EditorIdentity; Update-OverviewTab
                return
            }
        }

        $script:Mbx = Resolve-MailboxTarget -Smtp $script:Mbx.Smtp
        Update-EditorIdentity
        Update-OverviewTab
        $tail = $(if ($opKind -in @('missing','unavailable')) { '  On-premises was NOT changed -- directory sync may undo this.' } else { '' })
        Set-EditorStatus -Text ("Converted to $to -- " + ($did -join ' and ') + '.' + $tail) -Bad:($opKind -in @('missing','unavailable'))
    } 'Convert mailbox type'
})
$btnOvArchAuto.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Enabling auto-expanding archiving')) { return }
        if ($script:ArchAuto -ne 'off') {
            Set-EditorStatus -Text 'Auto-expanding archiving is only offered when it is known to be off.' -Bad; return
        }
        $a = Show-ToolDialog(
            ("Turn on auto-expanding archiving for $($script:Mbx.Display)?" + "`r`n`r`n" +
             'Extra storage is added to the archive automatically as it approaches its limit.' + "`r`n`r`n" +
             'THIS CANNOT BE UNDONE. Microsoft documents auto-expanding archiving as permanent once enabled for a mailbox -- there is no way to turn it off again, here or anywhere else.'),
            'Enable auto-expanding archiving', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        # A SWITCH, NOT A BOOLEAN. It has its own parameter set on
        # Enable-Mailbox, so -Archive is not sent alongside it.
        Invoke-EXOCommand -Name 'Enable-Mailbox' -Parameters @{
            Identity = $script:Mbx.Smtp; AutoExpandingArchive = $true; Confirm = $false; ErrorAction = 'Stop'
        } | Out-Null
        Write-EditorAudit -Action 'MailboxAutoExpandingArchive' -Detail 'Auto-expanding archiving enabled (not reversible).'
        $script:Mbx = Resolve-MailboxTarget -Smtp $script:Mbx.Smtp
        Update-OverviewTab
        Set-EditorStatus -Text 'Auto-expanding archiving enabled. It can take a while before the extra space appears.'
    } 'Enable auto-expanding archiving'
})

$btnOvArchive.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Enabling the archive')) { return }
        if (-not (Test-EditorFieldWritable -Field 'ArchiveEnabled')) { return }
        $a = Show-ToolDialog(
            ("Enable the online archive for $($script:Mbx.Display)?" + "`r`n`r`n" +
             'This is not reversible from this tool. It can take a while to appear in Outlook.'),
            'Enable archive', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Invoke-EXOCommand -Name 'Enable-Mailbox' -Parameters @{
            Identity = $script:Mbx.Smtp; Archive = $true; ErrorAction = 'Stop'
        } | Out-Null
        Write-EditorAudit -Action 'MailboxArchiveEnable' -Detail 'Online archive enabled.'
        $script:Mbx = Resolve-MailboxTarget -Smtp $script:Mbx.Smtp
        Update-OverviewTab
        Set-EditorStatus -Text 'Archive enabled. It takes a while to appear in Outlook.'
    } 'Enable archive'
})

function Update-OverviewHide {
    <#
        The hide control on the first tab, and whether it can be used.

        THE GATE IS THE SAME ONE EVERY OTHER FIELD GOES THROUGH. Hide-from-GAL
        is directory-mastered, so on a synced mailbox in Cloud mode Exchange
        Online accepts the write and directory sync puts it back within about
        half an hour, with no error anywhere. Greying it here, with the reason
        beside it, is the whole point of the window.
    #>
    param([bool]$Hidden)
    $cmbOvHide.SelectedItem = $(if ($Hidden) { 'True' } else { 'False' })
    Update-OverviewTypeArchive
    if ($script:Mode -eq 'OnPrem') {
        $ok = ($script:OnPrem -and $script:OnPrem.Ok -and $script:OnPrem.SetCmdlet)
        $cmbOvHide.Enabled = $ok; $btnOvHide.Enabled = $ok
        $ovHideNote.Text = $(if ($ok) { "Written on-premises with $($script:OnPrem.SetCmdlet)." } else { 'No on-premises mailbox object to change.' })
        $ovHideNote.ForeColor = (Get-ToolStatusColour -Kind Warn)
        return
    }
    $t = Get-MailboxWriteTarget -Field 'HiddenFromAddressListsEnabled' -IsSynced $script:Mbx.IsSynced
    $ok = ($t.State -ne 'Directory')
    $cmbOvHide.Enabled = $ok; $btnOvHide.Enabled = $ok
    $ovHideNote.Text = $(if ($ok) { 'Removes it from the address book. Mail sent straight to the address still delivers.' } else { $t.Reason })
    $ovHideNote.ForeColor = $(if ($ok) { (Get-ToolStatusColour -Kind Muted) } else { (Get-ToolStatusColour -Kind Warn) })
}

$btnIdApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not $script:IdWas) { Set-EditorStatus -Text 'Nothing loaded to write to.' -Bad; return }
        if (-not (Test-EditorFieldWritable -Field 'DisplayName')) { return }

        $wantName  = "$($ovName.Text)".Trim()
        $wantAlias = "$($ovAlias.Text)".Trim()
        if (-not $wantName)  { Set-EditorStatus -Text 'Display name cannot be empty.' -Bad; return }
        # AN ALIAS IS NOT FREE TEXT. Exchange rejects spaces and most
        # punctuation, and the failure comes back as a generic parameter
        # error -- so it is caught here with a reason instead.
        if (-not $wantAlias) { Set-EditorStatus -Text 'Alias cannot be empty.' -Bad; return }
        if ($wantAlias -notmatch '^[A-Za-z0-9._-]+$') {
            Set-EditorStatus -Text 'An alias may contain only letters, digits, full stops, hyphens and underscores.' -Bad; return
        }

        # -Confirm:$false -- see the note on the directory-fields Apply.
        $p = @{ Identity = $script:Mbx.Smtp; Confirm = $false; ErrorAction = 'Stop' }
        $changed = @()
        if ($wantName -ne "$($script:IdWas.DisplayName)".Trim())   { $p['DisplayName'] = $wantName;  $changed += "DisplayName='$wantName'" }
        if ($wantAlias -ne "$($script:IdWas.Alias)".Trim())        { $p['Alias']       = $wantAlias; $changed += "Alias='$wantAlias'" }
        if ($changed.Count -eq 0) { Set-EditorStatus -Text 'Nothing changed.'; return }

        # CHANGING AN ALIAS CAN MOVE THE PRIMARY ADDRESS. On a mailbox with the
        # e-mail address policy enabled, the policy rebuilds addresses from the
        # alias -- so this is said before it happens rather than discovered
        # afterwards by whoever stops receiving mail.
        $warn = $(if ($p.ContainsKey('Alias')) {
            "`r`n`r`nIf an e-mail address policy applies to this mailbox, changing the alias can rebuild its addresses and move the primary one."
        } else { '' })
        $a = Show-ToolDialog(("On $($script:Mbx.Display):" + "`r`n`r`n  " + ($changed -join "`r`n  ") + $warn),
                             'Apply name and alias', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        if ($script:IdWas.Cloud) {
            Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters $p | Out-Null
        } else {
            $p['Identity'] = $script:OnPrem.Smtp
            Invoke-OnPremCommand -Name $script:OnPrem.SetCmdlet -Parameters $p | Out-Null
        }
        Write-EditorAudit -Action 'MailboxIdentity' -Detail ($changed -join '; ')
        Invoke-EditorReload
        Update-OverviewTab
        Set-EditorStatus -Text ("Applied: $($changed -join '; ')." + $(if (-not $script:IdWas.Cloud) { (Get-EditorSyncSuffix) }))
    } 'Apply name and alias'
})

$btnOvHide.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorFieldWritable -Field 'HiddenFromAddressListsEnabled')) { return }
        $want = ("$($cmbOvHide.SelectedItem)" -eq 'True')
        $now  = $(if ($script:Mode -eq 'OnPrem') { [bool]$script:OnPrem.Hidden } else { [bool]$script:Mbx.Raw.HiddenFromAddressListsEnabled })
        if ($want -eq $now) { Set-EditorStatus -Text 'Nothing changed.'; return }
        Invoke-EditorSetMailbox -Parameters @{ HiddenFromAddressListsEnabled = $want }
        Write-EditorAudit -Action 'MailboxHideFromGal' -Detail "Hidden from address list = $want."
        Invoke-EditorReload
        Update-OverviewTab
        Set-EditorStatus -Text ($(if ($want) { 'Hidden from the address list.' } else { 'Visible in the address list.' }) + (Get-EditorSyncSuffix))
    } 'Apply address list visibility'
})

function Update-OverviewTab {
    if ($script:Mode -eq 'OnPrem') { Update-OverviewOnPrem; return }
    # EDITABLE IN THE CLOUD ONLY WHEN NOTHING ON-PREMISES OWNS THESE FIELDS.
    #
    # On a SYNCED mailbox they are mastered in Active Directory: Exchange
    # Online accepts the write, directory sync puts it back within the hour,
    # and nothing anywhere reports an error. That silent revert is the trap
    # this window was built around and the boxes stay read-only for it -- the
    # Editing switch is the route, and the note under them says so.
    #
    # On a CLOUD-ONLY mailbox -- a cloud-native one here, or every mailbox in
    # a cloud-only tenant -- there is no directory to revert anything, and
    # Set-User is Applicable to Exchange Online. Leaving those read-only made
    # the fields permanently uneditable for a whole class of tenant, which is
    # not a safety rule, just a gap.
    $m = $script:Mbx
    $mb = $m.Raw
    # THIS CALL MUST SIT BELOW $m, not above it. The line it replaced was
    # 'Set-EditorDirectoryEditable -On $false', which referenced nothing and so
    # did not care where it sat. This one reads $m.IsSynced, and on a $null $m
    # that yields $null -- '-not $null' is TRUE, so every mailbox including a
    # synced one came back editable: the exact silent-revert trap the gate
    # exists to prevent, inverted, with no error anywhere.
    Set-EditorDirectoryEditable -On (-not $m.IsSynced) -Cloud $true
    $ovName.Text   = "$($mb.DisplayName)"
    $ovType.Text   = "$($mb.RecipientTypeDetails)"
    $ovSmtp.Text   = "$($mb.PrimarySmtpAddress)"
    $ovAlias.Text  = "$($mb.Alias)"
    $script:IdWas = [PSCustomObject]@{ DisplayName = "$($mb.DisplayName)"; Alias = "$($mb.Alias)"; Cloud = $true }
    # ONE ENGINE, NOT A SECOND OPINION. Both fields are 'Directory' in
    # $script:MailboxFieldOrigin, so a synced mailbox comes back 'Directory'
    # here and a cloud-only one comes back Ok -- which is exactly the rule
    # every other directory-mastered control on this tab already follows.
    $tId = Get-MailboxWriteTarget -Field 'DisplayName' -IsSynced $m.IsSynced
    Set-EditorIdentityEditable -On ($tId.State -ne 'Directory') `
        -Reason $(if ($tId.State -ne 'Directory') { 'Written with Set-Mailbox in Exchange Online.' } else { $tId.Reason })
    $ovUpn.Text    = "$($mb.UserPrincipalName)"
    $ovGuid.Text   = "$($mb.ExternalDirectoryObjectId)"
    $ovSource.Text = $(if ($m.IsSynced) { 'Synced from Active Directory' } else { 'Cloud only' })
    Update-OverviewHide -Hidden ([bool]$mb.HiddenFromAddressListsEnabled)

    $d = Get-EditorDirectoryFields -Mbx $m
    $ovTitle.Text  = $d.Title
    $ovDept.Text   = $d.Department
    $ovMgr.Text    = $d.Manager
    $ovOffice.Text = $d.Office
    $ovUserNotes.Text = $d.Notes
    $script:OvDirWas = [PSCustomObject]@{
        Title = "$($d.Title)"; Department = "$($d.Department)"; Office = "$($d.Office)"
        Manager = "$($d.Manager)"; Notes = "$($d.Notes)"; Identity = "$($m.Smtp)"
    }
    # NO DESCRIPTION IN CLOUD MODE, and not because it does not exist. The
    # only distinguished name on hand here is Exchange Online's own, which
    # Get-ADObject could never resolve -- so the control is hidden rather than
    # shown holding a value read from nowhere. On a synced mailbox, flip
    # Editing to on-premises to see and edit it.
    Update-OverviewDescription -Dn ''
    $ovDirNote.Text = $(if ($m.IsSynced) {
        '-- mastered in Active Directory. Set Editing (top right) to On-premises to change them here; a cloud write would be reverted by the next directory sync without an error.'
    } else {
        '-- editable here; this mailbox has no on-premises object, so Exchange Online is where they live.'
    })

    $notes = New-Object System.Text.StringBuilder
    if ($d.Note) { [void]$notes.AppendLine($d.Note) }

    # Statistics are a separate, slower call. Fetched here rather than on open
    # so the identity block appears immediately on a large mailbox.
    $ovSize.Text = '(reading...)'; $ovItems.Text = ''; $ovQuota.Text = ''; $ovLogon.Text = ''
    $ovArchSize.Text = ''; $ovArchAuto.Text = ''
    $ovSize.Refresh()
    try {
        $st = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $m.Smtp; ErrorAction = 'Stop' }
        $ovSize.Text  = "$($st.TotalItemSize)"
        $ovItems.Text = "$($st.ItemCount)"
        $ovLogon.Text = "$($st.LastLogonTime)"
        # LastLogonTime IS NOT 'WHEN THE PERSON LAST USED IT'. Service
        # processes touch a mailbox and move this value, so it overstates
        # activity -- which matters because it gets used to decide whether an
        # account is dormant. Sign-in Activity answers that question properly.
        $ovLogonNote.Text = 'Service processes move this. For real user sign-ins use the Sign-in Activity tool.'
    } catch {
        $ovSize.Text = '(unavailable)'
        $script:TabReadFailed = $true
        [void]$notes.AppendLine("Mailbox statistics could not be read: $($_.Exception.Message)")
    }
    $ovQuota.Text   = "$($mb.ProhibitSendReceiveQuota)"
    # ARCHIVEGUID IS TESTED AS A STRING, not against [guid]::Empty. A mailbox
    # with no archive returns the all-zero GUID, which compares correctly --
    # but a mailbox where the property is ABSENT returns $null, and
    # '$null -ne [guid]::Empty' is TRUE, so the check reported an archive on
    # every mailbox that did not have the property at all.
    $arcGuid = "$($mb.ArchiveGuid)"
    $hasArchive = ("$($mb.ArchiveStatus)" -eq 'Active') -or
                  ($arcGuid -and $arcGuid -ne '00000000-0000-0000-0000-000000000000')
    $ovArchive.Text = $(if ($hasArchive) { "enabled ($($mb.ArchiveStatus))" } else { 'not enabled' })

    # ARCHIVE SIZE AND AUTO-EXPANDING.
    #
    # ONLY ASKED FOR WHEN THERE IS AN ARCHIVE. Get-MailboxStatistics -Archive
    # against a mailbox without one is an error, not an empty result, so the
    # guard is what keeps a perfectly normal mailbox from posting a red line
    # about a failure that is really just 'there is no archive'.
    #
    # ITS OWN TRY, like the licence line below. An archive that cannot be read
    # must cost the archive line and nothing else -- the mailbox size above it
    # was already fetched and is still true.
    #
    # THE AUTO-EXPANDING PROPERTY IS PROBED, NOT ASSUMED. Exchange Online
    # reports it on the mailbox, on-premises Exchange does not have the feature
    # at all, and the name has not been stable across builds -- so the two
    # spellings are looked for and anything else reads '(not reported)'. A
    # missing property must not be shown as 'off': 'off' is a fact about the
    # archive, '(not reported)' is a fact about what we could see.
    if (-not $hasArchive) {
        $ovArchSize.Text = 'no archive'
        $ovArchAuto.Text = ''
        # THE ARCHIVE COMES FIRST, and the reference is explicit that it must.
        # Offering auto-expanding here would produce a service error that reads
        # as a bug in this window rather than as the prerequisite it is.
        $script:ArchAuto = 'unknown'
        $btnOvArchAuto.Enabled = $false
        $ovArchNote.Text += '  Auto-expanding needs an archive first.'
    } else {
        try {
            $ast = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $m.Smtp; Archive = $true; ErrorAction = 'Stop' }
            $ovArchSize.Text = "$($ast.TotalItemSize)  --  $($ast.ItemCount) item(s)"
        } catch {
            $ovArchSize.Text = '(unavailable)'
            [void]$notes.AppendLine("Archive statistics could not be read: $($_.Exception.Message)")
        }

        $auto = $null
        foreach ($n in @('AutoExpandingArchiveEnabled', 'AutoExpandingArchiveState')) {
            if ($mb.PSObject.Properties[$n]) { $auto = $mb.$n; break }
        }
        # THREE STATES, NOT TWO. 'We could not read it' is not 'it is off',
        # and only a definite off earns an enabled button: offering the switch
        # on a guess would mean offering an irreversible change on a guess.
        #
        # A NON-BOOLEAN VALUE COUNTS AS UNKNOWN. The property name is probed
        # rather than assumed -- a build reporting AutoExpandingArchiveState
        # returns words, and which of those words means 'off' is not something
        # this file should decide from the outside.
        if ($null -eq $auto) {
            $ovArchAuto.Text = '(not reported)'
            $script:ArchAuto = 'unknown'
        } elseif ($auto -is [bool]) {
            $ovArchAuto.Text = $(if ($auto) { 'ON' } else { 'off' })
            $script:ArchAuto = $(if ($auto) { 'on' } else { 'off' })
        } else {
            $ovArchAuto.Text = "$auto"
            $script:ArchAuto = 'unknown'
        }
        $btnOvArchAuto.Enabled = ($script:ArchAuto -eq 'off')
        # A GREYED BUTTON MUST SAY WHY, ON THE LINE NEXT TO IT.
        #
        # This window greys the auto-expanding button in three different
        # situations and the only clue was a word in a field on the row below.
        # 'I don't see the ability to enable auto-expanding archive' is what
        # that costs: the control and its reason were in different places, so
        # the absence read as the feature missing rather than as a state.
        if ($script:ArchAuto -eq 'on') {
            # THE FIELD JUST SAYS 'ON'. It briefly read 'ON (cannot be turned
            # off)', which repeated what the note on the archive row above
            # already says in a full sentence. The value column states the
            # value; the note column explains it. One fact, one place.
            $ovArchNote.Text += '  Auto-expanding is already on and cannot be turned off.'
        } elseif ($script:ArchAuto -eq 'unknown') {
            $ovArchNote.Text += '  Auto-expanding state could not be read, so it is not offered.'
            [void]$notes.AppendLine('Auto-expanding archiving could not be read, so it is not offered -- it is a one-way change and is not made on a guess.')
        } else {
            $ovArchNote.Text += '  Auto-expanding is off -- the button turns it on, permanently.'
        }

        # A QUOTA WORTH SAYING OUT LOUD. The archive quota is the number that
        # decides whether auto-expanding is needed, and it is not on any other
        # line of this tab.
        if ("$($mb.ArchiveQuota)".Trim()) { $ovArchSize.Text += "   (quota $($mb.ArchiveQuota))" }
    }

    # ITS OWN TRY, AND LAST. This is the only line on the tab that asks a
    # DIRECTORY rather than Exchange Online, so it is the only one that can
    # fail for a reason that has nothing to do with the mailbox. An unreachable
    # domain controller costs the licence line and nothing else.
    $ovLicence.Text = '(reading...)'
    try {
        $script:Licence = Get-UserLicenseState -Smtp $m.Smtp
        $ovLicence.Text = (Format-UserLicenseState -State $script:Licence)
    } catch {
        $script:Licence = $null
        $ovLicence.Text = "could not be read -- $($_.Exception.Message)"
    }

    # FINDINGS FROM WHAT THIS TAB ALREADY READ. No extra remote call: every
    # fact here is on the mailbox object or the licence state that the rows
    # above just displayed. The Health tab does the expensive checks --
    # protocols, hybrid GUIDs, inbox rules -- and the last note says so.
    $f = New-Object System.Collections.ArrayList
    $rtd = "$($mb.RecipientTypeDetails)"
    if ($mb.PSObject.Properties['AccountDisabled'] -and $mb.AccountDisabled -and $rtd -eq 'UserMailbox') {
        [void]$f.Add('The account is DISABLED but this is still a user mailbox -- a leaver nobody converted to shared, still consuming a licence.')
    }
    if ($script:Licence -and $script:Licence.PSObject.Properties['Licensed']) {
        if ($rtd -eq 'SharedMailbox' -and $script:Licence.Licensed) { [void]$f.Add('Shared mailbox IS licensed -- a shared mailbox under 50GB does not need one.') }
        if ($rtd -eq 'UserMailbox' -and -not $script:Licence.Licensed) { [void]$f.Add('User mailbox is NOT licensed. Mail stops after the grace period.') }
    }
    if ("$($mb.ForwardingSmtpAddress)" -or "$($mb.ForwardingAddress)") { [void]$f.Add('This mailbox forwards its mail. See Mail Flow for where to, and whether a copy is kept.') }
    if ($mb.LitigationHoldEnabled) { [void]$f.Add('Litigation hold is on -- nothing in this mailbox can be permanently deleted. See Retention.') }
    if ($mb.HiddenFromAddressListsEnabled) { [void]$f.Add('Hidden from the address list. Mail sent directly still delivers.') }
    $extra = @()
    if ($notes.Length -gt 0) { $extra += $notes.ToString().TrimEnd() }
    $extra += 'Protocols, hybrid GUIDs and inbox rules are checked on the Health tab.'
    Set-EditorVerdict -Findings @($f) -Notes $extra -Clean 'Nothing obviously wrong from this tab. Licensed as expected for its type, no forwarding, no hold, visible in the address book.' -Box $ovNotes
}

$btnDirApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        # WHICHEVER SIDE THE BOXES WERE FILLED FROM IS THE SIDE THIS WRITES
        # TO. Set-EditorDirectoryEditable decided that when it decided whether
        # the boxes were editable at all, so the two can never disagree.
        $u = $script:OvDirWas
        if (-not $u) { Set-EditorStatus -Text 'There is nothing loaded to write to.' -Bad; return }
        if (-not $script:OvDirCloud) {
            if ($script:Mode -ne 'OnPrem') { Set-EditorStatus -Text 'Directory fields are only editable in on-premises mode.' -Warn; return }
            if (-not (Test-EditorFieldWritable -Field 'DisplayName')) { return }
            if (-not $script:OnPrem -or -not $script:OnPrem.User) { Set-EditorStatus -Text 'There is no on-premises user object to write to.' -Bad; return }
        }

        # ONLY WHAT CHANGED. Set-User with all four every time overwrites a
        # field somebody else edited between this window loading and the
        # button being pressed, and claims four changes in the audit line for
        # the one the operator made.
        # -Confirm:$false IS NOT OPTIONAL IN A FORM.
        #
        # This is what hung. An Exchange write that supports ShouldProcess can
        # raise a confirmation prompt, and a WinForms process has NO CONSOLE
        # to render one on -- so the cmdlet sits there waiting for an answer
        # that can never be typed. The window greys out and goes to 'not
        # responding' with the write neither done nor failed.
        #
        # Every other write in this file already passed it. These new handlers
        # did not, which is why reads were fine and only Apply hung.
        $p = @{ Identity = "$($u.Identity)"; Confirm = $false; ErrorAction = 'Stop' }
        $changed = @()
        foreach ($f in @(
            @{ Name = 'Title';      Box = $ovTitle;      Was = "$($u.Title)" }
            @{ Name = 'Department'; Box = $ovDept;       Was = "$($u.Department)" }
            @{ Name = 'Office';     Box = $ovOffice;     Was = "$($u.Office)" }
            @{ Name = 'Manager';    Box = $ovMgr;        Was = "$($u.Manager)" }
            @{ Name = 'Notes';      Box = $ovUserNotes;  Was = "$($u.Notes)" }
        )) {
            $now = "$($f.Box.Text)".Trim()
            if ($now -ne "$($f.Was)".Trim()) {
                $p[$f.Name] = $(if ($now) { $now } else { $null })
                $changed += "$($f.Name)='$now'"
            }
        }
        # DESCRIPTION IS A SEPARATE WRITE TO A SEPARATE PLACE. The other five
        # are Exchange properties on the user object; description is an Active
        # Directory attribute with no Exchange parameter behind it on-premises,
        # so it goes through Set-ADObject. Doing it FIRST means a directory
        # failure stops the whole apply rather than leaving half of it done
        # with the audit line already claiming both.
        $descChanged = ''
        if ($script:OvDescDn -and -not $script:OvDescMulti) {
            $wantDesc = "$($ovDesc.Text)".Trim()
            Set-EditorStatus -Text 'Reading the current description from Active Directory...'
            $curDesc  = Get-AdDescription -Dn $script:OvDescDn
            if ($curDesc.Ok -and $wantDesc -ne "$($curDesc.Text)".Trim()) {
                Set-EditorStatus -Text 'Writing the description to Active Directory...'
                Set-AdDescription -Dn $script:OvDescDn -Value $wantDesc
                $descChanged = "Description='$(if ($wantDesc) { $wantDesc } else { '(cleared)' })'"
                $changed += $descChanged
            }
        }

        if ($changed.Count -eq 0) { Set-EditorStatus -Text 'Nothing changed.'; return }

        # Set-User only runs if something it owns actually changed -- a
        # description-only edit must not send an empty Set-User.
        if ($p.Keys.Count -gt 2) {
            if ($script:OvDirCloud) {
                Set-EditorStatus -Text 'Writing to Exchange Online (Set-User)...'
                Invoke-EXOCommand -Name 'Set-User' -Parameters $p | Out-Null
            } else {
                Set-EditorStatus -Text 'Writing to on-premises Exchange (Set-User)...'
                Invoke-OnPremCommand -Name 'Set-User' -Parameters $p | Out-Null
            }
        }
        Write-EditorAudit -Action 'MailboxDirectoryFields' -Detail ($changed -join '; ')
        # THE WRITE IS DONE AND LOGGED BY THIS POINT. Everything below is a
        # re-read, and it is several more remote calls -- so if the window
        # stops responding AFTER this message, the change has already landed
        # and it is the refresh that is stuck, not the write. Those are
        # different faults and the status line now tells them apart.
        Set-EditorStatus -Text 'Written. Re-reading the mailbox...'
        Invoke-EditorReload
        Update-OverviewTab
        Set-EditorStatus -Text $(if ($script:OvDirCloud) {
            "Applied in Exchange Online: $($changed -join '; ')."
        } else {
            "Applied on-premises: $($changed -join '; ')." + (Get-EditorSyncSuffix)
        })
    } 'Apply directory fields'
})

$btnOvRefresh.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction { Update-OverviewTab; Set-EditorStatus -Text 'Overview refreshed.' } 'Refresh overview'
})
#endregion

#region Email Address tab -------------------------------------------------------
$adY = 8
[void](New-TabLabel $tabAddresses 'EMAIL ADDRESSES' 8 $adY -Bold)
$adY += 22
$adNote = New-TabLabel $tabAddresses '' 8 $adY 1050 -Muted
$adY += 22

$gridAddr = New-TabGrid $tabAddresses @(@('Address', 460), @('Type', 90), @('Primary', 80)) 8 $adY 700 320

$adBtnY = $adY
$btnAddrPrimary = New-TabButton $tabAddresses 'Make primary...' 720 $adBtnY 160
$adBtnY += 34
$btnAddrAdd     = New-TabButton $tabAddresses 'Add alias...' 720 $adBtnY 160
$adBtnY += 34
$btnAddrRemove  = New-TabButton $tabAddresses 'Remove selected' 720 $adBtnY 160
$adBtnY += 44
$adWarn = New-TabLabel $tabAddresses '' 720 $adBtnY 340 -Muted
$adWarn.Size = New-Object System.Drawing.Size(340, 120)
$adY += 332

# ---- the e-mail address policy, directly below the list --------------------
<#
    THE SWITCH THAT DECIDES WHETHER THE LIST ABOVE IS YOURS TO EDIT.

    While EmailAddressPolicyEnabled is set, Exchange REBUILDS the proxy list
    from the address policy, so anything written here is rejected outright or
    quietly reverted. Assert-OnPremAddressEditable already refuses an
    on-premises address edit for exactly this reason -- and told the operator
    to go and clear the flag somewhere else. This is that somewhere else, on
    the tab where the problem appears.

    IT IS AN ON-PREMISES FIELD, NOT A CLOUD ONE. Microsoft documents
    -EmailAddressPolicyEnabled on Set-RemoteMailbox as applicable to Exchange
    Server 2010 through Subscription Edition only; e-mail address policies are
    an on-premises concept and the attribute lives on the directory object. So
    the control is live in On-premises mode and greyed in Cloud mode, with the
    reason beside it rather than an Apply that fails.
#>
[void](New-TabLabel $tabAddresses 'E-MAIL ADDRESS POLICY' 8 $adY -Bold)
$adY += 22
$chkAddrPolicy = New-Object System.Windows.Forms.CheckBox
$chkAddrPolicy.Location = New-Object System.Drawing.Point(8, $adY)
$chkAddrPolicy.Size     = New-Object System.Drawing.Size(660, 20)
$chkAddrPolicy.Text     = 'Automatically update e-mail addresses based on the e-mail address policy'
$tabAddresses.Controls.Add($chkAddrPolicy)
$btnAddrPolicy = New-TabButton $tabAddresses 'Apply' 680 ($adY - 3) 100
$adY += 26
$adPolNote = New-TabLabel $tabAddresses '' 8 $adY 1040 -Muted
$adPolNote.Size = New-Object System.Drawing.Size(1040, 32)
$adY += 36

function Update-AddressPolicyRow {
    <#
        Read the flag from the ON-PREMISES object, always -- it is where the
        attribute lives. In Cloud mode that object may not even be loaded, so
        the row reports what it can and offers nothing it cannot deliver.
    #>
    $onPrem = ($script:Mode -eq 'OnPrem')
    $raw    = $(if ($script:OnPrem -and $script:OnPrem.Ok) { $script:OnPrem.Raw } else { $null })
    $known  = ($raw -and $raw.PSObject.Properties['EmailAddressPolicyEnabled'])

    $chkAddrPolicy.Checked = [bool]($known -and $raw.EmailAddressPolicyEnabled)

    if (-not $onPrem) {
        $chkAddrPolicy.Enabled = $false; $btnAddrPolicy.Enabled = $false
        $adPolNote.Text = $(if ($known) {
                "Currently $(if ($raw.EmailAddressPolicyEnabled) { 'ON' } else { 'off' }) on-premises. This is an Active Directory attribute -- switch Editing to On-premises to change it."
            } else {
                'This is an on-premises attribute. Switch Editing to On-premises to read and change it.'
            })
        $adPolNote.ForeColor = (Get-ToolStatusColour -Kind Muted)
        return
    }

    if (-not $known) {
        $chkAddrPolicy.Enabled = $false; $btnAddrPolicy.Enabled = $false
        $adPolNote.Text = 'The on-premises object does not report this attribute, so it is not offered. An object with no mailbox has no address policy to apply.'
        $adPolNote.ForeColor = (Get-ToolStatusColour -Kind Warn)
        return
    }
    if (-not $script:OnPrem.SetCmdlet) {
        $chkAddrPolicy.Enabled = $false; $btnAddrPolicy.Enabled = $false
        $adPolNote.Text = 'This on-premises object is a plain directory user; there is no recipient cmdlet to write this with.'
        $adPolNote.ForeColor = (Get-ToolStatusColour -Kind Warn)
        return
    }

    $chkAddrPolicy.Enabled = $true; $btnAddrPolicy.Enabled = $true
    $adPolNote.Text = $(if ($raw.EmailAddressPolicyEnabled) {
            'ON: Exchange builds this mailbox''s addresses from the policy, and will revert or reject anything set in the list above. Untick and Apply to take manual control.'
        } else {
            'Off: the addresses above are maintained by hand. Ticking this hands them back to the policy, which may REPLACE them.'
        })
    $adPolNote.ForeColor = (Get-ToolStatusColour -Kind Muted)
}

$btnAddrPolicy.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if ($script:Mode -ne 'OnPrem') { Set-EditorStatus -Text 'Switch Editing to On-premises first.' -Warn; return }
        if (-not $script:OnPrem -or -not $script:OnPrem.Ok -or -not $script:OnPrem.SetCmdlet) {
            Set-EditorStatus -Text 'No writable on-premises object is loaded.' -Bad; return
        }
        $now  = [bool]$script:OnPrem.Raw.EmailAddressPolicyEnabled
        $want = [bool]$chkAddrPolicy.Checked
        if ($now -eq $want) { Set-EditorStatus -Text 'Already set that way.'; return }

        # THE TWO DIRECTIONS ARE NOT THE SAME RISK, so they do not share a
        # sentence. Turning it OFF freezes the current addresses. Turning it
        # ON hands them to the policy, which can replace the primary -- and
        # that changes what the person's mail comes from.
        $msg = $(if ($want) {
                ("Hand $($script:Mbx.Display)'s addresses back to the e-mail address policy?" + "`r`n`r`n" +
                 "Exchange will REBUILD the address list from the policy at the next stamp. Addresses the policy does not produce can be removed, and the PRIMARY address can change -- which changes what this person's mail appears to come from." + "`r`n`r`n" +
                 "X500 entries are normally preserved, but check the list afterwards.")
            } else {
                ("Take manual control of $($script:Mbx.Display)'s addresses?" + "`r`n`r`n" +
                 "The current addresses stay exactly as they are and the policy stops maintaining them. That also means future policy changes will NOT reach this mailbox -- it has to be maintained by hand from now on." + "`r`n`r`n" +
                 "This is what has to be off before an address on this tab can be edited on-premises.")
            })
        $a = Show-ToolDialog($msg, 'E-mail address policy', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        Invoke-OnPremCommand -Name $script:OnPrem.SetCmdlet -Parameters @{
            Identity = $script:OnPrem.Smtp; EmailAddressPolicyEnabled = $want; ErrorAction = 'Stop' } | Out-Null
        Write-EditorAudit -Action 'MailboxAddressPolicy' -Environment 'OnPrem' `
            -Detail "EmailAddressPolicyEnabled $now -> $want via $($script:OnPrem.SetCmdlet)."

        $script:OnPrem = Resolve-OnPremMailboxTarget -Identity $script:OnPrem.Smtp
        Update-AddressesTab
        Set-EditorStatus -Text "E-mail address policy $(if ($want) { 'enabled' } else { 'disabled' }). $($script:OnPremSyncNotice)"
    } 'Change the e-mail address policy setting'
})

function Update-AddressesTab {
    $gridAddr.Rows.Clear()
    # THE LIST COMES FROM WHICHEVER OBJECT IS LOADED. Showing the cloud
    # addresses under an on-premises banner would invite somebody to remove an
    # address that is not in the collection they are about to write.
    $onPrem = ($script:Mode -eq 'OnPrem')
    if ($onPrem) {
        if (-not $script:OnPrem -or -not $script:OnPrem.Ok) {
            $adNote.Text = 'No on-premises object is loaded.'
            $btnAddrPrimary.Enabled = $false; $btnAddrAdd.Enabled = $false; $btnAddrRemove.Enabled = $false
            return
        }
        $addresses = @($script:OnPrem.Addresses)
    } else {
        $addresses = @($script:Mbx.Raw.EmailAddresses)
    }
    foreach ($a in $addresses) {
        $s = "$a"
        if (-not $s) { continue }
        $prefix = ''; $addr = $s
        if ($s -match '^([A-Za-z0-9]+):(.*)$') { $prefix = $Matches[1]; $addr = $Matches[2] }
        $isPrimary = ($prefix -ceq 'SMTP')
        $kind = $(if ($prefix -match '(?i)^smtp$') { 'SMTP' } else { $prefix })
        $i = $gridAddr.Rows.Add($addr, $kind, $(if ($isPrimary) { 'PRIMARY' } else { '' }))
        $gridAddr.Rows[$i].Tag = $s
        if ($isPrimary) {
            $gridAddr.Rows[$i].DefaultCellStyle.Font = New-Object System.Drawing.Font('Consolas', 9, [System.Drawing.FontStyle]::Bold)
        } elseif ($kind -notmatch '(?i)^smtp$') {
            # X500 and the rest are dimmed: they are not addresses anyone
            # types, and they must survive any rebuild of the list.
            $gridAddr.Rows[$i].DefaultCellStyle.ForeColor = (Get-ToolStatusColour -Kind Muted)
        }
    }
    if ($onPrem) {
        $can = [bool]$script:OnPrem.SetCmdlet
        $adNote.Text = "$($gridAddr.Rows.Count) on-premises address(es). The primary is in bold; X500 entries are what let replies to old mail resolve and are never removed automatically."
        $adWarn.Text = $(if ($can) { "Changes are written with $($script:OnPrem.SetCmdlet). $($script:OnPremSyncNotice)" }
                         else { 'This on-premises object has no mailbox, so it has no address list to change.' })
        $adWarn.ForeColor = (Get-ToolStatusColour -Kind Warn)
        $btnAddrPrimary.Enabled = $can; $btnAddrAdd.Enabled = $can; $btnAddrRemove.Enabled = $can
        Update-AddressPolicyRow
        return
    }
    $t = Get-MailboxWriteTarget -Field 'EmailAddresses' -IsSynced $script:Mbx.IsSynced
    if ($t.State -eq 'Ok') {
        $adNote.Text  = "$($gridAddr.Rows.Count) address(es). The primary is in bold; X500 entries are what let replies to old mail resolve and are never removed automatically."
        $adWarn.Text  = ''
        $btnAddrPrimary.Enabled = $true; $btnAddrAdd.Enabled = $true; $btnAddrRemove.Enabled = $true
    } else {
        $adNote.Text  = "$($gridAddr.Rows.Count) address(es). READ ONLY for this mailbox."
        $adWarn.Text  = $t.Reason
        $adWarn.ForeColor = (Get-ToolStatusColour -Kind Warn)
        $btnAddrPrimary.Enabled = $false; $btnAddrAdd.Enabled = $false; $btnAddrRemove.Enabled = $false
    }
    Update-AddressPolicyRow
}

$btnAddrPrimary.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorFieldWritable -Field 'EmailAddresses')) { return }
        if ($gridAddr.SelectedRows.Count -eq 0) { Set-EditorStatus -Text 'Select the address to make primary.' -Bad; return }
        $row  = $gridAddr.SelectedRows[0]
        $addr = "$($row.Cells['Address'].Value)"
        $kind = "$($row.Cells['Type'].Value)"
        if ($kind -notmatch '(?i)^smtp$') { Set-EditorStatus -Text 'Only an SMTP address can be the primary.' -Bad; return }
        if ("$($row.Cells['Primary'].Value)") { Set-EditorStatus -Text "$addr is already the primary."; return }

        $a = Show-ToolDialog(
            ("Make $addr the primary address for $($script:Mbx.Display)?" + "`r`n`r`n" +
             "The current primary is kept as an alias, so mail to it still delivers." + "`r`n" +
             "Every other entry, including X500, is carried across unchanged." + "`r`n`r`n" +
             "Outlook profiles and Teams can take a while to catch up."),
            'Change primary address', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        # The SAME function Unified Lookup's Attributes editor calls. One
        # implementation of the EmailAddresses rebuild, so the X500 handling
        # cannot drift between the two front doors.
        #
        # RE-READ THE ADDRESSES FIRST. Set-CloudPrimaryAddress rebuilds and
        # REPLACES the whole collection, so it has to be given the collection
        # as it is NOW. Handing it the snapshot taken at Load would delete any
        # alias or X500 added since -- by another admin, or by a mailbox move
        # that stamps one -- and an X500 deleted that way breaks replies to
        # every message already sent.
        if ($script:Mode -eq 'OnPrem') {
            Set-OnPremPrimaryAddress -Target $script:OnPrem -NewPrimary $addr
            Write-EditorAudit -Action 'MailboxPrimaryAddress' -Detail "Primary address changed to $addr on-premises."
            $script:OnPrem = Resolve-OnPremMailboxTarget -Identity $addr
        } else {
            $fresh = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
                        Identity = $script:Mbx.Smtp; ErrorAction = 'Stop' }
            $plan = [PSCustomObject]@{ Cmdlet = 'Set-Mailbox'; PrimaryParam = '' }
            Set-CloudPrimaryAddress -Plan $plan -Identity $script:Mbx.Smtp -NewPrimary $addr `
                -ExistingAddresses @($fresh.EmailAddresses)
            Write-EditorAudit -Action 'MailboxPrimaryAddress' -Detail "Primary address changed to $addr."
            $script:Mbx = Resolve-MailboxTarget -Smtp $addr
        }
        Update-EditorIdentity
        Update-AddressesTab; $script:Loaded['Overview'] = $false
        Set-EditorStatus -Text ("Primary address is now $addr." + (Get-EditorSyncSuffix))
    } 'Change primary address'
})

$btnAddrAdd.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorFieldWritable -Field 'EmailAddresses')) { return }
        $new = Show-InputDialog -Title 'Add alias' -Width 520 `
                -Prompt "New SMTP alias for $($script:Mbx.Display). It becomes an alias, not the primary:"
        $new = "$new".Trim()
        if (-not $new) { return }
        if ($new -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { Set-EditorStatus -Text 'That does not look like an email address.' -Bad; return }
        Invoke-EditorSetMailbox -Parameters @{ EmailAddresses = @{ Add = "smtp:$new" } }
        Write-EditorAudit -Action 'MailboxAddressAdd' -Detail "Alias $new added."
        Invoke-EditorReload
        Update-AddressesTab; $script:Loaded['Overview'] = $false
        Set-EditorStatus -Text ("Added alias $new." + (Get-EditorSyncSuffix))
    } 'Add alias'
})

$btnAddrRemove.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorFieldWritable -Field 'EmailAddresses')) { return }
        if ($gridAddr.SelectedRows.Count -eq 0) { Set-EditorStatus -Text 'Select the address to remove.' -Bad; return }
        $row  = $gridAddr.SelectedRows[0]
        $full = "$($row.Tag)"
        $addr = "$($row.Cells['Address'].Value)"
        if ("$($row.Cells['Primary'].Value)") {
            $null = Show-ToolDialog('The primary address cannot be removed. Make another address primary first -- the old one is kept as an alias.',
                                    'Primary address', 'OK', 'Information')
            return
        }
        # X500 REMOVAL IS THE DANGEROUS ONE, so it gets its own sentence
        # rather than the generic confirmation.
        $extra = ''
        if ("$($row.Cells['Type'].Value)" -match '(?i)^x500$') {
            $extra = "`r`n`r`nTHIS IS AN X500 ADDRESS. It is what lets replies to older mail find this mailbox. Removing it makes those replies bounce, and the failure appears weeks later."
        }
        $a = Show-ToolDialog(("Remove $addr from $($script:Mbx.Display)?" + "`r`n`r`nMail sent to it will stop delivering." + $extra),
                             'Remove address', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Invoke-EditorSetMailbox -Parameters @{ EmailAddresses = @{ Remove = $full } }
        Write-EditorAudit -Action 'MailboxAddressRemove' -Detail "Address $full removed."
        Invoke-EditorReload
        Update-AddressesTab; $script:Loaded['Overview'] = $false
        Set-EditorStatus -Text ("Removed $addr." + (Get-EditorSyncSuffix))
    } 'Remove address'
})
#endregion

#region Full Access tab ---------------------------------------------------------
<#
    NOTHING ON THIS TAB IS TYPED ANY MORE.

    The first version asked for the trustee in one prompt and the words
    'FullAccess', 'SendAs' or 'SendOnBehalf' in another. Three problems with
    that, and the third is the one that mattered: a typo in the second prompt
    was rejected with a message about valid values; granting the same person
    two kinds of access meant doing the whole thing twice; and granting four
    people access to a shared mailbox meant doing it eight times.

    So: tick boxes for the kinds, a list for the people, and a list for the
    MAILBOXES -- because the request that arrives is 'give these three people
    access to these two shared mailboxes', not one grant at a time.

    THE WINDOW STILL NAMES ONE MAILBOX, and that is a real tension. It is
    resolved by auditing each mailbox against its own address and by reporting
    per mailbox rather than saying 'done' -- so a grant that worked on two of
    three is visible as exactly that.
#>
$acY = 8
[void](New-TabLabel $tabAccess 'WHO CAN GET AT THIS MAILBOX' 8 $acY -Bold)
$acY += 22
$acNote = New-TabLabel $tabAccess '' 8 $acY 1050 -Muted
$acY += 22

$gridAccess = New-TabGrid $tabAccess @(@('Kind', 120), @('User', 380), @('Note', 120)) 8 $acY 700 300
Add-GridPickColumn -Grid $gridAccess

$acBtnY = $acY
$btnAccAdd     = New-TabButton $tabAccess 'Grant access...'    720 $acBtnY 170
$acBtnY += 34
$btnAccChange  = New-TabButton $tabAccess 'Change ticked...'   720 $acBtnY 170
$acBtnY += 34
$btnAccRemove  = New-TabButton $tabAccess 'Remove ticked'      720 $acBtnY 170
$acBtnY += 34
# 'Fix automapping...' LEFT THIS TAB ON 2026-09-19 for its own tab, which does
# the job the way the standalone tool did: remove, hold ten seconds, re-add,
# read back, and say what to tell the user.
$btnAccAll     = New-TabButton $tabAccess 'Tick all'            720 $acBtnY 82
$btnAccNone    = New-TabButton $tabAccess 'None'                808 $acBtnY 82
$acBtnY += 40
$acHint = New-TabLabel $tabAccess '' 720 $acBtnY 340 -Muted
$acHint.Size = New-Object System.Drawing.Size(340, 100)
$acHint.Text = ('Full Access opens the mailbox. Send As sends from it as the mailbox. ' +
                'Send on Behalf sends showing both names.' + "`r`n`r`n" +
                'Change and Remove act on every TICKED row, not on the highlighted one.' + "`r`n`r`n" +
                'A mailbox that will not appear in someone''s Outlook: the Automap Fix tab.')
$acY += 312

# The two MessageCopy settings sit HERE rather than in mail flow, because the
# question they answer -- 'why are sent items not showing up in the shared
# mailbox' -- arrives while somebody is looking at who has Send As.
[void](New-TabLabel $tabAccess 'COPIES OF SENT MAIL' 8 $acY -Bold)
$acY += 22
$chkCopySentAs = New-Object System.Windows.Forms.CheckBox
$chkCopySentAs.Text     = 'Keep a copy in this mailbox when someone sends AS it'
$chkCopySentAs.Location = New-Object System.Drawing.Point(8, $acY)
$chkCopySentAs.Size     = New-Object System.Drawing.Size(460, 20)
$tabAccess.Controls.Add($chkCopySentAs)
$acY += 24
$chkCopyOnBehalf = New-Object System.Windows.Forms.CheckBox
$chkCopyOnBehalf.Text     = 'Keep a copy in this mailbox when someone sends ON BEHALF OF it'
$chkCopyOnBehalf.Location = New-Object System.Drawing.Point(8, $acY)
$chkCopyOnBehalf.Size     = New-Object System.Drawing.Size(460, 20)
$tabAccess.Controls.Add($chkCopyOnBehalf)
$btnCopyApply = New-TabButton $tabAccess 'Apply' 480 ($acY - 14) 100
$acY += 30

# The three kinds, in one place. The label is what a person says; the Key is
# what Add-MailboxAccessEntry wants. Typing the second was the old way.
$script:AccessKind = @(
    [PSCustomObject]@{ Key = 'FullAccess';   Label = 'Full Access -- open the mailbox' }
    [PSCustomObject]@{ Key = 'SendAs';       Label = 'Send As -- send as the mailbox' }
    [PSCustomObject]@{ Key = 'SendOnBehalf'; Label = 'Send on Behalf -- send showing both names' }
)

function Update-AccessTab {
    $gridAccess.Rows.Clear()
    $m = $script:Mbx
    $r = Get-MailboxAccessEntry -Smtp $m.Smtp
    foreach ($e in @($r.Entries)) {
        $i = $gridAccess.Rows.Add($false, $e.Kind, $e.User, $e.Detail)
        $gridAccess.Rows[$i].Tag = $e
        if ($e.Detail -eq 'DENY') { $gridAccess.Rows[$i].DefaultCellStyle.ForeColor = (Get-ToolStatusColour -Kind Bad) }
    }
    $acNote.Text = "$($gridAccess.Rows.Count) grant(s). Inherited and self entries are not shown. $($r.Note)"
    # A READ THAT REPORTED A PROBLEM IS NOT A LOADED TAB. Without this the
    # grid is cached empty after a transient failure and re-selecting the tab
    # will not try again -- it looks like the mailbox has no grants.
    if ("$($r.Note)".Trim()) { $script:TabReadFailed = $true }

    $mb = $m.Raw
    $chkCopySentAs.Checked   = [bool]$mb.MessageCopyForSentAsEnabled
    $chkCopyOnBehalf.Checked = [bool]$mb.MessageCopyForSendOnBehalfEnabled
}

function Show-AccessGrantDialog {
    <#
        Who, on what, with which kinds of access -- all on one dialog.

        The loaded mailbox is prefilled and the box is editable rather than
        fixed, because the common request names more than one mailbox and
        retyping the first one is the sort of friction that sends people back
        to the shell.

        Returns $null when cancelled, otherwise the three lists and the
        automapping answer.
    #>
    param([string]$DefaultMailbox, [string]$DisplayName)

    $f = New-Object System.Windows.Forms.Form
    $f.Text          = 'Grant access'
    $f.ClientSize    = New-Object System.Drawing.Size(720, 430)
    $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'
    $f.MaximizeBox   = $false
    $f.MinimizeBox   = $false
    $f.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $l1 = New-Object System.Windows.Forms.Label
    $l1.Text = "MAILBOXES to grant access ON -- one per line"
    $l1.Location = New-Object System.Drawing.Point(12, 12); $l1.Size = New-Object System.Drawing.Size(340, 18)
    $l1.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
    $f.Controls.Add($l1)
    $txtMbxList = New-DialogList $f 12 34 340 150 $DefaultMailbox

    $l2 = New-Object System.Windows.Forms.Label
    $l2.Text = 'PEOPLE to give access TO -- one per line'
    $l2.Location = New-Object System.Drawing.Point(368, 12); $l2.Size = New-Object System.Drawing.Size(340, 18)
    $l2.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
    $f.Controls.Add($l2)
    $txtWhoList = New-DialogList $f 368 34 340 150 ''

    $l3 = New-Object System.Windows.Forms.Label
    $l3.Text = 'ACCESS to grant -- tick as many as apply'
    $l3.Location = New-Object System.Drawing.Point(12, 196); $l3.Size = New-Object System.Drawing.Size(400, 18)
    $l3.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
    $f.Controls.Add($l3)
    $y = 218
    $checks = @{}
    foreach ($k in $script:AccessKind) {
        $checks[$k.Key] = New-DialogCheck $f $k.Label 20 $y 420
        $y += 24
    }

    $chkAuto = New-DialogCheck $f 'Automapping: open it in their Outlook automatically (Full Access only)' 20 ($y + 6) 640 $true
    $y += 34
    $note = New-Object System.Windows.Forms.Label
    $note.Text = ('Automapping is right when one person is covering another, and wrong for a busy shared mailbox -- ' +
                  'it lands in every profile and cannot be removed per person without restamping the permission.')
    $note.Location = New-Object System.Drawing.Point(20, ($y + 4)); $note.Size = New-Object System.Drawing.Size(680, 34)
    $note.ForeColor = (Get-ToolStatusColour -Kind Muted)
    $f.Controls.Add($note)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Grant'; $ok.Location = New-Object System.Drawing.Point(508, 388); $ok.Size = New-Object System.Drawing.Size(96, 28)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $f.Controls.Add($ok)
    $ca = New-Object System.Windows.Forms.Button
    $ca.Text = 'Cancel'; $ca.Location = New-Object System.Drawing.Point(612, 388); $ca.Size = New-Object System.Drawing.Size(96, 28)
    $ca.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $f.Controls.Add($ca)
    $f.AcceptButton = $ok; $f.CancelButton = $ca

    Set-ConsoleTheme -Form $f
    $r = $f.ShowDialog()
    $res = $null
    if ($r -eq [System.Windows.Forms.DialogResult]::OK) {
        $kinds = @()
        foreach ($k in $script:AccessKind) { if ($checks[$k.Key].Checked) { $kinds += $k.Key } }
        $res = [PSCustomObject]@{
            Mailboxes = (ConvertFrom-DialogList -Text $txtMbxList.Text -CompleteDomain)
            Trustees  = (ConvertFrom-DialogList -Text $txtWhoList.Text -CompleteDomain)
            Kinds     = @($kinds)
            AutoMap   = [bool]$chkAuto.Checked
        }
    }
    $f.Dispose()
    return $res
}

function Show-AccessChangeDialog {
    <#
        Add and remove kinds of access across every ticked row at once.

        TWO COLUMNS, ADD AND REMOVE, rather than one set of checkboxes. A
        single tri-state column would have to mean 'leave alone' somewhere, and
        a checkbox that has three meanings is a checkbox nobody trusts. Two
        explicit lists cannot be misread.
    #>
    param([string[]]$Trustees, [string]$MailboxName)

    $f = New-Object System.Windows.Forms.Form
    $f.Text          = 'Change access'
    $f.ClientSize    = New-Object System.Drawing.Size(640, 330)
    $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'
    $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $l0 = New-Object System.Windows.Forms.Label
    $l0.Text = "Applies to $(@($Trustees).Count) person/people on $MailboxName" + ":"
    $l0.Location = New-Object System.Drawing.Point(12, 12); $l0.Size = New-Object System.Drawing.Size(616, 18)
    $l0.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
    $f.Controls.Add($l0)

    $who = New-Object System.Windows.Forms.TextBox
    $who.Multiline = $true; $who.ReadOnly = $true; $who.ScrollBars = 'Vertical'
    $who.Location = New-Object System.Drawing.Point(12, 34); $who.Size = New-Object System.Drawing.Size(616, 70)
    $who.Font = New-Object System.Drawing.Font('Consolas', 9)
    $who.Text = (@($Trustees) -join "`r`n")
    $who.Tag = 'NoTheme'; $who.BackColor = [System.Drawing.Color]::WhiteSmoke; $who.ForeColor = [System.Drawing.Color]::Black
    $f.Controls.Add($who)

    $lAdd = New-Object System.Windows.Forms.Label
    $lAdd.Text = 'GIVE them:'; $lAdd.Location = New-Object System.Drawing.Point(12, 116); $lAdd.Size = New-Object System.Drawing.Size(300, 18)
    $lAdd.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
    $f.Controls.Add($lAdd)
    $lDel = New-Object System.Windows.Forms.Label
    $lDel.Text = 'TAKE AWAY:'; $lDel.Location = New-Object System.Drawing.Point(330, 116); $lDel.Size = New-Object System.Drawing.Size(300, 18)
    $lDel.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
    $f.Controls.Add($lDel)

    $y = 138
    $add = @{}; $del = @{}
    foreach ($k in $script:AccessKind) {
        $short = ($k.Label -split ' -- ')[0]
        $add[$k.Key] = New-DialogCheck $f $short 20 $y 290
        $del[$k.Key] = New-DialogCheck $f $short 338 $y 290
        $y += 24
    }

    $chkAuto = New-DialogCheck $f 'Automapping for any Full Access being GIVEN' 20 ($y + 10) 500 $true
    $y += 40
    $note = New-Object System.Windows.Forms.Label
    $note.Text = 'Anything ticked in both columns is taken away and given back -- which is how automapping gets restamped.'
    $note.Location = New-Object System.Drawing.Point(20, ($y + 2)); $note.Size = New-Object System.Drawing.Size(600, 34)
    $note.ForeColor = (Get-ToolStatusColour -Kind Muted)
    $f.Controls.Add($note)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Apply'; $ok.Location = New-Object System.Drawing.Point(428, 290); $ok.Size = New-Object System.Drawing.Size(96, 28)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK; $f.Controls.Add($ok)
    $ca = New-Object System.Windows.Forms.Button
    $ca.Text = 'Cancel'; $ca.Location = New-Object System.Drawing.Point(532, 290); $ca.Size = New-Object System.Drawing.Size(96, 28)
    $ca.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $f.Controls.Add($ca)
    $f.AcceptButton = $ok; $f.CancelButton = $ca

    Set-ConsoleTheme -Form $f
    $r = $f.ShowDialog()
    $res = $null
    if ($r -eq [System.Windows.Forms.DialogResult]::OK) {
        $a = @(); $d = @()
        foreach ($k in $script:AccessKind) {
            if ($add[$k.Key].Checked) { $a += $k.Key }
            if ($del[$k.Key].Checked) { $d += $k.Key }
        }
        $res = [PSCustomObject]@{ Add = @($a); Remove = @($d); AutoMap = [bool]$chkAuto.Checked }
    }
    $f.Dispose()
    return $res
}

$btnAccAll.Add_Click({  Set-GridPicked -Grid $gridAccess -On $true })
$btnAccNone.Add_Click({ Set-GridPicked -Grid $gridAccess -On $false })

$btnAccAdd.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Granting access')) { return }
        $d = Show-AccessGrantDialog -DefaultMailbox $script:Mbx.Smtp -DisplayName $script:Mbx.Display
        if (-not $d) { return }
        if ($d.Mailboxes.Count -eq 0) { Set-EditorStatus -Text 'No mailboxes listed.' -Bad; return }
        if ($d.Trustees.Count -eq 0)  { Set-EditorStatus -Text 'No people listed.' -Bad; return }
        if ($d.Kinds.Count -eq 0)     { Set-EditorStatus -Text 'Tick at least one kind of access.' -Bad; return }

        # SEND ON BEHALF IS MASTERED IN THE DIRECTORY ON A SYNCED MAILBOX, and
        # the gate has to be asked of EVERY mailbox in the list -- not of the
        # one in the banner. Checking only the loaded mailbox meant a
        # cloud-only mailbox in the box suppressed the warning for the synced
        # ones beside it, whose write directory sync would quietly revert.
        #
        # One extra read per mailbox, against N x M x K writes. Cheap.
        $syncedSob = @()
        if ($d.Kinds -contains 'SendOnBehalf') {
            foreach ($mbx in $d.Mailboxes) {
                try {
                    $t = Resolve-MailboxTarget -Smtp $mbx
                    if ($t.Ok -and $t.IsSynced) { $syncedSob += $mbx }
                } catch { }
            }
            if ($syncedSob.Count) {
                $w = Show-ToolDialog(
                    ("Send on Behalf is mastered in Active Directory for these mailboxes:" + "`r`n  " +
                     ($syncedSob -join "`r`n  ") + "`r`n`r`n" +
                     'Exchange Online will accept the change and directory sync will put it back, usually within half an hour, with no error anywhere.' + "`r`n`r`n" +
                     'Full Access and Send As are unaffected and will apply normally. Go ahead anyway?'),
                    'Send on Behalf will not stick', 'OKCancel', 'Warning')
                if ($w -ne [System.Windows.Forms.DialogResult]::OK) { return }
            }
        }

        $total = $d.Mailboxes.Count * $d.Trustees.Count * $d.Kinds.Count
        $a = Show-ToolDialog(
            ("Grant $($d.Kinds -join ', ') to $($d.Trustees.Count) person/people on $($d.Mailboxes.Count) mailbox(es)?" +
             "`r`n`r`nMailboxes:`r`n  " + ($d.Mailboxes -join "`r`n  ") +
             "`r`n`r`nPeople:`r`n  " + ($d.Trustees -join "`r`n  ") +
             "`r`n`r`n$total grant(s) in total." +
             $(if ($d.Kinds -contains 'FullAccess') { "`r`nAutomapping: $($d.AutoMap)" })),
            'Grant access', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $lines = @(); $ok = 0; $bad = 0
        foreach ($mbx in $d.Mailboxes) {
            foreach ($who in $d.Trustees) {
                foreach ($kind in $d.Kinds) {
                    try {
                        Add-MailboxAccessEntry -Smtp $mbx -Kind $kind -Trustee $who -AutoMap $d.AutoMap
                        $lines += "OK      $mbx  <-  $who  ($kind)"
                        $ok++
                        # AUDITED AGAINST THE MAILBOX THAT WAS ACTUALLY CHANGED,
                        # not against the one in the banner. A bulk grant whose
                        # log all points at one mailbox is a log that cannot
                        # answer 'who gave them access to that one'.
                        Write-CompletedLog -Action 'MailboxAccessAdd' -Target (Get-LogTarget $mbx) `
                            -Result 'Success' -Environment 'Cloud' `
                            -Detail "$kind granted to $who$(if ($kind -eq 'FullAccess') { "; automapping $($d.AutoMap)" }) (Mailbox Editor, bulk)"
                    } catch {
                        $lines += "FAILED  $mbx  <-  $who  ($kind): $($_.Exception.Message)"
                        $bad++
                        Write-FailedLog -Action 'MailboxAccessAdd' -Target (Get-LogTarget $mbx) `
                            -ErrorRecord $_ -Environment 'Cloud' -Detail "$kind for $who"
                    }
                }
            }
        }
        Show-ResultListDialog -Title "Grant access -- $ok done, $bad failed" -Lines $lines
        Update-AccessTab
        Set-EditorStatus -Text "Granted $ok, failed $bad." -Bad:($bad -gt 0)
    } 'Grant access'
})

$btnAccChange.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Changing access')) { return }
        $picked = Get-GridPicked -Grid $gridAccess
        if ($picked.Count -eq 0) { Set-EditorStatus -Text 'Tick one or more rows to change.' -Bad; return }
        # ONE PERSON MAY HOLD THREE ROWS. The dialog is about PEOPLE, so the
        # ticked rows are reduced to the distinct trustees behind them --
        # otherwise ticking somebody's three grants would apply the change to
        # them three times.
        $who = @($picked | ForEach-Object { "$($_.User)" } | Select-Object -Unique)

        $d = Show-AccessChangeDialog -Trustees $who -MailboxName $script:Mbx.Display
        if (-not $d) { return }
        if ($d.Add.Count -eq 0 -and $d.Remove.Count -eq 0) { Set-EditorStatus -Text 'Nothing ticked to give or take away.'; return }
        if ($d.Add -contains 'SendOnBehalf' -or $d.Remove -contains 'SendOnBehalf') {
            if (-not (Test-EditorFieldWritable -Field 'GrantSendOnBehalfTo')) { return }
        }

        $a = Show-ToolDialog(
            ("On $($script:Mbx.Display), for:" + "`r`n  " + ($who -join "`r`n  ") + "`r`n`r`n" +
             "Give:       $(if ($d.Add) { $d.Add -join ', ' } else { '(nothing)' })" + "`r`n" +
             "Take away:  $(if ($d.Remove) { $d.Remove -join ', ' } else { '(nothing)' })"),
            'Change access', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        # A KIND TICKED IN BOTH COLUMNS IS A RESTAMP, NOT A REMOVE AND AN ADD.
        #
        # Doing it as two independent operations is what Repair-MailboxAutoMapping
        # exists to avoid: the re-add commonly fails on throttling, and a bare
        # pair leaves the trustee with the permission REMOVED and nothing put
        # back -- reported as one FAILED line in a list. The repair function
        # sleeps, retries three times, and if it still cannot restore access
        # says so in plain words. So anything in both lists goes through it.
        $both   = @($d.Add | Where-Object { $d.Remove -contains $_ })
        $addOnly = @($d.Add    | Where-Object { $both -notcontains $_ })
        $delOnly = @($d.Remove | Where-Object { $both -notcontains $_ })

        $lines = @(); $ok = 0; $bad = 0
        foreach ($u in $who) {
            foreach ($kind in $delOnly) {
                try { Remove-MailboxAccessEntry -Smtp $script:Mbx.Smtp -Kind $kind -Trustee $u; $lines += "OK      removed $kind from $u"; $ok++ }
                catch { $lines += "FAILED  removing $kind from $u`: $($_.Exception.Message)"; $bad++ }
            }
        }
        foreach ($u in $who) {
            foreach ($kind in $addOnly) {
                try { Add-MailboxAccessEntry -Smtp $script:Mbx.Smtp -Kind $kind -Trustee $u -AutoMap $d.AutoMap; $lines += "OK      gave $kind to $u"; $ok++ }
                catch { $lines += "FAILED  giving $kind to $u`: $($_.Exception.Message)"; $bad++ }
            }
        }
        foreach ($u in $who) {
            foreach ($kind in $both) {
                try {
                    if ($kind -eq 'FullAccess') {
                        Repair-MailboxAutoMapping -Smtp $script:Mbx.Smtp -Trustee $u -AutoMap $d.AutoMap
                        $lines += "OK      restamped FullAccess for $u (automapping $($d.AutoMap))"
                    } else {
                        Remove-MailboxAccessEntry -Smtp $script:Mbx.Smtp -Kind $kind -Trustee $u
                        Add-MailboxAccessEntry -Smtp $script:Mbx.Smtp -Kind $kind -Trustee $u -AutoMap $d.AutoMap
                        $lines += "OK      restamped $kind for $u"
                    }
                    $ok++
                } catch { $lines += "FAILED  restamping $kind for $u`: $($_.Exception.Message)"; $bad++ }
            }
        }
        Write-EditorAudit -Action 'MailboxAccessChange' `
            -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
            -Detail "give [$($d.Add -join ', ')], take [$($d.Remove -join ', ')] for $($who -join '; ')"
        Show-ResultListDialog -Title "Change access -- $ok done, $bad failed" -Lines $lines
        Update-AccessTab
        Set-EditorStatus -Text "Changed $ok, failed $bad." -Bad:($bad -gt 0)
    } 'Change access'
})

$btnAccRemove.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Removing access')) { return }
        $picked = Get-GridPicked -Grid $gridAccess
        if ($picked.Count -eq 0) { Set-EditorStatus -Text 'Tick one or more grants to remove.' -Bad; return }
        $names = @($picked | ForEach-Object { "  $($_.Kind)  --  $($_.User)" }) -join "`r`n"
        $a = Show-ToolDialog(("Remove $($picked.Count) grant(s) from $($script:Mbx.Display)?" + "`r`n`r`n" + $names),
                             'Remove access', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $ok = 0; $bad = 0; $lines = @()
        foreach ($e in $picked) {
            try { Remove-MailboxAccessEntry -Smtp $script:Mbx.Smtp -Kind $e.Kind -Trustee $e.User; $lines += "OK      $($e.Kind)  $($e.User)"; $ok++ }
            catch { $lines += "FAILED  $($e.Kind)  $($e.User): $($_.Exception.Message)"; $bad++ }
        }
        Write-EditorAudit -Action 'MailboxAccessRemove' `
            -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
            -Detail "$ok removed, $bad failed: $(($picked | ForEach-Object { "$($_.Kind) $($_.User)" }) -join '; ')"
        if ($bad) { Show-ResultListDialog -Title "Remove access -- $ok done, $bad failed" -Lines $lines }
        Update-AccessTab
        Set-EditorStatus -Text "Removed $ok, failed $bad." -Bad:($bad -gt 0)
    } 'Remove access'
})

$btnCopyApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'The sent-items copy settings')) { return }
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
            Identity                          = $script:Mbx.Smtp
            MessageCopyForSentAsEnabled       = [bool]$chkCopySentAs.Checked
            MessageCopyForSendOnBehalfEnabled = [bool]$chkCopyOnBehalf.Checked
            ErrorAction                       = 'Stop'
        } | Out-Null
        Write-EditorAudit -Action 'MailboxSentItemsCopy' `
            -Detail "SentAs copy = $($chkCopySentAs.Checked); SendOnBehalf copy = $($chkCopyOnBehalf.Checked)."
        $script:Mbx = Resolve-MailboxTarget -Smtp $script:Mbx.Smtp
        Set-EditorStatus -Text 'Sent-items copy settings applied.'
    } 'Apply sent-items copy settings'
})
#endregion


#region Automap Fix tab ---------------------------------------------------------
<#
    THE AUTOMAP REPAIR, ON ITS OWN TAB, THE WAY THE STANDALONE TOOL DID IT.

    The Full Access tab had a 'Fix automapping...' button that did the job in two
    seconds with a retry. Glenn's push-back, 2026-09-19: the job is not two
    seconds and a retry. It is remove, WAIT TEN SECONDS SO THE REMOVAL SETTLES,
    re-add, READ BACK, and then tell the operator what to tell the user --
    because the user watching their Outlook will see nothing happen and
    conclude it failed. That is a tab's worth of window, not a button's, and
    Automap-Fix.ps1 stays until this tab does all of it. Now it does.

    WHY THE FIX IS A RE-GRANT AT ALL. Automapping is decided when Full Access
    is granted (msExchDelegateListLink is stamped then) and never revisited.
    Exchange Online does not even report it back, so the tab cannot say
    whether it is currently on -- only that the permission exists. Granting
    again is the only way to change it.

    ONLY DIRECT FULL ACCESS IS LISTED. Access through a group never automaps,
    however many times it is re-granted, and inherited entries are not grants
    anybody made. A person who is not in this grid has nothing here to fix,
    and the note says what to do instead.
#>
$amY = 8
[void](New-TabLabel $tabAutomap 'FULL ACCESS THAT WILL NOT SHOW IN OUTLOOK' 8 $amY -Bold)
$amY += 22
$amNote = New-TabLabel $tabAutomap '' 8 $amY 1050 -Muted
$amNote.Size = New-Object System.Drawing.Size(1050, 34)
$amNote.Text = ('Automapping is set when Full Access is GRANTED and never revisited, and Exchange Online does not report it back. ' +
                'The fix is to remove the grant, wait for it to settle, and grant it again. Only direct Full Access is listed: access through a group never automaps.')
$amY += 38

$gridAutomap = New-TabGrid $tabAutomap @(@('User', 420), @('Note', 200)) 8 $amY 640 200
Add-GridPickColumn -Grid $gridAutomap

$amBtnY = $amY
$chkAmOn = New-Object System.Windows.Forms.CheckBox
$chkAmOn.Text     = 'Automapping ON afterwards (untick to KEEP access but stop it opening in Outlook)'
$chkAmOn.Location = New-Object System.Drawing.Point(660, $amBtnY)
$chkAmOn.Size     = New-Object System.Drawing.Size(400, 40)
$chkAmOn.Checked  = $true
$tabAutomap.Controls.Add($chkAmOn)
$amBtnY += 44
$btnAmFix     = New-TabButton $tabAutomap 'Fix ticked' 660 $amBtnY 170
$amBtnY += 34
$btnAmRefresh = New-TabButton $tabAutomap 'Refresh list' 660 $amBtnY 170
$amBtnY += 40
$btnAmAll     = New-TabButton $tabAutomap 'Tick all' 660 $amBtnY 82
$btnAmNone    = New-TabButton $tabAutomap 'None'     748 $amBtnY 82
$amBtnY += 40
$amHint = New-TabLabel $tabAutomap '' 660 $amBtnY 400 -Muted
$amHint.Size = New-Object System.Drawing.Size(400, 70)
$amHint.Text = ('Each ticked person loses access for about 10 seconds while the grant is redone. ' +
                'Do not run it while they are working in the mailbox.' + "`r`n`r`n" +
                'Someone not listed? They have no DIRECT Full Access. Grant it on the Full Access tab -- automapping is set then.')
$amY += 208

$amStage = New-Object System.Windows.Forms.Label
$amStage.Tag       = 'NoTheme'
$amStage.Location  = New-Object System.Drawing.Point(8, $amY)
$amStage.Size      = New-Object System.Drawing.Size(800, 20)
$amStage.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$amStage.ForeColor = (Get-ToolStatusColour -Kind Muted)
$amStage.Text      = ''
$tabAutomap.Controls.Add($amStage)
$amBar = New-Object System.Windows.Forms.ProgressBar
$amBar.Location = New-Object System.Drawing.Point(820, $amY)
$amBar.Size     = New-Object System.Drawing.Size(238, 18)
$amBar.Visible  = $false
$tabAutomap.Controls.Add($amBar)
$amY += 26

$amOut = New-Object System.Windows.Forms.TextBox
$amOut.Multiline  = $true
$amOut.ReadOnly   = $true
$amOut.ScrollBars = 'Vertical'
$amOut.WordWrap   = $true
$amOut.Location   = New-Object System.Drawing.Point(8, $amY)
$amOut.Size       = New-Object System.Drawing.Size(1050, 270)
$amOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$amOut.Tag        = 'NoTheme'
$amOut.BackColor  = [System.Drawing.Color]::WhiteSmoke
$amOut.ForeColor  = [System.Drawing.Color]::Black
$tabAutomap.Controls.Add($amOut)

# TEN SECONDS, and the number is not arbitrary: long enough for the removal
# to settle before the grant is written, which is the whole point of doing
# them as two steps rather than one Set-. Carried over from Automap-Fix.ps1.
$script:AutomapWaitSeconds = 10

function Update-AutomapTab {
    $gridAutomap.Rows.Clear()
    $r = Get-MailboxAccessEntry -Smtp $script:Mbx.Smtp
    $full = @($r.Entries | Where-Object { $_.Kind -eq 'FullAccess' -and $_.Detail -ne 'DENY' })
    foreach ($e in $full) {
        $i = $gridAutomap.Rows.Add($false, $e.User, 'direct Full Access')
        $gridAutomap.Rows[$i].Tag = $e
    }
    $amStage.Text = ''
    $amStage.ForeColor = (Get-ToolStatusColour -Kind Muted)
    if ($full.Count -eq 0) {
        $amOut.Text = ('NOBODY HAS DIRECT FULL ACCESS ON THIS MAILBOX, so there is nothing here to fix.' + "`r`n`r`n" +
                       'If somebody expects it to appear in their Outlook, either they get access through a GROUP -- ' +
                       'which never automaps, and needs a direct grant or the mailbox added to their profile by hand -- ' +
                       'or they have no access at all. Grant it on the Full Access tab; automapping is set at that moment.')
    } else {
        $amOut.Text = ("$($full.Count) direct Full Access grant(s). Tick the people whose Outlook is not showing this mailbox and press Fix ticked." + "`r`n`r`n" +
                       'AUTOMAPPING CANNOT BE READ BACK. Exchange Online does not report it, so this tab cannot tell you whether it is on now -- only that the permission exists.')
    }
    if ("$($r.Note)".Trim()) { $amOut.Text = "$($r.Note)`r`n`r`n" + $amOut.Text; $script:TabReadFailed = $true }
}

function Set-AutomapStage {
    param([string]$Text, [ValidateSet('Muted', 'Warn', 'Bad', 'Good')][string]$Kind = 'Muted')
    $amStage.Text = $Text
    $amStage.ForeColor = (Get-ToolStatusColour -Kind $Kind)
    $amStage.Refresh()
}

function Invoke-AutomapFixOne {
    <#
        Remove, wait, re-add, verify -- for ONE trustee. Returns the lines
        for the output box and an Ok flag.

        THE REMOVE AND THE RE-ADD ARE ONE OPERATION. If the remove succeeds
        and the re-add fails, the person has LESS access than before this
        ran, and that is the loudest thing on the screen -- with the exact
        cmdlet to put it back. Failing at the remove is the safe failure:
        nothing changed.
    #>
    param([string]$Trustee, [bool]$AutoMap)
    $L = New-Object System.Collections.ArrayList
    $smtp = $script:Mbx.Smtp
    $r = [PSCustomObject]@{ Ok = $false; Removed = $false; Lines = @() }
    [void]$L.Add("== $Trustee")

    # ---- 1. remove -------------------------------------------------------
    Set-AutomapStage -Text "Removing Full Access for $Trustee ..." -Kind Warn
    try {
        Remove-MailboxAccessEntry -Smtp $smtp -Kind 'FullAccess' -Trustee $Trustee
        $r.Removed = $true
        [void]$L.Add('   1. REMOVED Full Access.')
    } catch {
        [void]$L.Add('   THE REMOVE FAILED, SO NOTHING WAS CHANGED. Their access is exactly as it was; it is safe to try again.')
        [void]$L.Add("      $($_.Exception.Message)")
        $r.Lines = @($L); return $r
    }

    # ---- 2. wait, counted down on screen ---------------------------------
    # A tab that sits still for ten seconds looks hung, and this is the exact
    # moment the person has no access -- the operator should see it is
    # deliberate and how long is left.
    $amBar.Maximum = $script:AutomapWaitSeconds; $amBar.Value = 0; $amBar.Visible = $true
    for ($sec = $script:AutomapWaitSeconds; $sec -ge 1; $sec--) {
        if ($script:Closing) { break }
        Set-AutomapStage -Text "Waiting $sec second(s) -- $Trustee has NO access right now ..." -Kind Warn
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Seconds 1
        $amBar.Value = [Math]::Min($amBar.Maximum, $amBar.Value + 1)
    }
    $amBar.Visible = $false
    [void]$L.Add("   2. Waited $($script:AutomapWaitSeconds) seconds.")

    # ---- 3. re-add, three tries -------------------------------------------
    Set-AutomapStage -Text "Re-adding Full Access for $Trustee ..." -Kind Warn
    $last = $null; $added = $false
    foreach ($attempt in 1..3) {
        try { Add-MailboxAccessEntry -Smtp $smtp -Kind 'FullAccess' -Trustee $Trustee -AutoMap $AutoMap; $added = $true; break }
        catch { $last = $_; if ($attempt -lt 3) { Start-Sleep -Seconds (4 * $attempt) } }
    }
    if (-not $added) {
        [void]$L.Add('')
        [void]$L.Add('   *** THE RE-ADD FAILED AFTER THREE TRIES. THIS PERSON NOW HAS NO ACCESS. ***')
        [void]$L.Add("      $($last.Exception.Message)")
        [void]$L.Add('   Put it back now, from the Full Access tab (Grant access...) or:')
        [void]$L.Add("      Add-MailboxPermission -Identity '$smtp' -User '$Trustee' -AccessRights FullAccess -InheritanceType All -AutoMapping `$$AutoMap")
        $r.Lines = @($L); return $r
    }
    [void]$L.Add("   3. RE-ADDED Full Access with automapping $(if ($AutoMap) { 'ON' } else { 'OFF' }).")

    # ---- 4. read back -----------------------------------------------------
    # 'No exception' is not evidence the grant landed, and this is the run
    # where being sure matters most.
    $after = Get-MailboxAccessEntry -Smtp $smtp
    $present = @($after.Entries | Where-Object { $_.Kind -eq 'FullAccess' -and ("$($_.User)" -eq $Trustee -or "$($_.User)" -like "*$Trustee*") }).Count -gt 0
    if ($present) { [void]$L.Add('   4. VERIFIED -- the permission is present again.') }
    else { [void]$L.Add('   4. The grant reported success but the read-back does not show it yet. That is usually replication; press Refresh list in a minute.') }
    $r.Ok = $true
    $r.Lines = @($L)
    return $r
}

$btnAmFix.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Fixing automapping')) { return }
        $picked = @(Get-GridPicked -Grid $gridAutomap)
        if ($picked.Count -eq 0) { Set-EditorStatus -Text 'Tick one or more people to fix.' -Bad; return }
        $who  = @($picked | ForEach-Object { "$($_.User)" } | Select-Object -Unique)
        $auto = [bool]$chkAmOn.Checked
        $a = Show-ToolDialog(
            ("Remove and re-add Full Access on $($script:Mbx.Display) for:" + "`r`n  " + ($who -join "`r`n  ") + "`r`n`r`n" +
             "Automapping afterwards: $(if ($auto) { 'ON -- it opens in their Outlook automatically' } else { 'OFF -- they keep access, it stops opening automatically' })" + "`r`n`r`n" +
             "Each person's access is REMOVED for about $($script:AutomapWaitSeconds) seconds and then granted again." + "`r`n" +
             'Do not do this while they are working in the mailbox.'),
            'Fix automapping', 'OKCancel', 'Warning', 'Button2')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $all = New-Object System.Collections.ArrayList
        $ok = 0; $bad = 0; $stranded = @()
        $btnAmFix.Enabled = $false; $btnAmRefresh.Enabled = $false
        try {
            foreach ($u in $who) {
                if ($script:Closing) { break }
                $res = Invoke-AutomapFixOne -Trustee $u -AutoMap $auto
                foreach ($l in $res.Lines) { [void]$all.Add($l) }
                [void]$all.Add('')
                $amOut.Text = ($all -join "`r`n")
                if ($res.Ok) { $ok++ } else { $bad++; if ($res.Removed) { $stranded += $u } }
            }
        } finally {
            $btnAmFix.Enabled = $true; $btnAmRefresh.Enabled = $true
            $amBar.Visible = $false
        }
        if ($stranded.Count) {
            Set-AutomapStage -Text ('ACCESS REMOVED AND NOT RESTORED for: ' + ($stranded -join ', ')) -Kind Bad
        } elseif ($bad) {
            Set-AutomapStage -Text "$ok done, $bad failed at the remove (nothing changed for those)." -Kind Bad
        } else {
            Set-AutomapStage -Text "Done for $ok -- Outlook picks it up on its next reconnect." -Kind Good
            [void]$all.Add('WHAT TO TELL THEM')
            [void]$all.Add('   Close Outlook completely and open it again. The mailbox appears when Outlook next')
            [void]$all.Add('   reconnects -- usually a few minutes, sometimes not until a restart. Watching the')
            [void]$all.Add('   folder list right now will show nothing.')
            [void]$all.Add('   If it still has not appeared after a restart and an hour, automapping is not the')
            [void]$all.Add('   problem -- check whether the mailbox was already added to their profile by hand,')
            [void]$all.Add('   which blocks the automatic one from showing.')
            $amOut.Text = ($all -join "`r`n")
        }
        Write-EditorAudit -Action 'MailboxAutomapFix' `
            -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
            -Detail ("Full Access removed, held $($script:AutomapWaitSeconds)s and re-added with automapping $auto for $($who -join '; ')" +
                     $(if ($stranded.Count) { "; REMOVED AND NOT RESTORED: $($stranded -join '; ')" } else { '' }))
        # The grid is refreshed from the server; the narrative above stays.
        $keep = $amOut.Text
        $r = Get-MailboxAccessEntry -Smtp $script:Mbx.Smtp
        $gridAutomap.Rows.Clear()
        foreach ($e in @($r.Entries | Where-Object { $_.Kind -eq 'FullAccess' -and $_.Detail -ne 'DENY' })) {
            $i = $gridAutomap.Rows.Add($false, $e.User, 'direct Full Access'); $gridAutomap.Rows[$i].Tag = $e
        }
        $amOut.Text = $keep
        $script:Loaded['Access'] = $false   # the Full Access tab's grid is stale now
        Set-EditorStatus -Text "Automapping fix: $ok done, $bad failed." -Bad:($bad -gt 0)
    } 'Fix automapping'
})
$btnAmRefresh.Add_Click({ if (Test-EditorLoaded) { Invoke-EditorAction { Update-AutomapTab } 'Refresh automap list' } })
$btnAmAll.Add_Click({ Set-GridPicked -Grid $gridAutomap -Value $true })
$btnAmNone.Add_Click({ Set-GridPicked -Grid $gridAutomap -Value $false })
#endregion

#region Folder Perm tab ---------------------------------------------------------
<#
    NOTHING TYPED HERE EITHER.

    Grant used to ask for the address in one prompt and the role name in
    another, against a list of eleven spellings. Ticking rows, choosing a role
    from a list and naming several mailboxes at once replaces all of it.

    ROLES ARE A CHOICE, NOT CHECKBOXES, and that is Exchange's doing rather
    than a UI preference: a folder permission stores ONE role per user per
    folder. Owner and Reviewer at the same time is not a thing that can be
    written, so offering two ticks would be offering a state the server cannot
    hold.
#>
$fpY = 8
[void](New-TabLabel $tabFolders 'FOLDER PERMISSIONS' 8 $fpY -Bold)
$fpY += 22
$fpNote = New-TabLabel $tabFolders '' 8 $fpY 1040 -Muted
$fpNote.Text = 'Send on Behalf is the other half of delegation and lives on the Full Access tab -- it is a mailbox property, not a folder right.'
$fpY += 24

[void](New-TabLabel $tabFolders 'Folder:' 8 ($fpY + 3) 60)
$cmbFolder               = New-Object System.Windows.Forms.ComboBox
$cmbFolder.DropDownStyle = 'DropDownList'
$cmbFolder.Location      = New-Object System.Drawing.Point(70, $fpY)
$cmbFolder.Size          = New-Object System.Drawing.Size(200, 24)
[void]$cmbFolder.Items.Add('All folders')
foreach ($f in $script:MailboxWellKnownFolder) { [void]$cmbFolder.Items.Add($f.Label) }
$cmbFolder.SelectedIndex = 0
$tabFolders.Controls.Add($cmbFolder)

$btnFolderRefresh = New-TabButton $tabFolders 'Refresh' 282 ($fpY - 2) 90
$fpY += 34

$gridFolder = New-TabGrid $tabFolders @(@('Folder', 110), @('User', 330), @('Rights', 220)) 8 $fpY 700 340
Add-GridPickColumn -Grid $gridFolder

$fpBtnY = $fpY
$btnFolderAdd    = New-TabButton $tabFolders 'Grant...'        720 $fpBtnY 170
$fpBtnY += 34
$btnFolderChange = New-TabButton $tabFolders 'Change ticked...' 720 $fpBtnY 170
$fpBtnY += 34
$btnFolderRemove = New-TabButton $tabFolders 'Remove ticked'    720 $fpBtnY 170
$fpBtnY += 34
$btnFolderAll    = New-TabButton $tabFolders 'Tick all'         720 $fpBtnY 82
$btnFolderNone   = New-TabButton $tabFolders 'None'             808 $fpBtnY 82
$fpBtnY += 42
$fpHint = New-TabLabel $tabFolders '' 720 $fpBtnY 320 -Muted
$fpHint.Size = New-Object System.Drawing.Size(320, 200)
$fpHint.Text = ('Default is what everyone in the organisation gets; Anonymous is what people outside it get. ' +
                'Both are shown in amber because changing them changes what EVERYBODY sees.' + "`r`n`r`n" +
                'Change and Remove act on every TICKED row.' + "`r`n`r`n" +
                'Reviewer reads. Editor reads and changes. Author adds and edits their own. ' +
                'AvailabilityOnly is free/busy with no detail, LimitedDetails adds subject and location.')
$fpY += 352

function Update-FoldersTab {
    $gridFolder.Rows.Clear()
    $m = $script:Mbx
    $wanted = @()
    if ($cmbFolder.SelectedIndex -le 0) {
        $wanted = @($script:MailboxWellKnownFolder)
    } else {
        $wanted = @($script:MailboxWellKnownFolder[$cmbFolder.SelectedIndex - 1])
    }

    $problems = @()
    foreach ($f in $wanted) {
        $r = Get-MailboxFolderAccess -Smtp $m.Smtp -FolderType $f.Type
        if (-not $r.Ok) {
            # A FOLDER THAT DOES NOT EXIST IS NOT AN ERROR. A room mailbox has
            # no Tasks folder worth speaking of, and a mailbox created in
            # another language may not resolve on the English name at all --
            # Get-MailboxFolderPath already tries the folder TYPE first. Report
            # it once, quietly, and carry on with the folders that did work.
            $problems += "$($f.Label): $($r.Reason)"
            continue
        }
        foreach ($e in @($r.Entries)) {
            # None carries no information and clutters the delegate view.
            if ("$($e.Rights)" -eq 'None') { continue }
            $i = $gridFolder.Rows.Add($false, $f.Type, $e.User, $e.Rights)
            $gridFolder.Rows[$i].Tag = $e
            if ("$($e.User)" -match '(?i)^(Default|Anonymous)$') {
                $gridFolder.Rows[$i].DefaultCellStyle.ForeColor = (Get-ToolStatusColour -Kind Warn)
            }
        }
    }
    if ($cmbFolder.SelectedIndex -le 0) {
        # Sorted by user, not folder: the delegate question is 'what does Jane
        # have', and that answer is only readable when Jane's rows are together.
        $gridFolder.Sort($gridFolder.Columns['User'], [System.ComponentModel.ListSortDirection]::Ascending)
    }
    $msg = "$($gridFolder.Rows.Count) permission(s)."
    if ($problems.Count) { $msg += '  Not read: ' + ($problems -join '; ') }
    # EVERY folder failing is a failed read, not a mailbox with no permissions.
    # One folder failing is ordinary -- a room mailbox has no useful Tasks
    # folder -- so only the total failure marks the tab as needing a retry.
    if ($problems.Count -ge @($wanted).Count) { $script:TabReadFailed = $true }
    Set-EditorStatus -Text $msg
}

function New-RoleList {
    # The roles, as a list you click. Reviewer is preselected because it is the
    # answer to 'let them see my calendar', which is most of these requests.
    param($Parent, [int]$X, [int]$Y, [int]$W = 200, [int]$H = 150)
    $l = New-Object System.Windows.Forms.ListBox
    $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.Size     = New-Object System.Drawing.Size($W, $H)
    foreach ($r in $script:MailboxFolderRole) { [void]$l.Items.Add($r) }
    $i = $l.Items.IndexOf('Reviewer')
    $l.SelectedIndex = $(if ($i -ge 0) { $i } else { 0 })
    $Parent.Controls.Add($l)
    return $l
}

function Show-FolderGrantDialog {
    <#
        Several mailboxes, several people, several folders, one role.

        ONE ROLE AND NOT ONE PER FOLDER, on purpose. 'Give these two Reviewer
        on the calendar and the inbox' is a request; 'give them Reviewer on the
        calendar and Editor on the inbox' is two requests, and building one
        dialog that does both would mean a role picker per folder and a screen
        nobody can read at a glance.
    #>
    param([string]$DefaultMailbox)

    $f = New-Object System.Windows.Forms.Form
    $f.Text          = 'Grant folder permission'
    $f.ClientSize    = New-Object System.Drawing.Size(740, 410)
    $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'
    $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $mk = {
        param($text, $x, $y, $w)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $text; $l.Location = New-Object System.Drawing.Point($x, $y); $l.Size = New-Object System.Drawing.Size($w, 18)
        $l.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
        $f.Controls.Add($l)
    }
    & $mk 'MAILBOXES -- one per line' 12 12 330
    $txtMbxList = New-DialogList $f 12 34 330 140 $DefaultMailbox
    & $mk 'PEOPLE -- one per line, or Default / Anonymous' 356 12 372
    $txtWhoList = New-DialogList $f 356 34 372 140 ''

    & $mk 'FOLDERS -- tick as many as apply' 12 186 330
    $y = 208
    $fchecks = @{}
    foreach ($k in $script:MailboxWellKnownFolder) {
        $fchecks[$k.Type] = New-DialogCheck $f $k.Label 20 $y 200 ($k.Type -eq 'Calendar')
        $y += 22
    }

    & $mk 'ROLE -- one, applied to every folder ticked' 356 186 372
    $lstRole = New-RoleList $f 356 208 200 120

    $note = New-Object System.Windows.Forms.Label
    $note.Text = 'Giving somebody a role they already hold on a folder replaces the old one -- it does not add a second entry.'
    $note.Location = New-Object System.Drawing.Point(12, 340); $note.Size = New-Object System.Drawing.Size(716, 30)
    $note.ForeColor = (Get-ToolStatusColour -Kind Muted)
    $f.Controls.Add($note)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Grant'; $ok.Location = New-Object System.Drawing.Point(528, 372); $ok.Size = New-Object System.Drawing.Size(96, 28)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK; $f.Controls.Add($ok)
    $ca = New-Object System.Windows.Forms.Button
    $ca.Text = 'Cancel'; $ca.Location = New-Object System.Drawing.Point(632, 372); $ca.Size = New-Object System.Drawing.Size(96, 28)
    $ca.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $f.Controls.Add($ca)
    $f.AcceptButton = $ok; $f.CancelButton = $ca

    Set-ConsoleTheme -Form $f
    $r = $f.ShowDialog()
    $res = $null
    if ($r -eq [System.Windows.Forms.DialogResult]::OK) {
        $folders = @()
        foreach ($k in $script:MailboxWellKnownFolder) { if ($fchecks[$k.Type].Checked) { $folders += $k.Type } }
        $res = [PSCustomObject]@{
            Mailboxes = (ConvertFrom-DialogList -Text $txtMbxList.Text -CompleteDomain)
            Trustees  = (ConvertFrom-DialogList -Text $txtWhoList.Text -CompleteDomain)
            Folders   = @($folders)
            Role      = "$($lstRole.SelectedItem)"
        }
    }
    $f.Dispose()
    return $res
}

function Show-FolderRoleDialog {
    # One role, for every ticked row. No typing, and no way to pick a
    # combination Exchange cannot store.
    param([string[]]$Rows)

    $f = New-Object System.Windows.Forms.Form
    $f.Text          = 'Change folder permission'
    $f.ClientSize    = New-Object System.Drawing.Size(560, 360)
    $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'
    $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $l0 = New-Object System.Windows.Forms.Label
    $l0.Text = "Set the role on $(@($Rows).Count) ticked permission(s):"
    $l0.Location = New-Object System.Drawing.Point(12, 12); $l0.Size = New-Object System.Drawing.Size(536, 18)
    $l0.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
    $f.Controls.Add($l0)

    $box = New-Object System.Windows.Forms.TextBox
    $box.Multiline = $true; $box.ReadOnly = $true; $box.ScrollBars = 'Vertical'; $box.WordWrap = $false
    $box.Location = New-Object System.Drawing.Point(12, 34); $box.Size = New-Object System.Drawing.Size(320, 260)
    $box.Font = New-Object System.Drawing.Font('Consolas', 9)
    $box.Text = (@($Rows) -join "`r`n")
    $box.Tag = 'NoTheme'; $box.BackColor = [System.Drawing.Color]::WhiteSmoke; $box.ForeColor = [System.Drawing.Color]::Black
    $f.Controls.Add($box)

    $l1 = New-Object System.Windows.Forms.Label
    $l1.Text = 'New role:'; $l1.Location = New-Object System.Drawing.Point(348, 34); $l1.Size = New-Object System.Drawing.Size(200, 18)
    $l1.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
    $f.Controls.Add($l1)
    $lstRole = New-RoleList $f 348 56 200 238

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Apply'; $ok.Location = New-Object System.Drawing.Point(348, 322); $ok.Size = New-Object System.Drawing.Size(96, 28)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK; $f.Controls.Add($ok)
    $ca = New-Object System.Windows.Forms.Button
    $ca.Text = 'Cancel'; $ca.Location = New-Object System.Drawing.Point(452, 322); $ca.Size = New-Object System.Drawing.Size(96, 28)
    $ca.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $f.Controls.Add($ca)
    $f.AcceptButton = $ok; $f.CancelButton = $ca

    Set-ConsoleTheme -Form $f
    $r = $f.ShowDialog()
    $role = $(if ($r -eq [System.Windows.Forms.DialogResult]::OK) { "$($lstRole.SelectedItem)" } else { '' })
    $f.Dispose()
    return $role
}

$btnFolderRefresh.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction { Update-FoldersTab } 'Read folder permissions'
})
$cmbFolder.Add_SelectedIndexChanged({
    if ($script:Busy) { return }
    if (-not $script:Mbx -or -not $script:Mbx.Ok) { return }
    Invoke-EditorAction { Update-FoldersTab } 'Read folder permissions'
})
$btnFolderAll.Add_Click({  Set-GridPicked -Grid $gridFolder -On $true })
$btnFolderNone.Add_Click({ Set-GridPicked -Grid $gridFolder -On $false })

$btnFolderAdd.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Granting a folder permission')) { return }
        $d = Show-FolderGrantDialog -DefaultMailbox $script:Mbx.Smtp
        if (-not $d) { return }
        if ($d.Mailboxes.Count -eq 0) { Set-EditorStatus -Text 'No mailboxes listed.' -Bad; return }
        if ($d.Trustees.Count -eq 0)  { Set-EditorStatus -Text 'No people listed.' -Bad; return }
        if ($d.Folders.Count -eq 0)   { Set-EditorStatus -Text 'Tick at least one folder.' -Bad; return }
        if (-not $d.Role)             { Set-EditorStatus -Text 'Pick a role.' -Bad; return }

        # DEFAULT AND ANONYMOUS CHANGE WHAT EVERYONE SEES, so they get their own
        # sentence rather than being folded into the count.
        $everyone = @($d.Trustees | Where-Object { $_ -match '(?i)^(Default|Anonymous)$' })
        $extra = ''
        if ($everyone.Count) {
            $extra = "`r`n`r`nTHIS INCLUDES $($everyone -join ' AND ') -- that is not one person, it is what EVERYBODY gets on those folders."
        }
        $total = $d.Mailboxes.Count * $d.Trustees.Count * $d.Folders.Count
        $a = Show-ToolDialog(
            ("Give $($d.Role) on $($d.Folders -join ', ') to $($d.Trustees.Count) person/people, on $($d.Mailboxes.Count) mailbox(es)?" +
             "`r`n`r`nMailboxes:`r`n  " + ($d.Mailboxes -join "`r`n  ") +
             "`r`n`r`nPeople:`r`n  " + ($d.Trustees -join "`r`n  ") +
             "`r`n`r`n$total permission(s) in total." + $extra),
            'Grant folder permission', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $lines = @(); $ok = 0; $bad = 0
        foreach ($mbx in $d.Mailboxes) {
            foreach ($who in $d.Trustees) {
                foreach ($ft in $d.Folders) {
                    try {
                        # Set-MailboxFolderAccess adds when there is nothing
                        # there and sets when there is -- so granting somebody a
                        # role they already hold replaces it rather than failing
                        # with 'an existing permission entry was found'.
                        Set-MailboxFolderAccess -Smtp $mbx -FolderType $ft -User $who -Role $d.Role
                        $lines += "OK      $mbx  $ft  <-  $who  ($($d.Role))"
                        $ok++
                        Write-CompletedLog -Action 'MailboxFolderPermissionSet' -Target (Get-LogTarget $mbx) `
                            -Result 'Success' -Environment 'Cloud' `
                            -Detail "$ft : $who -> $($d.Role) (Mailbox Editor, bulk)"
                    } catch {
                        $lines += "FAILED  $mbx  $ft  <-  $who`: $($_.Exception.Message)"
                        $bad++
                        Write-FailedLog -Action 'MailboxFolderPermissionSet' -Target (Get-LogTarget $mbx) `
                            -ErrorRecord $_ -Environment 'Cloud' -Detail "$ft : $who -> $($d.Role)"
                    }
                }
            }
        }
        Show-ResultListDialog -Title "Grant folder permission -- $ok done, $bad failed" -Lines $lines
        Update-FoldersTab
        Set-EditorStatus -Text "Granted $ok, failed $bad." -Bad:($bad -gt 0)
    } 'Grant folder permission'
})

$btnFolderChange.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Changing a folder permission')) { return }
        $picked = Get-GridPicked -Grid $gridFolder
        if ($picked.Count -eq 0) { Set-EditorStatus -Text 'Tick one or more rows to change.' -Bad; return }
        $shown = @($picked | ForEach-Object { "$($_.Folder)  $($_.User)  ($($_.Rights))" })
        $role = Show-FolderRoleDialog -Rows $shown
        if (-not $role) { return }

        $everyone = @($picked | Where-Object { "$($_.User)" -match '(?i)^(Default|Anonymous)$' })
        $extra = ''
        if ($everyone.Count) {
            $extra = "`r`n`r`nTHIS INCLUDES Default or Anonymous -- those rows are what EVERYBODY gets, not one person."
        }
        $a = Show-ToolDialog(
            ("Set $role on $($picked.Count) permission(s) on $($script:Mbx.Display)?" + "`r`n`r`n  " +
             ($shown -join "`r`n  ") + $extra),
            'Change folder permission', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $lines = @(); $ok = 0; $bad = 0
        foreach ($e in $picked) {
            try {
                Set-MailboxFolderAccess -Smtp $script:Mbx.Smtp -FolderType "$($e.Folder)" -User "$($e.User)" -Role $role
                $lines += "OK      $($e.Folder)  $($e.User)  $($e.Rights) -> $role"; $ok++
            } catch {
                $lines += "FAILED  $($e.Folder)  $($e.User): $($_.Exception.Message)"; $bad++
            }
        }
        Write-EditorAudit -Action 'MailboxFolderPermissionChange' `
            -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
            -Detail "set $role on: $(($picked | ForEach-Object { "$($_.Folder)/$($_.User)" }) -join '; ')"
        Show-ResultListDialog -Title "Change folder permission -- $ok done, $bad failed" -Lines $lines
        Update-FoldersTab
        Set-EditorStatus -Text "Changed $ok, failed $bad." -Bad:($bad -gt 0)
    } 'Change folder permission'
})

$btnFolderRemove.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Removing a folder permission')) { return }
        $picked = Get-GridPicked -Grid $gridFolder
        if ($picked.Count -eq 0) { Set-EditorStatus -Text 'Tick one or more rows to remove.' -Bad; return }

        # DEFAULT AND ANONYMOUS CANNOT BE REMOVED, only set to None. Attempting
        # it produces an Exchange error that reads like a permissions problem,
        # so they are separated out and handled as what they are.
        $special = @($picked | Where-Object { "$($_.User)" -match '(?i)^(Default|Anonymous)$' })
        $normal  = @($picked | Where-Object { "$($_.User)" -notmatch '(?i)^(Default|Anonymous)$' })
        $extra = ''
        if ($special.Count) {
            $extra = "`r`n`r`n$($special.Count) of these are Default or Anonymous. Those cannot be removed -- they will be set to None, which is the same thing in effect and is what EVERYBODY then gets."
        }
        $shown = @($picked | ForEach-Object { "  $($_.Folder)  $($_.User)  ($($_.Rights))" }) -join "`r`n"
        $a = Show-ToolDialog(("Remove $($picked.Count) folder permission(s) from $($script:Mbx.Display)?" + "`r`n`r`n" + $shown + $extra),
                             'Remove folder permission', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $lines = @(); $ok = 0; $bad = 0
        foreach ($e in $normal) {
            try { Remove-MailboxFolderAccess -Smtp $script:Mbx.Smtp -FolderType "$($e.Folder)" -User "$($e.User)"
                  $lines += "OK      removed $($e.Folder)  $($e.User)"; $ok++ }
            catch { $lines += "FAILED  $($e.Folder)  $($e.User): $($_.Exception.Message)"; $bad++ }
        }
        foreach ($e in $special) {
            try { Set-MailboxFolderAccess -Smtp $script:Mbx.Smtp -FolderType "$($e.Folder)" -User "$($e.User)" -Role 'None'
                  $lines += "OK      set None   $($e.Folder)  $($e.User)"; $ok++ }
            catch { $lines += "FAILED  $($e.Folder)  $($e.User): $($_.Exception.Message)"; $bad++ }
        }
        Write-EditorAudit -Action 'MailboxFolderPermissionRemove' `
            -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
            -Detail "$ok removed, $bad failed: $(($picked | ForEach-Object { "$($_.Folder)/$($_.User)" }) -join '; ')"
        if ($bad) { Show-ResultListDialog -Title "Remove folder permission -- $ok done, $bad failed" -Lines $lines }
        Update-FoldersTab
        Set-EditorStatus -Text "Removed $ok, failed $bad." -Bad:($bad -gt 0)
    } 'Remove folder permission'
})
#endregion


#region Mail Flow tab -----------------------------------------------------------
<#
    FORWARDING, DELIVERY RESTRICTIONS, SIZE AND RECIPIENT LIMITS, MAIL TIP.

    EVERY FIELD ON THIS TAB IS 'Unverified' IN THE ORIGIN TABLE, and that is
    not laziness. ForwardingAddress maps to altRecipient in Active Directory;
    the delivery restrictions map to authOrig, unauthOrig and
    dLMemSubmitPerms. Whether those flow from the directory depends on which
    attributes the sync is configured to write, which varies per deployment.
    ForwardingSmtpAddress has no directory counterpart and is the safer of the
    two forwarding fields.

    So on a synced mailbox these ask before writing and say plainly that the
    value needs re-checking after a sync cycle. On a cloud-only mailbox they
    are ordinary writes and nothing is asked.
#>
$mfY = 8
[void](New-TabLabel $tabMailFlow 'FORWARDING' 8 $mfY -Bold)
$mfY += 22
$mfFwdNote = New-TabLabel $tabMailFlow '' 8 $mfY 1050 -Muted
$mfFwdNote.Text = 'Forwarding to an address outside the organisation is how a compromised mailbox exfiltrates mail. Check it is expected before adding one.'
$mfY += 24

[void](New-TabLabel $tabMailFlow 'Forward to:' 8 ($mfY + 3) 110)
$txtFwd          = New-Object System.Windows.Forms.TextBox
$txtFwd.Location = New-Object System.Drawing.Point(124, $mfY)
$txtFwd.Size     = New-Object System.Drawing.Size(380, 24)
$txtFwd.Font     = New-Object System.Drawing.Font('Consolas', 9)
$tabMailFlow.Controls.Add($txtFwd)
$lblFwdKind = New-TabLabel $tabMailFlow '' 514 ($mfY + 3) 300 -Muted
$mfY += 28

$chkFwdKeep          = New-Object System.Windows.Forms.CheckBox
$chkFwdKeep.Text     = 'Keep a copy in this mailbox as well as forwarding'
$chkFwdKeep.Location = New-Object System.Drawing.Point(124, $mfY)
$chkFwdKeep.Size     = New-Object System.Drawing.Size(400, 20)
$tabMailFlow.Controls.Add($chkFwdKeep)
$mfY += 26

$btnFwdApply = New-TabButton $tabMailFlow 'Apply forwarding' 124 $mfY 150
$btnFwdClear = New-TabButton $tabMailFlow 'Stop forwarding'  284 $mfY 150
$mfY += 40

[void](New-TabLabel $tabMailFlow 'DELIVERY RESTRICTIONS' 8 $mfY -Bold)
$mfY += 22
$mfDelNote = New-TabLabel $tabMailFlow '' 8 $mfY 1050 -Muted
$mfDelNote.Text = 'Semicolon-separated addresses. Leave a box empty for no restriction. "Require authentication" blocks all mail from outside the organisation.'
$mfY += 24

[void](New-TabLabel $tabMailFlow 'Accept only from:' 8 ($mfY + 3) 110)
$txtAccept          = New-Object System.Windows.Forms.TextBox
$txtAccept.Location = New-Object System.Drawing.Point(124, $mfY)
$txtAccept.Size     = New-Object System.Drawing.Size(660, 24)
$txtAccept.Font     = New-Object System.Drawing.Font('Consolas', 9)
$tabMailFlow.Controls.Add($txtAccept)
$mfY += 28
[void](New-TabLabel $tabMailFlow 'Reject from:' 8 ($mfY + 3) 110)
$txtReject          = New-Object System.Windows.Forms.TextBox
$txtReject.Location = New-Object System.Drawing.Point(124, $mfY)
$txtReject.Size     = New-Object System.Drawing.Size(660, 24)
$txtReject.Font     = New-Object System.Drawing.Font('Consolas', 9)
$tabMailFlow.Controls.Add($txtReject)
$mfY += 28
$chkAuthOnly          = New-Object System.Windows.Forms.CheckBox
$chkAuthOnly.Text     = 'Require senders to be authenticated (blocks all external mail)'
$chkAuthOnly.Location = New-Object System.Drawing.Point(124, $mfY)
$chkAuthOnly.Size     = New-Object System.Drawing.Size(460, 20)
$tabMailFlow.Controls.Add($chkAuthOnly)
$mfY += 26
$btnDelApply = New-TabButton $tabMailFlow 'Apply restrictions' 124 $mfY 150
$mfY += 40

[void](New-TabLabel $tabMailFlow 'LIMITS AND MAIL TIP' 8 $mfY -Bold)
$mfY += 22
[void](New-TabLabel $tabMailFlow 'Max send size:' 8 ($mfY + 3) 110)
$txtMaxSend          = New-Object System.Windows.Forms.TextBox
$txtMaxSend.Location = New-Object System.Drawing.Point(124, $mfY)
$txtMaxSend.Size     = New-Object System.Drawing.Size(150, 24)
$tabMailFlow.Controls.Add($txtMaxSend)
[void](New-TabLabel $tabMailFlow 'Max receive size:' 290 ($mfY + 3) 120)
$txtMaxRecv          = New-Object System.Windows.Forms.TextBox
$txtMaxRecv.Location = New-Object System.Drawing.Point(414, $mfY)
$txtMaxRecv.Size     = New-Object System.Drawing.Size(150, 24)
$tabMailFlow.Controls.Add($txtMaxRecv)
[void](New-TabLabel $tabMailFlow 'Recipient limit:' 580 ($mfY + 3) 110)
$txtRcpLimit          = New-Object System.Windows.Forms.TextBox
$txtRcpLimit.Location = New-Object System.Drawing.Point(694, $mfY)
$txtRcpLimit.Size     = New-Object System.Drawing.Size(90, 24)
$tabMailFlow.Controls.Add($txtRcpLimit)
$mfY += 26
$mfLimNote = New-TabLabel $tabMailFlow '' 124 $mfY 900 -Muted
$mfLimNote.Text = 'Sizes as 35MB or 36700160. Recipient limit is per message. Leave any box as it is to keep the current value; the word unlimited is accepted.'
$mfY += 24
[void](New-TabLabel $tabMailFlow 'Mail tip:' 8 ($mfY + 3) 110)
$txtMailTip          = New-Object System.Windows.Forms.TextBox
$txtMailTip.Location = New-Object System.Drawing.Point(124, $mfY)
$txtMailTip.Size     = New-Object System.Drawing.Size(660, 24)
$tabMailFlow.Controls.Add($txtMailTip)
$mfY += 26
$mfTipNote = New-TabLabel $tabMailFlow '' 124 $mfY 900 -Muted
$mfTipNote.Text = 'Shown in Outlook above the message when someone addresses mail here. Empty removes it.'
$mfY += 24
$btnLimApply = New-TabButton $tabMailFlow 'Apply limits and tip' 124 $mfY 170

function ConvertTo-AddressListText {
    # The restriction properties come back as recipient objects, not strings.
    param($Value)
    $out = @()
    foreach ($v in @($Value)) {
        if (-not "$v") { continue }
        $s = "$v"
        # A distinguished name is unreadable in a text box; take the last leaf.
        if ($s -match '^CN=([^,]+),') { $s = $Matches[1] }
        $out += $s
    }
    return ($out -join '; ')
}

function ConvertFrom-AddressListText {
    param([string]$Text)
    $t = "$Text".Trim()
    if (-not $t) { return @() }
    return @($t -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Update-MailFlowTab {
    $mb = $script:Mbx.Raw
    # ForwardingSmtpAddress wins in the box because it is the one that works
    # for an external address; ForwardingAddress is shown when that is what is
    # set, with the label saying which.
    $fwdSmtp = "$($mb.ForwardingSmtpAddress)" -replace '^smtp:', ''
    $fwdInt  = "$($mb.ForwardingAddress)"
    if ($fwdSmtp) {
        $txtFwd.Text = $fwdSmtp
        $lblFwdKind.Text = 'ForwardingSmtpAddress (any address)'
    } elseif ($fwdInt) {
        # THE BOX IS LEFT EMPTY ON PURPOSE. ForwardingAddress is a recipient,
        # and it renders as a name or a distinguished name -- neither of which
        # this tab will accept back, so filling the box with it offers an Apply
        # that is guaranteed to fail. The value goes in the label instead,
        # where it reads as information rather than as an edit.
        $txtFwd.Text = ''
        $lblFwdKind.Text = "ForwardingAddress (internal): $(ConvertTo-AddressListText $fwdInt) -- type an address to replace it"

    } else {
        $txtFwd.Text = ''
        $lblFwdKind.Text = 'not forwarding'
    }
    $chkFwdKeep.Checked = [bool]$mb.DeliverToMailboxAndForward

    # THE RENDERED TEXT IS REMEMBERED IN .Tag. ConvertTo-AddressListText is
    # lossy -- a distinguished name comes back as its CN leaf -- so writing
    # these lists back unchanged does not round-trip: it can resolve to a
    # different recipient or fail outright. Apply compares against the Tag and
    # sends a list ONLY when the operator actually edited that box.
    $txtAccept.Text   = (ConvertTo-AddressListText $mb.AcceptMessagesOnlyFromSendersOrMembers)
    $txtAccept.Tag    = $txtAccept.Text
    $txtReject.Text   = (ConvertTo-AddressListText $mb.RejectMessagesFromSendersOrMembers)
    $txtReject.Tag    = $txtReject.Text
    $chkAuthOnly.Checked = [bool]$mb.RequireSenderAuthenticationEnabled

    $txtMaxSend.Text  = "$($mb.MaxSendSize)"
    $txtMaxRecv.Text  = "$($mb.MaxReceiveSize)"
    $txtRcpLimit.Text = "$($mb.RecipientLimits)"
    $txtMailTip.Text  = "$($mb.MailTip)"
}

$btnFwdApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Setting forwarding')) { return }
        if (-not (Test-EditorFieldWritable -Field 'ForwardingSmtpAddress')) { return }
        $to = "$($txtFwd.Text)".Trim()
        if (-not $to) { Set-EditorStatus -Text 'Enter an address to forward to, or use Stop forwarding.' -Bad; return }
        if ($to -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { Set-EditorStatus -Text 'That does not look like an email address.' -Bad; return }

        # EXTERNAL FORWARDING GETS ITS OWN WARNING. Auto-forwarding out of the
        # organisation is the standard exfiltration route from a compromised
        # mailbox, and the operator setting it should have to confirm they
        # meant it rather than have it happen on one keystroke.
        $ourDomains = @()
        try { $ourDomains = @((Get-TenantConfig).PrimaryDomain, (Get-TenantConfig).OnMicrosoftDomain) | Where-Object { $_ } } catch { }
        $isExternal = $true
        foreach ($d in $ourDomains) { if ($to -imatch ([regex]::Escape("@$d") + '$')) { $isExternal = $false; break } }
        $warn = ''
        if ($isExternal) {
            $warn = "`r`n`r`nTHIS ADDRESS LOOKS EXTERNAL. Auto-forwarding to an outside address is the usual way mail leaves a compromised mailbox. Only set this if you know it was asked for."
        }
        $keep = $(if ($chkFwdKeep.Checked) { 'A copy stays in this mailbox.' } else { 'NO copy is kept -- mail leaves and does not stay here.' })
        $a = Show-ToolDialog(
            ("Forward $($script:Mbx.Display)'s mail to $to?" + "`r`n`r`n" + $keep + $warn),
            'Set forwarding', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
            Identity                   = $script:Mbx.Smtp
            ForwardingSmtpAddress      = $to
            # ForwardingAddress IS CLEARED IN THE SAME WRITE. The two
            # properties are independent: setting the SMTP one on a mailbox
            # that already forwards to an internal recipient leaves both set,
            # and mail keeps going to the old target as well. Stop forwarding
            # already clears both; this makes the two paths agree.
            ForwardingAddress          = $null
            DeliverToMailboxAndForward = [bool]$chkFwdKeep.Checked
            ErrorAction                = 'Stop'
        } | Out-Null
        Write-EditorAudit -Action 'MailboxForwardingSet' -Detail "Forwarding to $to; keep a copy = $($chkFwdKeep.Checked)$(if ($isExternal) { '; EXTERNAL address' })."
        $script:Mbx = Resolve-MailboxTarget -Smtp $script:Mbx.Smtp
        Update-MailFlowTab
        Set-EditorStatus -Text "Forwarding to $to."
    } 'Set forwarding'
})

$btnFwdClear.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Stopping forwarding')) { return }
        if (-not (Test-EditorFieldWritable -Field 'ForwardingSmtpAddress')) { return }
        # BOTH properties are cleared. Clearing only the one that happened to
        # be shown leaves the other still forwarding, which looks exactly like
        # the tool having done nothing.
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
            Identity              = $script:Mbx.Smtp
            ForwardingSmtpAddress = $null
            ForwardingAddress     = $null
            ErrorAction           = 'Stop'
        } | Out-Null
        Write-EditorAudit -Action 'MailboxForwardingClear' -Detail 'Forwarding removed (both ForwardingSmtpAddress and ForwardingAddress).'
        $script:Mbx = Resolve-MailboxTarget -Smtp $script:Mbx.Smtp
        Update-MailFlowTab
        Set-EditorStatus -Text 'Forwarding removed.'
    } 'Stop forwarding'
})

$btnDelApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Delivery restrictions')) { return }
        if (-not (Test-EditorFieldWritable -Field 'AcceptMessagesOnlyFrom')) { return }
        $accept = ConvertFrom-AddressListText $txtAccept.Text
        $reject = ConvertFrom-AddressListText $txtReject.Text
        $p = @{
            Identity                           = $script:Mbx.Smtp
            RequireSenderAuthenticationEnabled = [bool]$chkAuthOnly.Checked
            ErrorAction                        = 'Stop'
        }
        # An EMPTY box means no restriction, and that has to be written as an
        # empty collection rather than skipped -- skipping it would silently
        # keep a restriction the operator has just cleared. But a box that was
        # NOT edited is not written at all, because the text in it is a lossy
        # rendering of recipient objects and would not round-trip.
        $acceptEdited = ("$($txtAccept.Text)".Trim() -ne "$($txtAccept.Tag)".Trim())
        $rejectEdited = ("$($txtReject.Text)".Trim() -ne "$($txtReject.Tag)".Trim())
        if ($acceptEdited) { $p['AcceptMessagesOnlyFromSendersOrMembers'] = @($accept) }
        if ($rejectEdited) { $p['RejectMessagesFromSendersOrMembers']     = @($reject) }
        if ($acceptEdited -and $accept.Count -gt 0) {
            $a = Show-ToolDialog(
                ("Restrict $($script:Mbx.Display) to accept mail ONLY from:" + "`r`n`r`n  " + ($accept -join "`r`n  ") +
                 "`r`n`r`nEverything from anyone else will be rejected."),
                'Delivery restriction', 'OKCancel', 'Warning')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        }
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters $p | Out-Null
        Write-EditorAudit -Action 'MailboxDeliveryRestriction' `
            -Detail "Accept only from: $(if (-not $acceptEdited) { '(unchanged)' } elseif ($accept) { $accept -join '; ' } else { '(none)' }); reject from: $(if (-not $rejectEdited) { '(unchanged)' } elseif ($reject) { $reject -join '; ' } else { '(none)' }); require auth = $($chkAuthOnly.Checked)."
        $script:Mbx = Resolve-MailboxTarget -Smtp $script:Mbx.Smtp
        Update-MailFlowTab
        Set-EditorStatus -Text 'Delivery restrictions applied.'
    } 'Apply delivery restrictions'
})

$btnLimApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Limits and the mail tip')) { return }
        $p = @{ Identity = $script:Mbx.Smtp; ErrorAction = 'Stop' }
        $changed = @()
        $mb = $script:Mbx.Raw
        # Only send what actually changed. Writing every field every time
        # makes the audit line claim six changes when the operator made one.
        if ("$($txtMaxSend.Text)".Trim() -and "$($txtMaxSend.Text)".Trim() -ne "$($mb.MaxSendSize)") {
            $p['MaxSendSize'] = "$($txtMaxSend.Text)".Trim(); $changed += "MaxSendSize=$($p['MaxSendSize'])"
        }
        if ("$($txtMaxRecv.Text)".Trim() -and "$($txtMaxRecv.Text)".Trim() -ne "$($mb.MaxReceiveSize)") {
            $p['MaxReceiveSize'] = "$($txtMaxRecv.Text)".Trim(); $changed += "MaxReceiveSize=$($p['MaxReceiveSize'])"
        }
        if ("$($txtRcpLimit.Text)".Trim() -and "$($txtRcpLimit.Text)".Trim() -ne "$($mb.RecipientLimits)") {
            $p['RecipientLimits'] = "$($txtRcpLimit.Text)".Trim(); $changed += "RecipientLimits=$($p['RecipientLimits'])"
        }
        if ("$($txtMailTip.Text)" -ne "$($mb.MailTip)") {
            $p['MailTip'] = "$($txtMailTip.Text)"; $changed += "MailTip"
        }
        if ($changed.Count -eq 0) { Set-EditorStatus -Text 'Nothing changed.'; return }
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters $p | Out-Null
        Write-EditorAudit -Action 'MailboxLimits' -Detail ($changed -join '; ')
        $script:Mbx = Resolve-MailboxTarget -Smtp $script:Mbx.Smtp
        Update-MailFlowTab
        Set-EditorStatus -Text "Applied: $($changed -join '; ')"
    } 'Apply limits and mail tip'
})
#endregion

#region Settings tab ------------------------------------------------------------
<#
    THIS TAB IS GONE, AND SAYING SO IS THE POINT.

    Everything it held -- mailbox type, hide from the address book, the online
    archive -- moved to Overview on 2026-09-18. All three are things somebody
    changes about a mailbox without needing to know anything else about it,
    and making people find a second tab for them was exactly the friction this
    window exists to remove.

    THE TAB ITSELF IS NOT CREATED any more. A tab left behind holding a
    sentence about where things went is a tab people still click, and a
    read-only echo of a control that lives elsewhere is two places for one
    fact to disagree -- which is the rule this file follows everywhere else.
#>
#endregion

#region Retention tab -----------------------------------------------------------
<#
    LITIGATION HOLD AND RETENTION POLICY.

    Read and set, which is what Hold & Retention did, plus the read-only
    Retention-Policy view folded in. Both are properties of the mailbox in
    Exchange Online, so neither is affected by directory sync.
#>
$rtY = 8
[void](New-TabLabel $tabRetention 'LITIGATION HOLD' 8 $rtY -Bold)
$rtY += 22
$rtHoldNow = New-TabLabel $tabRetention '' 8 $rtY 900 -Muted
$rtY += 24
$chkHold          = New-Object System.Windows.Forms.CheckBox
$chkHold.Text     = 'Litigation hold enabled'
$chkHold.Location = New-Object System.Drawing.Point(8, $rtY)
$chkHold.Size     = New-Object System.Drawing.Size(220, 20)
$tabRetention.Controls.Add($chkHold)
[void](New-TabLabel $tabRetention 'Duration (days, blank = forever):' 240 ($rtY + 1) 210)
$txtHoldDays          = New-Object System.Windows.Forms.TextBox
$txtHoldDays.Location = New-Object System.Drawing.Point(456, ($rtY - 2))
$txtHoldDays.Size     = New-Object System.Drawing.Size(90, 24)
$tabRetention.Controls.Add($txtHoldDays)
$btnHoldApply = New-TabButton $tabRetention 'Apply hold' 560 ($rtY - 5) 120
$rtY += 28
$rtHoldNote = New-TabLabel $tabRetention '' 8 $rtY 1050 -Muted
$rtHoldNote.Size = New-Object System.Drawing.Size(1050, 50)
$rtHoldNote.Text = ('While a hold is on, nothing the user deletes actually leaves the mailbox, and the mailbox cannot be deleted.' + "`r`n" +
                    'Removing a hold can let previously retained items age out permanently. Check with whoever asked for it before turning one off.')
$rtY += 60

[void](New-TabLabel $tabRetention 'RETENTION POLICY' 8 $rtY -Bold)
$rtY += 22
$rtPolNow = New-TabLabel $tabRetention '' 8 $rtY 900 -Muted
$rtY += 24
[void](New-TabLabel $tabRetention 'Policy:' 8 ($rtY + 3) 60)
$cmbPolicy               = New-Object System.Windows.Forms.ComboBox
$cmbPolicy.DropDownStyle = 'DropDownList'
$cmbPolicy.Location      = New-Object System.Drawing.Point(70, $rtY)
$cmbPolicy.Size          = New-Object System.Drawing.Size(340, 24)
$tabRetention.Controls.Add($cmbPolicy)
$btnPolicyApply = New-TabButton $tabRetention 'Apply policy' 420 ($rtY - 2) 120
$rtY += 30
$rtPolNote = New-TabLabel $tabRetention '' 8 $rtY 1050 -Muted
$rtPolNote.Size = New-Object System.Drawing.Size(1050, 50)
$rtPolNote.Text = ('A retention policy decides what is archived or deleted and when. The managed folder assistant applies it, which is not immediate.' + "`r`n" +
                    'This list is every policy in the organisation; picking the wrong one can start deleting mail on a schedule.')

function Update-RetentionTab {
    $mb = $script:Mbx.Raw
    $on = [bool]$mb.LitigationHoldEnabled
    $chkHold.Checked = $on
    $dur = "$($mb.LitigationHoldDuration)"
    $txtHoldDays.Text = $(if ($dur -and $dur -notmatch '(?i)unlimited') { ($dur -split '\.')[0] } else { '' })
    $rtHoldNow.Text = $(if ($on) {
        "Currently ON$(if ($dur) { " -- duration $dur" })$(if ("$($mb.LitigationHoldDate)") { "; set $($mb.LitigationHoldDate)" })$(if ("$($mb.LitigationHoldOwner)") { " by $($mb.LitigationHoldOwner)" })"
    } else { 'Currently OFF' })

    $rtPolNow.Text = "Currently: $(if ("$($mb.RetentionPolicy)") { "$($mb.RetentionPolicy)" } else { '(none assigned)' })"
    if ($cmbPolicy.Items.Count -eq 0) {
        try {
            foreach ($p in @(Invoke-EXOCommand -Name 'Get-RetentionPolicy' -Parameters @{ ErrorAction = 'Stop' } | Sort-Object Name)) {
                [void]$cmbPolicy.Items.Add("$($p.Name)")
            }
        } catch {
            $script:TabReadFailed = $true
            $rtPolNote.Text = "The list of retention policies could not be read: $($_.Exception.Message)"
        }
    }
    $cur = "$($mb.RetentionPolicy)"
    if ($cur -and $cmbPolicy.Items.Contains($cur)) { $cmbPolicy.SelectedItem = $cur }
}

$btnHoldApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Litigation hold')) { return }
        if (-not (Test-EditorFieldWritable -Field 'LitigationHoldEnabled')) { return }
        $on = [bool]$chkHold.Checked
        $days = "$($txtHoldDays.Text)".Trim()
        if ($days -and $days -notmatch '^\d+$') { Set-EditorStatus -Text 'Duration must be a whole number of days, or blank for forever.' -Bad; return }

        $was = [bool]$script:Mbx.Raw.LitigationHoldEnabled
        $msg = $(if ($on) {
            "Put $($script:Mbx.Display) on litigation hold$(if ($days) { " for $days days" } else { ' indefinitely' })?" +
            "`r`n`r`nNothing the user deletes will leave the mailbox, and the mailbox cannot be deleted while it is held."
        } else {
            "REMOVE the litigation hold from $($script:Mbx.Display)?" +
            "`r`n`r`nItems that were being retained only because of this hold can then age out and be permanently lost." +
            "`r`n`r`nIf this hold was placed for a legal matter, confirm with whoever asked for it first."
        })
        $a = Show-ToolDialog($msg, $(if ($on) { 'Enable litigation hold' } else { 'REMOVE litigation hold' }), 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $p = @{ Identity = $script:Mbx.Smtp; LitigationHoldEnabled = $on; ErrorAction = 'Stop' }
        if ($on -and $days) { $p['LitigationHoldDuration'] = [int]$days }
        if ($on -and -not $days) { $p['LitigationHoldDuration'] = 'Unlimited' }
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters $p | Out-Null
        Write-EditorAudit -Action 'MailboxLitigationHold' -Detail "Hold was $was, now $on$(if ($on) { "; duration $(if ($days) { "$days days" } else { 'unlimited' })" })."
        $script:Mbx = Resolve-MailboxTarget -Smtp $script:Mbx.Smtp
        Update-RetentionTab
        Set-EditorStatus -Text $(if ($on) { 'Litigation hold is on.' } else { 'Litigation hold removed.' })
    } 'Apply litigation hold'
})

$btnPolicyApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'The retention policy')) { return }
        if (-not (Test-EditorFieldWritable -Field 'RetentionPolicy')) { return }
        $pol = "$($cmbPolicy.SelectedItem)"
        if (-not $pol) { Set-EditorStatus -Text 'Pick a policy first.' -Bad; return }
        if ($pol -eq "$($script:Mbx.Raw.RetentionPolicy)") { Set-EditorStatus -Text "$pol is already assigned."; return }
        $a = Show-ToolDialog(
            ("Assign the retention policy '$pol' to $($script:Mbx.Display)?" + "`r`n`r`n" +
             'A retention policy archives or deletes mail on a schedule. The wrong one starts removing mail the user expects to keep.'),
            'Assign retention policy', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
            Identity = $script:Mbx.Smtp; RetentionPolicy = $pol; ErrorAction = 'Stop'
        } | Out-Null
        Write-EditorAudit -Action 'MailboxRetentionPolicy' -Detail "Retention policy set to $pol."
        $script:Mbx = Resolve-MailboxTarget -Smtp $script:Mbx.Smtp
        Update-RetentionTab
        Set-EditorStatus -Text "Retention policy is now $pol."
    } 'Assign retention policy'
})
#endregion

#region Auto Replies tab --------------------------------------------------------
<#
    THE REAL EDITOR, IN A TAB. Not a summary with a button any more.

    This is the same Show-OutOfOfficeForm the standalone Out of Office tool
    opens -- same toolbar, same fonts and colours, same templates, same
    preview, same HTML. It is built ONTO this tab page rather than into a
    window of its own, so there is exactly one out-of-office editor in the
    toolkit and no second place for its behaviour to drift.

    This was tried on 2026-09-18 and taken back out the same day, because a
    plain PowerShell scriptblock is not a closure -- every handler in that
    function resolved against THIS file when it fired, where $form and $tabs
    exist and are the wrong objects. Pressing Close inside the tab closed the
    Mailbox Editor. The function has since been rewritten around one $ui state
    bag with .GetNewClosure() on every handler, which is what makes hosting it
    safe; the long note at the top of that function explains why both halves
    were needed.

    WHAT THIS TAB OWNS: the target, and when the first load happens. A
    WebBrowser's ActiveX is not created until the control has been VISIBLE, so
    FirstRun cannot be called at build time -- it runs the first time the
    operator actually looks at this tab, which is what Invoke-EditorTabLoad
    already gives us.
#>
$script:Oof      = $null
$script:OofFirst = $false
try {
    $script:Oof = Show-OutOfOfficeForm -Container $tabReplies -NoTargetBar
} catch {
    # A FAILURE HERE MUST NOT TAKE THE WINDOW WITH IT. Twelve other tabs work
    # without this one, and an editor that will not build is a tab with a
    # sentence in it, not a dead launch.
    $oofDead = New-TabLabel $tabReplies '' 8 8 1040
    $oofDead.Size = New-Object System.Drawing.Size(1040, 80)
    $oofDead.ForeColor = (Get-ToolStatusColour -Kind Bad)
    $oofDead.Text = ("The automatic replies editor could not be built: $($_.Exception.Message)" + "`r`n`r`n" +
                     'Every other tab still works. Out of Office is also a tool of its own in the console.')
}

function Update-RepliesTab {
    if (-not $script:Oof) { return }
    # POINT IT AT THIS MAILBOX FIRST, EVERY TIME. The editor keeps its own
    # target, and the operator can load a different mailbox into this window
    # without the Replies tab being touched in between.
    # METHOD CALL SYNTAX, NOT '& scriptblock'. These are ScriptMethods on the
    # object the editor returns, which is how they reach their own state
    # without a closure -- see the note at the top of Show-OutOfOfficeForm.
    $script:Oof.SetTarget($script:Mbx.Smtp)
    try {
        if (-not $script:OofFirst) {
            # ONCE. FirstRun flips to the external page and back to force both
            # WebBrowser controls to create their ActiveX, which only happens
            # once a control has been shown. Doing that on every visit would
            # throw away unsaved edits.
            $script:OofFirst = $true
            $script:Oof.FirstRun()
        } else {
            $script:Oof.Load()
        }
    } catch {
        $script:TabReadFailed = $true
        Set-EditorStatus -Text "Automatic replies could not be read: $($_.Exception.Message)" -Bad
    }
}

# THE TAB CACHE IS CLEARED WHEN THE MAILBOX CHANGES, so Invoke-EditorTabLoad
# calls Update-RepliesTab again and the editor is re-pointed and re-read. The
# FirstRun flag is deliberately NOT reset -- the ActiveX only needs creating
# once per window, not once per mailbox.
#endregion

#region Recovery tab ------------------------------------------------------------
<#
    RECOVERABLE ITEMS.

    WHAT THIS DOES AND DOES NOT DO. It reports how much is sitting in the
    Recoverable Items folder and how long deleted items are kept, and it can
    lengthen that retention. It does NOT restore individual messages: putting
    items back is a per-item operation the user does from Outlook's Recover
    Deleted Items, or a compliance search for anything larger. A button here
    claiming to 'recover' would be promising something this interface cannot
    deliver.

    The useful admin action is the one that buys time: if a user has just
    deleted something important, extending the retention window stops it
    ageing out while they go and find it.
#>
$rcY = 8
[void](New-TabLabel $tabRecovery 'RECOVERABLE ITEMS' 8 $rcY -Bold)
$rcY += 22
$rcNote = New-TabLabel $tabRecovery '' 8 $rcY 1050 -Muted
$rcNote.Size = New-Object System.Drawing.Size(1050, 50)
$rcNote.Text = ('This is the hidden folder deleted mail goes to. The user recovers individual items from Outlook: Folder tab, Recover Deleted Items.' + "`r`n" +
                'What an administrator can usefully do here is lengthen the window so nothing ages out while it is being looked for.')
$rcY += 58

[void](New-TabLabel $tabRecovery 'Recoverable items size:' 8 $rcY 160)
$rcSize = New-TabValue $tabRecovery 176 $rcY 260
[void](New-TabLabel $tabRecovery 'Items:' 460 $rcY 60)
$rcCount = New-TabValue $tabRecovery 524 $rcY 160
$rcY += 26
[void](New-TabLabel $tabRecovery 'Quota for that folder:' 8 $rcY 160)
$rcQuota = New-TabValue $tabRecovery 176 $rcY 260
[void](New-TabLabel $tabRecovery 'Single item recovery:' 460 $rcY 150)
$rcSir = New-TabValue $tabRecovery 614 $rcY 160
$rcY += 34

[void](New-TabLabel $tabRecovery 'Keep deleted items for (days):' 8 ($rcY + 3) 200)
$txtRetainDays          = New-Object System.Windows.Forms.TextBox
$txtRetainDays.Location = New-Object System.Drawing.Point(214, $rcY)
$txtRetainDays.Size     = New-Object System.Drawing.Size(90, 24)
$tabRecovery.Controls.Add($txtRetainDays)
$btnRetainApply = New-TabButton $tabRecovery 'Apply' 314 ($rcY - 2) 100
$rcY += 30
$rcRetNote = New-TabLabel $tabRecovery '' 8 $rcY 1050 -Muted
$rcRetNote.Size = New-Object System.Drawing.Size(1050, 50)
$rcRetNote.Text = ('The tenant maximum is 30 days unless a hold is in place, in which case items are kept regardless of this number.' + "`r`n" +
                    'Shortening this can permanently destroy items that were still recoverable. Lengthening it is the safe direction.')
$rcY += 58
$btnSirEnable = New-TabButton $tabRecovery 'Turn on single item recovery' 8 $rcY 230
$rcY += 34
$rcSirNote = New-TabLabel $tabRecovery '' 8 $rcY 1050 -Muted
$rcSirNote.Size = New-Object System.Drawing.Size(1050, 34)
$rcSirNote.Text = 'Single item recovery stops a user purging items from Recoverable Items themselves, so deleted mail survives until the window expires.'

function Update-RecoveryTab {
    $mb = $script:Mbx.Raw
    $rcSir.Text   = $(if ($mb.SingleItemRecoveryEnabled) { 'ON' } else { 'off' })
    $btnSirEnable.Enabled = -not [bool]$mb.SingleItemRecoveryEnabled
    $dur = "$($mb.RetainDeletedItemsFor)"
    $txtRetainDays.Text = $(if ($dur) { ($dur -split '\.')[0] } else { '' })
    $rcSize.Text = '(reading...)'; $rcCount.Text = ''; $rcQuota.Text = "$($mb.RecoverableItemsQuota)"
    $rcSize.Refresh()
    try {
        # THE FOLDER SCOPE MATTERS. Without -FolderScope RecoverableItems this
        # returns the visible mailbox size, which is not what this tab is
        # about and would read as reassuringly small.
        $st = @(Invoke-EXOCommand -Name 'Get-MailboxFolderStatistics' -Parameters @{
                    Identity = $script:Mbx.Smtp; FolderScope = 'RecoverableItems'; ErrorAction = 'Stop' })
        $root = @($st | Where-Object { "$($_.FolderType)" -eq 'RecoverableItemsRoot' }) | Select-Object -First 1
        if ($root) {
            $rcSize.Text  = "$($root.FolderAndSubfolderSize)"
            $rcCount.Text = "$($root.ItemsInFolderAndSubfolders)"
        } else {
            $rcSize.Text  = '(no recoverable items folder reported)'
        }
    } catch {
        $script:TabReadFailed = $true
        $rcSize.Text = "(unavailable: $($_.Exception.Message))"
    }
}

$btnRetainApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Deleted-item retention')) { return }
        if (-not (Test-EditorFieldWritable -Field 'RetainDeletedItemsFor')) { return }
        $d = "$($txtRetainDays.Text)".Trim()
        if ($d -notmatch '^\d+$') { Set-EditorStatus -Text 'Enter a whole number of days.' -Bad; return }
        $n = [int]$d
        $was = "$($script:Mbx.Raw.RetainDeletedItemsFor)"
        $wasDays = 0; [void][int]::TryParse((($was -split '\.')[0]), [ref]$wasDays)
        if ($n -lt $wasDays) {
            $a = Show-ToolDialog(
                ("Shorten the recovery window from $wasDays to $n days?" + "`r`n`r`n" +
                 'Items older than the new window can be purged permanently and are not recoverable afterwards.'),
                'Shortening the window', 'OKCancel', 'Warning')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        }
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
            Identity = $script:Mbx.Smtp; RetainDeletedItemsFor = $n; ErrorAction = 'Stop'
        } | Out-Null
        Write-EditorAudit -Action 'MailboxRetainDeleted' -Detail "Deleted-item retention changed from $wasDays to $n days."
        $script:Mbx = Resolve-MailboxTarget -Smtp $script:Mbx.Smtp
        Update-RecoveryTab
        Set-EditorStatus -Text "Deleted items are now kept for $n days."
    } 'Apply deleted-item retention'
})

$btnSirEnable.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Single item recovery')) { return }
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
            Identity = $script:Mbx.Smtp; SingleItemRecoveryEnabled = $true; ErrorAction = 'Stop'
        } | Out-Null
        Write-EditorAudit -Action 'MailboxSingleItemRecovery' -Detail 'Single item recovery turned on.'
        $script:Mbx = Resolve-MailboxTarget -Smtp $script:Mbx.Smtp
        Update-RecoveryTab
        Set-EditorStatus -Text 'Single item recovery is on.'
    } 'Turn on single item recovery'
})
#endregion

#region Health tab --------------------------------------------------------------
<#
    WHAT IS WRONG WITH THIS MAILBOX -- AND ONLY WHAT NO OTHER TAB SHOWS.

    The old Mailbox Health tool reported size, forwarding, delegates, holds and
    the archive alongside everything here. All five of those now have a tab of
    their own in this window, and repeating them would mean two places showing
    the same fact and eventually disagreeing about it.

    So this tab carries the four things nothing else does:

      PROTOCOL ACCESS   Two rows. Outlook on the web, MAPI and EWS are what
                        people use, and a finding when they are OFF.
                        ActiveSync, POP and IMAP should be off, and are a
                        finding only when they are ON. The answer to "it works
                        in the browser but not on the phone", and the first
                        thing a compromise assessment looks at.
      ACCOUNT STATE     Whether the directory account is disabled. A disabled
                        account with a user mailbox is a leaver nobody
                        converted; a disabled account with a shared mailbox is
                        correct and normal.
      HYBRID GUIDS      ExchangeGuid and ArchiveGuid, cloud against
                        on-premises. When these drift nothing errors and mail
                        keeps flowing -- until a move fails or the archive
                        stops being reachable, months later.
      INBOX RULES       How many, and which of them forward, redirect or
                        delete. That combination is the standard footprint of
                        a compromised mailbox.

    THE VERDICT AT THE TOP DOES cover the other tabs -- but it states the
    FINDING and names the tab, rather than repeating the value. "Forwarding is
    set to an external address -- see Mail Flow" tells you what you need
    without becoming a second copy of Mail Flow.
#>
$hlY = 8
[void](New-TabLabel $tabHealth 'FINDINGS' 8 $hlY -Bold)
$hlY += 22
$hlVerdict            = New-Object System.Windows.Forms.TextBox
$hlVerdict.Multiline  = $true
$hlVerdict.ReadOnly   = $true
$hlVerdict.ScrollBars = 'Vertical'
$hlVerdict.WordWrap   = $true
$hlVerdict.Location   = New-Object System.Drawing.Point(8, $hlY)
$hlVerdict.Size       = New-Object System.Drawing.Size(1040, 130)
$hlVerdict.Font       = New-Object System.Drawing.Font('Consolas', 9)
$hlVerdict.Tag        = 'NoTheme'
$hlVerdict.BackColor  = [System.Drawing.Color]::WhiteSmoke
$hlVerdict.ForeColor  = [System.Drawing.Color]::Black
$tabHealth.Controls.Add($hlVerdict)
$hlY += 140

# ORDER IS DELIBERATE, AND IT IS NOT ALPHABETICAL OR THE CMDLET'S ORDER.
#
# Top row is what people actually use and what ought to be on: Outlook on the
# web, MAPI, EWS. Bottom row is what ought to be OFF: ActiveSync, POP, IMAP.
# Read down the second row and anything saying Enabled is a question to ask.
# THEY ARE EDITABLE NOW, AND THEY COST NO EXTRA ROOM.
#
# These six were read with Get-CASMailbox and rendered into read-only boxes;
# Set-CASMailbox appeared nowhere in this file. That was not a hybrid rule or
# a safety rule -- the write was simply never built, so the tab could tell you
# POP was on and leave you to go and turn it off somewhere else.
#
# COMBO BOXES IN PLACE OF THE TEXT BOXES, at the same coordinates, so the
# whole thing fits without a new row: this tab had about fifty pixels left.
# The Apply sits on the heading row for the same reason.
#
# ALL SIX PARAMETERS ARE Applicable TO EXCHANGE ONLINE as well as on-premises,
# and the tab is cloud-only anyway, so the write goes through Exchange Online
# exactly like the read it is paired with.
function New-ProtoCombo {
    param($Parent, [int]$X, [int]$Y, [int]$W = 120)
    $c = New-Object System.Windows.Forms.ComboBox
    $c.DropDownStyle = 'DropDownList'
    $c.Location      = New-Object System.Drawing.Point($X, ($Y - 3))
    $c.Size          = New-Object System.Drawing.Size($W, 24)
    # THE BLANK ENTRY IS 'NOT READ', and it is a real state: a mailbox whose
    # CAS settings could not be fetched must not show as 'Disabled'.
    foreach ($v in @('', 'Enabled', 'Disabled')) { [void]$c.Items.Add($v) }
    $c.SelectedIndex = 0
    $Parent.Controls.Add($c)
    return $c
}

[void](New-TabLabel $tabHealth 'PROTOCOL ACCESS' 8 $hlY -Bold)
$btnProtoApply = New-TabButton $tabHealth 'Apply protocols' 900 ($hlY - 4) 150
$hlY += 22
[void](New-TabLabel $tabHealth 'Outlook on the web:' 8 $hlY 150)
$hlOwa   = New-ProtoCombo $tabHealth 170 $hlY
[void](New-TabLabel $tabHealth 'MAPI (Outlook):' 320 $hlY 130)
$hlMapi  = New-ProtoCombo $tabHealth 460 $hlY
[void](New-TabLabel $tabHealth 'EWS:' 600 $hlY 120)
$hlEws   = New-ProtoCombo $tabHealth 730 $hlY
$hlY += 26
[void](New-TabLabel $tabHealth 'ActiveSync:' 8 $hlY 150)
$hlEas   = New-ProtoCombo $tabHealth 170 $hlY
[void](New-TabLabel $tabHealth 'POP:' 320 $hlY 130)
$hlPop   = New-ProtoCombo $tabHealth 460 $hlY
[void](New-TabLabel $tabHealth 'IMAP:' 600 $hlY 120)
$hlImap  = New-ProtoCombo $tabHealth 730 $hlY
$hlY += 26
$hlProtoNote = New-TabLabel $tabHealth '' 8 $hlY 1040 -Muted
$hlProtoNote.Text = 'The second row -- ActiveSync, POP, IMAP -- should be Disabled. Enabled is the finding, not disabled. Change any of the six and press Apply protocols.'
$hlY += 28

[void](New-TabLabel $tabHealth 'ACCOUNT AND HYBRID' 8 $hlY -Bold)
$hlY += 22
[void](New-TabLabel $tabHealth 'Directory account:' 8 $hlY 150)
$hlAcct  = New-TabValue $tabHealth 170 $hlY 360
$hlY += 24
[void](New-TabLabel $tabHealth 'ExchangeGuid:' 8 $hlY 150)
$hlEgOk  = New-TabValue $tabHealth 170 $hlY 560
$hlY += 24
[void](New-TabLabel $tabHealth 'ArchiveGuid:' 8 $hlY 150)
$hlAgOk  = New-TabValue $tabHealth 170 $hlY 560
$hlY += 24
$hlGuidNote = New-TabLabel $tabHealth '' 8 $hlY 1040 -Muted
$hlY += 28

[void](New-TabLabel $tabHealth 'INBOX RULES -- rules that send mail OUT of the tenant' 8 $hlY -Bold)
$hlY += 22
[void](New-TabLabel $tabHealth 'Rules:' 8 $hlY 150)
$hlRules = New-TabValue $tabHealth 170 $hlY 360
$hlY += 24
$hlRuleList            = New-Object System.Windows.Forms.TextBox
$hlRuleList.Multiline  = $true
$hlRuleList.ReadOnly   = $true
$hlRuleList.ScrollBars = 'Vertical'
$hlRuleList.Location   = New-Object System.Drawing.Point(8, $hlY)
$hlRuleList.Size       = New-Object System.Drawing.Size(1040, 70)
$hlRuleList.Font       = New-Object System.Drawing.Font('Consolas', 9)
$hlRuleList.Tag        = 'NoTheme'
$hlRuleList.BackColor  = [System.Drawing.Color]::WhiteSmoke
$hlRuleList.ForeColor  = [System.Drawing.Color]::Black
$tabHealth.Controls.Add($hlRuleList)
$hlY += 78
$btnHlRefresh = New-TabButton $tabHealth 'Re-check' 8 $hlY 130

function Format-HealthFlag {
    # A protocol that is ON is not automatically a finding; a protocol that is
    # OFF usually IS the answer to somebody's ticket, so both are stated
    # plainly rather than left as True/False for the reader to interpret.
    #
    # THE PROTOCOL ROWS NO LONGER USE THIS -- they are editable combo boxes now
    # and go through Set-ProtoState. It stays for the read-only flags below
    # them, which are still text.
    param($Value)
    if ($null -eq $Value) { return '(unknown)' }
    if ([bool]$Value) { return 'enabled' }
    return 'DISABLED'
}

function Set-ProtoState {
    # A combo filled from what the service said. $null -- the property was not
    # returned at all -- lands on the blank entry rather than on 'Disabled',
    # because 'we could not see it' and 'it is off' lead to different actions.
    param($Combo, $Value)
    $Combo.SelectedItem = $(if ($null -eq $Value) { '' } elseif ([bool]$Value) { 'Enabled' } else { 'Disabled' })
}

function Get-ProtoState {
    # $true / $false / $null, back out of a combo. $null means 'leave it alone'.
    param($Combo)
    switch ("$($Combo.SelectedItem)") {
        'Enabled'  { return $true }
        'Disabled' { return $false }
        default    { return $null }
    }
}

$btnProtoApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Changing protocol access')) { return }
        if (-not $script:CasWas) {
            Set-EditorStatus -Text 'The protocol settings have not been read yet -- open or refresh the Health tab first.' -Bad; return
        }

        # ONLY WHAT CHANGED, compared against what the combos were filled from.
        $p = @{ Identity = $script:Mbx.Smtp; Confirm = $false; ErrorAction = 'Stop' }
        $changed = @(); $risky = @()
        foreach ($f in @(
            @{ Name = 'OWAEnabled';        Combo = $hlOwa;  Label = 'Outlook on the web'; OffIsRisk = $true }
            @{ Name = 'MAPIEnabled';       Combo = $hlMapi; Label = 'MAPI (Outlook)';     OffIsRisk = $true }
            @{ Name = 'EwsEnabled';        Combo = $hlEws;  Label = 'EWS';                OffIsRisk = $true }
            @{ Name = 'ActiveSyncEnabled'; Combo = $hlEas;  Label = 'ActiveSync';         OffIsRisk = $false }
            @{ Name = 'PopEnabled';        Combo = $hlPop;  Label = 'POP';                OffIsRisk = $false }
            @{ Name = 'ImapEnabled';       Combo = $hlImap; Label = 'IMAP';               OffIsRisk = $false }
        )) {
            $want = Get-ProtoState $f.Combo
            # A BLANK COMBO MEANS 'LEAVE IT ALONE', not 'set it to false'.
            if ($null -eq $want) { continue }
            $was = $script:CasWas.($f.Name)
            if ($null -ne $was -and [bool]$was -eq $want) { continue }
            $p[$f.Name] = $want
            $changed += "$($f.Label): $(if ($want) { 'Enabled' } else { 'Disabled' })"
            # THE TWO DIRECTIONS ARE NOT EQUALLY DANGEROUS, and the
            # confirmation says which is which. Turning MAPI off stops Outlook
            # working for that person; turning POP on re-opens a protocol the
            # tenant probably disabled on purpose.
            if ($f.OffIsRisk -and -not $want) { $risky += "Turning $($f.Label) OFF will stop it working for this person." }
            if (-not $f.OffIsRisk -and $want) { $risky += "Turning $($f.Label) ON re-enables a legacy protocol -- these are usually off deliberately." }
        }
        if ($changed.Count -eq 0) { Set-EditorStatus -Text 'Nothing changed.'; return }

        $extra = $(if ($risky.Count) { "`r`n`r`n" + ($risky -join "`r`n") } else { '' })
        $a = Show-ToolDialog(("On $($script:Mbx.Display):" + "`r`n`r`n  " + ($changed -join "`r`n  ") + $extra),
                             'Apply protocol access', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        Invoke-EXOCommand -Name 'Set-CASMailbox' -Parameters $p | Out-Null
        Write-EditorAudit -Action 'MailboxProtocols' -Detail ($changed -join '; ')
        Update-HealthTab
        Set-EditorStatus -Text "Applied: $($changed -join '; ')."
    } 'Apply protocol access'
})

function Update-HealthTab {
    foreach ($t in @($hlOwa, $hlEas, $hlMapi, $hlPop, $hlImap, $hlEws)) { $t.SelectedItem = '' }
    $script:CasWas = $null
    $hlVerdict.Text = 'Checking...'
    $hlVerdict.Refresh()

    $m        = $script:Mbx
    $mb       = $m.Raw
    $findings = New-Object System.Collections.ArrayList
    $notes    = New-Object System.Collections.ArrayList

    # --- protocols -------------------------------------------------------
    try {
        $cas = Invoke-EXOCommand -Name 'Get-CASMailbox' -Parameters @{ Identity = $m.Smtp; ErrorAction = 'Stop' }
        Set-ProtoState $hlOwa  $cas.OWAEnabled
        Set-ProtoState $hlEas  $cas.ActiveSyncEnabled
        Set-ProtoState $hlMapi $cas.MAPIEnabled
        Set-ProtoState $hlPop  $cas.PopEnabled
        Set-ProtoState $hlImap $cas.ImapEnabled
        Set-ProtoState $hlEws  $cas.EwsEnabled
        # WHAT THE COMBOS WERE FILLED FROM, kept so Apply sends only what the
        # operator actually changed. Set-CASMailbox with all six every time
        # would overwrite a protocol somebody else changed between this tab
        # loading and the button being pressed.
        $script:CasWas = [PSCustomObject]@{
            OWAEnabled = $cas.OWAEnabled; ActiveSyncEnabled = $cas.ActiveSyncEnabled
            MAPIEnabled = $cas.MAPIEnabled; PopEnabled = $cas.PopEnabled
            ImapEnabled = $cas.ImapEnabled; EwsEnabled = $cas.EwsEnabled
        }
        # THE TOP ROW IS A FINDING WHEN IT IS OFF. These are how people work;
        # off means somebody cannot get to their mail.
        if (-not $cas.OWAEnabled)  { [void]$findings.Add('Outlook on the web is disabled.') }
        if (-not $cas.MAPIEnabled) { [void]$findings.Add('MAPI is disabled -- Outlook cannot connect at all.') }
        if (-not $cas.EwsEnabled)  { [void]$findings.Add('EWS is disabled -- free/busy lookups, some add-ins and several migration tools stop working.') }

        # THE BOTTOM ROW IS A FINDING WHEN IT IS ON, AND SILENT WHEN IT IS OFF.
        #
        # This was backwards. The tab used to report ActiveSync being disabled
        # as a problem, which meant the DESIRED state produced a finding and
        # the findings box was never empty on a correctly configured mailbox.
        # A box that always has something in it stops being read, and then the
        # real finding underneath goes unread with it.
        #
        # ActiveSync, POP and IMAP should be off. Off is correct and says
        # nothing; on is the thing worth a sentence.
        if ($cas.ActiveSyncEnabled) {
            [void]$findings.Add('ActiveSync is ENABLED. It should be off unless this mailbox has a specific need for it -- it is a standing route into the mailbox from any phone that has ever been paired.')
        }
        if ($cas.PopEnabled -or $cas.ImapEnabled) {
            $which = @(); if ($cas.PopEnabled) { $which += 'POP' }; if ($cas.ImapEnabled) { $which += 'IMAP' }
            [void]$findings.Add("$($which -join ' and ') is ENABLED. Legacy, poor support for modern authentication, and a common footing in mailbox compromises. Turn it off unless something named is using it.")
        }
    } catch {
        $script:TabReadFailed = $true
        # BLANK, NOT '(unavailable)'. These are ComboBoxes now, and setting
        # .Text on a DropDownList that has no such item silently does nothing
        # -- it would have left the PREVIOUS mailbox's protocol states on
        # screen under a failed read, which is worse than saying nothing.
        # $script:CasWas stays $null, so Apply refuses rather than writing
        # against a baseline it never got.
        foreach ($t in @($hlOwa, $hlEas, $hlMapi, $hlPop, $hlImap, $hlEws)) { $t.SelectedItem = '' }
        [void]$notes.Add("Protocol settings could not be read: $($_.Exception.Message)")
    }

    # --- account state ---------------------------------------------------
    $hlAcct.Text = '(unknown)'
    try {
        $u = Invoke-EXOCommand -Name 'Get-User' -Parameters @{ Identity = $m.Smtp; ErrorAction = 'Stop' }
        $disabled = [bool]$u.AccountDisabled
        $hlAcct.Text = $(if ($disabled) { 'DISABLED' } else { 'enabled' })
        $type = "$($mb.RecipientTypeDetails)"
        if ($disabled -and $type -imatch 'UserMailbox') {
            [void]$findings.Add('The account is DISABLED but this is still a user mailbox. That is what a leaver looks like when nobody converted the mailbox to shared -- and it is still consuming a licence.')
        }
        if (-not $disabled -and $type -imatch 'SharedMailbox') {
            [void]$findings.Add('Shared mailbox with an ENABLED account -- somebody can sign in to it directly. Shared mailboxes are normally left disabled.')
        }
    } catch {
        if ("$($m.Type)" -match '(?i)Shared|Room|Equipment') {
            $hlAcct.Text = 'no user object (normal for this mailbox type)'
        } else {
            [void]$notes.Add("The directory account state could not be read: $($_.Exception.Message)")
        }
    }

    # --- hybrid GUIDs ----------------------------------------------------
    #
    # SKIPPED ENTIRELY WITHOUT ON-PREMISES EXCHANGE. A cloud-only session
    # cannot answer this, and two blank fields read as "they do not match".
    $hlEgOk.Text = '(not checked)'; $hlAgOk.Text = '(not checked)'
    $ready = Test-OnPremMailboxReady
    if (-not $ready.Ok) {
        $hlGuidNote.Text = "Hybrid GUID check skipped -- $($ready.Reason)"
        $hlGuidNote.ForeColor = (Get-ToolStatusColour -Kind Muted)
    } else {
        try {
            $stub = Invoke-OnPremCommand -Name 'Get-RemoteMailbox' -Parameters @{ Identity = $m.Smtp; ErrorAction = 'Stop' }
            $cEg = "$($mb.ExchangeGuid)"; $oEg = "$($stub.ExchangeGuid)"
            $cAg = "$($mb.ArchiveGuid)";  $oAg = "$($stub.ArchiveGuid)"
            $zero = '00000000-0000-0000-0000-000000000000'

            if ($cEg -ieq $oEg) { $hlEgOk.Text = "match ($cEg)" }
            else {
                $hlEgOk.Text = "MISMATCH -- cloud $cEg, on-prem $oEg"
                [void]$findings.Add('ExchangeGuid does not match between the cloud mailbox and the on-premises stub. Nothing will error today, but the mailbox cannot be moved and a later migration fails with a message that names neither GUID.')
            }

            if (($cAg -eq $zero -or -not $cAg) -and ($oAg -eq $zero -or -not $oAg)) {
                $hlAgOk.Text = 'no archive on either side'
            } elseif ($cAg -ieq $oAg) {
                $hlAgOk.Text = "match ($cAg)"
            } else {
                $hlAgOk.Text = "MISMATCH -- cloud $cAg, on-prem $oAg"
                [void]$findings.Add('ArchiveGuid does not match. The archive stops being reachable from Outlook when these drift.')
            }
            $hlGuidNote.Text = 'Compared against the on-premises remote mailbox stub.'
            $hlGuidNote.ForeColor = (Get-ToolStatusColour -Kind Muted)
        } catch {
            # NOT-FOUND IS AN ANSWER; ANYTHING ELSE IS A FAILED QUESTION.
            # Treating every exception as 'there is no stub' turned a dropped
            # session or an access-denied into the hard finding that a synced
            # mailbox has nowhere to route its mail -- a false alarm on a
            # mailbox that is perfectly fine.
            $notFound = ("$($_.CategoryInfo.Category)" -eq 'ObjectNotFound') -or
                        ($_.Exception.Message -match "(?i)couldn't be found|not found")
            if ($notFound) {
                $hlEgOk.Text = '(no on-premises stub)'; $hlAgOk.Text = '(no on-premises stub)'
                $hlGuidNote.Text = 'No remote mailbox stub on-premises -- normal for a cloud-only mailbox, a finding for a synced one.'
                $hlGuidNote.ForeColor = (Get-ToolStatusColour -Kind Muted)
                if ($m.IsSynced) {
                    [void]$findings.Add('This mailbox is synced from Active Directory but has NO on-premises remote mailbox stub. On-premises mail has nowhere to route.')
                }
            } else {
                $hlEgOk.Text = '(not checked)'; $hlAgOk.Text = '(not checked)'
                $hlGuidNote.Text = 'The on-premises stub could not be read -- see the findings box.'
                $hlGuidNote.ForeColor = (Get-ToolStatusColour -Kind Warn)
                [void]$notes.Add("The on-premises remote mailbox could not be read, so the hybrid GUIDs were not compared: $($_.Exception.Message)")
            }
        }
    }

    # --- inbox rules -----------------------------------------------------
    $hlRules.Text = '(reading...)'; $hlRuleList.Text = ''
    try {
        $ir = Get-InboxRulesForMailbox -Smtp $m.Smtp
        if ($ir.Error) {
            $hlRules.Text = '(unavailable)'
            [void]$notes.Add("Inbox rules could not be read: $($ir.Error)")
        } else {
            $hlRules.Text = "$(@($ir.Rules).Count)"

            # ONLY MAIL LEAVING THE ORGANISATION IS A FINDING.
            #
            # This used to flag every rule that forwarded, redirected OR
            # deleted. Deleting is what inbox rules are FOR -- filing and
            # binning newsletters is the normal use -- and forwarding to a
            # colleague is a person covering their own desk. Flagging those
            # meant the list was full of ordinary housekeeping, which is how a
            # real finding gets lost.
            #
            # What matters is mail going OUT of the tenant. "External" is
            # measured against the accepted-domain list, which is what makes
            # it a real test rather than a guess at what a domain looks like.
            $accepted = @(Get-AcceptedDomainList)
            $blind    = ($accepted.Count -eq 0)

            $sus = @(@($ir.Rules) | Where-Object {
                $_.Enabled -and ($_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo)
            } | ForEach-Object {
                $r = $_
                $out = @()
                foreach ($to in @($r.ForwardTo) + @($r.ForwardAsAttachmentTo) + @($r.RedirectTo)) {
                    if (-not $to) { continue }
                    $dom = (Get-SmtpDomain -Address "$to").ToLower()
                    # NO DOMAIN MEANS INTERNAL, NOT UNKNOWN. A recipient that
                    # came back as a display name or a legacy DN resolved
                    # inside the directory -- that is what having no SMTP
                    # domain on it means.
                    if (-not $dom) { continue }
                    if ($blind -or ($accepted -notcontains $dom)) { $out += $dom }
                }
                if (-not $out) { return }
                $acts = @()
                if ($r.ForwardTo -or $r.ForwardAsAttachmentTo) { $acts += 'Forward' }
                if ($r.RedirectTo) { $acts += 'Redirect' }
                # THE DOMAIN IS NAMED. "A rule forwards externally" sends
                # somebody back to Outlook to find out where; naming it is the
                # difference between a finding and an errand.
                "$($r.Name)  [$($acts -join ', ') -> $(($out | Select-Object -Unique) -join ', ')]"
            })
            if ($sus.Count) {
                $hlRuleList.Text = ($sus -join "`r`n")
                if ($blind) {
                    # THE FALLBACK SAYS WHICH ONE IT IS. Without the accepted
                    # domain list this cannot tell inside from outside, so it
                    # shows every forward and says so rather than presenting a
                    # guess as a finding.
                    [void]$findings.Add("$($sus.Count) enabled rule(s) forward or redirect mail. The accepted-domain list could not be read, so these could NOT be filtered to external ones -- some of them are probably internal and harmless.")
                } else {
                    [void]$findings.Add("$($sus.Count) enabled rule(s) forward or redirect mail OUT OF THE TENANT. Mail leaving the organisation by inbox rule is the classic compromise footprint, especially paired with a delete -- check who created these and when.")
                }
            } else {
                # SAYS WHAT WAS ACTUALLY CHECKED. "No rules found" would be
                # false on a mailbox full of ordinary filing rules.
                $hlRuleList.Text = 'No enabled rule sends mail outside the organisation.'
            }
        }
    } catch {
        $hlRules.Text = '(unavailable)'
        [void]$notes.Add("Inbox rules could not be read: $($_.Exception.Message)")
    }

    # --- findings that belong to OTHER tabs ------------------------------
    #
    # STATED, NOT REPEATED. The finding says what is wrong and which tab holds
    # the detail; the values themselves stay in one place.
    if ($script:Licence -and $script:Licence.Known) {
        $t = "$($mb.RecipientTypeDetails)"
        if ($t -imatch 'SharedMailbox' -and $script:Licence.Licensed) {
            [void]$findings.Add('Shared mailbox IS licensed -- a shared mailbox under 50GB does not need one. See Overview.')
        }
        if ($t -imatch 'UserMailbox' -and -not $script:Licence.Licensed) {
            [void]$findings.Add('User mailbox is NOT licensed. Mail stops after the grace period. See Overview.')
        }
    }
    if ("$($mb.ForwardingSmtpAddress)" -or "$($mb.ForwardingAddress)") {
        [void]$findings.Add('This mailbox forwards its mail. See Mail Flow for where to, and whether a copy is kept.')
    }
    if ($mb.LitigationHoldEnabled) {
        [void]$findings.Add('Litigation hold is on -- nothing in this mailbox can be permanently deleted. See Retention.')
    }
    if ($mb.HiddenFromAddressListsEnabled) {
        # 'See Settings' WAS A DEAD POINTER. The Settings tab was removed and
        # Hide moved to Overview; the finding kept sending people to a tab
        # that is not there.
        [void]$findings.Add('Hidden from the address list. Mail sent directly still delivers. See Overview.')
    }

    $sb = New-Object System.Text.StringBuilder
    if ($findings.Count -eq 0) {
        [void]$sb.AppendLine('Nothing found. Protocols, account state, hybrid GUIDs and inbox rules all look ordinary.')
    } else {
        foreach ($f in $findings) { [void]$sb.AppendLine("* $f") }
    }
    if ($notes.Count) {
        [void]$sb.AppendLine('')
        foreach ($n in $notes) { [void]$sb.AppendLine("(could not check) $n") }
    }
    $hlVerdict.Text = $sb.ToString()
}

$btnHlRefresh.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'The health check')) { return }
        Update-HealthTab
        Set-EditorStatus -Text 'Health re-checked.'
    } 'Re-check health'
})
#endregion

#region MBX Stats tab -----------------------------------------------------------
<#
    WHERE THE SPACE WENT.

    NOTHING HERE RUNS WHEN THE TAB IS SELECTED. Get-MailboxFolderStatistics
    enumerates every folder in the mailbox, and on a large one that is tens of
    thousands of rows -- seconds at best, a hung-looking window at worst. Every
    other tab in this window loads on sight; this one loads on a button, so
    that opening Statistics by accident costs nothing and no other tab ever
    waits behind it.

    TOTAL SIZE AND QUOTA ARE NOT REPEATED HERE. They are on Overview. What this
    adds is the breakdown -- which folder holds it, which is almost always the
    actual answer: Deleted Items nobody emptied, a Sent Items going back nine
    years, or Recoverable Items quietly holding everything a retention policy
    was asked to keep.
#>
$stsY = 8
$stsNote = New-TabLabel $tabStats '' 8 $stsY 1040 -Muted
$stsNote.Size = New-Object System.Drawing.Size(1040, 34)
$stsNote.Text = ('Per-folder sizes are the slowest read in this window, so they are not fetched until you ask.' + "`r`n" +
                 'Recoverable Items is not a folder the user can see, and is the usual reason a mailbox is bigger than the sum of everything visible.')
$stsY += 40
$btnStsLoad = New-TabButton $tabStats 'Load folder statistics' 8 $stsY 190
$btnStsArch = New-TabButton $tabStats 'Include the archive' 208 $stsY 160
$stsY += 36
$stsSummary = New-TabLabel $tabStats '' 8 $stsY 1040 -Muted
$stsSummary.Size = New-Object System.Drawing.Size(1040, 34)
$stsSummary.Text = 'Not loaded.'
$stsY += 40
$gridStats = New-TabGrid $tabStats @(@('Folder', 420), @('Items', 100), @('Size', 160), @('Where', 120)) 8 $stsY 880 300

function Add-StatsRows {
    <#
        One folder list into the grid, biggest first.

        SORTED ON THE BYTE COUNT, not on the rendered string. 'FolderSize' comes
        back as a ByteQuantifiedSize whose text is '1.9 GB (2,040,109,465
        bytes)' -- sorting that as text puts 9 MB above 1.9 GB and makes the
        table useless for the one question it exists to answer.
    #>
    param($Folders, [string]$Where)
    $rows = @()
    foreach ($f in @($Folders)) {
        # THE TEXT IS PARSED, NOT JUST .ToBytes(). ByteQuantifiedSize has a
        # ToBytes() method on a live object, but over a remote Exchange Online
        # session the property arrives DESERIALISED -- a string like
        # '1.9 GB (2,040,109,465 bytes)' with no methods on it. Relying on the
        # method meant every row scored zero, the sort silently did nothing,
        # and the table came back in Exchange's own order looking sorted.
        $bytes = 0
        foreach ($src in @($f.FolderAndSubfolderSize, $f.FolderSize)) {
            if ($bytes) { break }
            try { $bytes = [int64]$src.ToBytes() } catch { }
            if (-not $bytes -and "$src" -match '\(([\d,]+)\s*bytes\)') {
                try { $bytes = [int64]($Matches[1] -replace ',', '') } catch { }
            }
        }
        $rows += [PSCustomObject]@{
            Name  = "$($f.FolderPath)"
            Items = "$($f.ItemsInFolderAndSubfolders)"
            Size  = "$($f.FolderAndSubfolderSize)"
            Bytes = $bytes
        }
    }
    foreach ($r in @($rows | Sort-Object -Property Bytes -Descending)) {
        $i = $gridStats.Rows.Add($r.Name, $r.Items, $r.Size, $Where)
        if ($r.Name -match '(?i)Recoverable') {
            $gridStats.Rows[$i].DefaultCellStyle.ForeColor = (Get-ToolStatusColour -Kind Warn)
        }
    }
}

function Update-StatsTab {
    param([switch]$WithArchive)
    $gridStats.Rows.Clear()
    $stsSummary.Text = 'Reading folder statistics...'
    $stsSummary.Refresh()
    $m = $script:Mbx

    $folders = @(Invoke-EXOCommand -Name 'Get-MailboxFolderStatistics' -Parameters @{
                    Identity = $m.Smtp; ErrorAction = 'Stop' })
    Add-StatsRows -Folders $folders -Where 'Mailbox'

    # WORKED OUT BEFORE THE ARCHIVE READ. The archive block returns early when
    # it fails, and this one sentence is the reason the tab exists -- losing it
    # to a failure in an optional extra would be the wrong thing to drop.
    $rec = @($folders | Where-Object { "$($_.FolderType)" -eq 'RecoverableItemsRoot' }) | Select-Object -First 1
    $recText = $(if ($rec) { "Recoverable Items: $($rec.FolderAndSubfolderSize), $($rec.ItemsInFolderAndSubfolders) item(s)." }
                 else { 'Recoverable Items was not reported.' })

    $arcRows = 0
    if ($WithArchive) {
        try {
            $af = @(Invoke-EXOCommand -Name 'Get-MailboxFolderStatistics' -Parameters @{
                        Identity = $m.Smtp; Archive = $true; ErrorAction = 'Stop' })
            Add-StatsRows -Folders $af -Where 'Archive'
            $arcRows = @($af).Count
        } catch {
            $stsSummary.Text = "Mailbox folders read; the archive could not be: $($_.Exception.Message)" + "`r`n" + $recText
            return
        }
    }

    # RECOVERABLE ITEMS IS CALLED OUT BY NAME, because it is the one folder
    # that explains a mailbox being far bigger than anything the user can see,
    # and the one nobody thinks to look for.
    $stsSummary.Text = ("$(@($folders).Count) mailbox folder(s)$(if ($arcRows) { ", $arcRows archive folder(s)" }). $recText" + "`r`n" +
                        'Total size and quota are on Overview. If the space is in Recoverable Items, the answer is a hold or retention question, not a housekeeping one -- see Retention.')
}

$btnStsLoad.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Folder statistics')) { return }
        Update-StatsTab
        Set-EditorStatus -Text 'Folder statistics loaded.'
    } 'Load folder statistics'
})

$btnStsArch.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Folder statistics')) { return }
        Update-StatsTab -WithArchive
        Set-EditorStatus -Text 'Folder statistics loaded, including the archive.'
    } 'Load folder statistics with the archive'
})
#endregion

#region Inbox Rules tab ---------------------------------------------------------
<#
    THE RULES THE USER MADE, WHICH ARE OFTEN THE ANSWER.

    'Mail is disappearing', 'she is not getting anything from that supplier',
    'the message went to a folder nobody looks at' -- all rules, and all
    invisible from every other tab. The Health tab counts them and flags the
    ones that forward, redirect or delete; this is where you read them and
    turn them off.

    NOTHING HERE WRITES A RULE. Disable, enable and remove are the whole
    verb list on purpose: rule CONDITIONS are a small language, and an editor
    for them that is not Outlook's own is a way to silently change what a
    rule matches. Turning a bad rule off is what an administrator actually
    needs to do.
#>
$irY = 8
[void](New-TabLabel $tabRules 'INBOX RULES' 8 $irY -Bold)
$irY += 22
$irNote = New-TabLabel $tabRules '' 8 $irY 1040 -Muted
$irNote.Size = New-Object System.Drawing.Size(1040, 32)
$irY += 40
$chkIrHidden = New-Object System.Windows.Forms.CheckBox
$chkIrHidden.Location = New-Object System.Drawing.Point(8, $irY)
$chkIrHidden.Size     = New-Object System.Drawing.Size(600, 20)
$chkIrHidden.Text     = 'Include hidden rules (the ones Outlook creates for itself)'
$tabRules.Controls.Add($chkIrHidden)
$irY += 26
$gridRules = New-TabGrid $tabRules @(@('Rule', 300), @('On', 60), @('Priority', 70), @('What it does', 280)) 8 $irY 760 300
Add-GridPickColumn -Grid $gridRules
$irBtnY = $irY
$btnIrDisable = New-TabButton $tabRules 'Turn off ticked' 780 $irBtnY 170
$irBtnY += 34
$btnIrEnable  = New-TabButton $tabRules 'Turn on ticked'  780 $irBtnY 170
$irBtnY += 34
$btnIrRemove  = New-TabButton $tabRules 'Delete ticked'   780 $irBtnY 170
$irBtnY += 34
$btnIrAll     = New-TabButton $tabRules 'Tick all'        780 $irBtnY 82
$btnIrNone    = New-TabButton $tabRules 'None'            868 $irBtnY 82
$irBtnY += 42
$btnIrRefresh = New-TabButton $tabRules 'Refresh'         780 $irBtnY 170
$irY += 312
$irDetail            = New-Object System.Windows.Forms.TextBox
$irDetail.Multiline  = $true
$irDetail.ReadOnly   = $true
$irDetail.ScrollBars = 'Vertical'
$irDetail.WordWrap   = $true
$irDetail.Location   = New-Object System.Drawing.Point(8, $irY)
$irDetail.Size       = New-Object System.Drawing.Size(1040, 150)
$irDetail.Font       = New-Object System.Drawing.Font('Consolas', 9)
$irDetail.Tag        = 'NoTheme'
$irDetail.BackColor  = [System.Drawing.Color]::WhiteSmoke
$irDetail.ForeColor  = [System.Drawing.Color]::Black
$tabRules.Controls.Add($irDetail)

function Format-RuleAction {
    # What a rule DOES, in a phrase. The full description goes in the box
    # below; this is what makes a bad rule obvious at a glance in the grid.
    param($Rule)
    $acts = @()
    if ($Rule.ForwardTo -or $Rule.ForwardAsAttachmentTo) { $acts += 'FORWARDS' }
    if ($Rule.RedirectTo)                                { $acts += 'REDIRECTS' }
    if ($Rule.DeleteMessage -or $Rule.SoftDeleteMessage) { $acts += 'DELETES' }
    if ($Rule.MoveToFolder)                              { $acts += "moves to $($Rule.MoveToFolder)" }
    if ($Rule.CopyToFolder)                              { $acts += "copies to $($Rule.CopyToFolder)" }
    if ($Rule.MarkAsRead)                                { $acts += 'marks read' }
    if ($Rule.StopProcessingRules)                       { $acts += 'stops other rules' }
    if ($acts.Count -eq 0) { return '(no action this tool recognises)' }
    return ($acts -join ', ')
}

function Update-RulesTab {
    $gridRules.Rows.Clear()
    $irDetail.Text = ''
    $rules = @()
    try {
        $p = @{ Mailbox = $script:Mbx.Smtp; ErrorAction = 'Stop' }
        if ($chkIrHidden.Checked) { $p['IncludeHidden'] = $true }
        $rules = @(Invoke-EXOCommand -Name 'Get-InboxRule' -Parameters $p)
    } catch {
        $script:TabReadFailed = $true
        $irNote.Text = "The inbox rules could not be read: $($_.Exception.Message)"
        $irNote.ForeColor = (Get-ToolStatusColour -Kind Bad)
        return
    }

    $risky = 0
    foreach ($r in ($rules | Sort-Object Priority)) {
        $what = Format-RuleAction -Rule $r
        $i = $gridRules.Rows.Add($false, "$($r.Name)", $(if ($r.Enabled) { 'on' } else { 'OFF' }), "$($r.Priority)", $what)
        $gridRules.Rows[$i].Tag = [PSCustomObject]@{
            Identity = "$($r.Identity)"; Name = "$($r.Name)"; Enabled = [bool]$r.Enabled; Raw = $r
        }
        if ($r.Enabled -and $what -match 'FORWARDS|REDIRECTS|DELETES') {
            $gridRules.Rows[$i].DefaultCellStyle.ForeColor = (Get-ToolStatusColour -Kind Bad)
            $risky++
        } elseif (-not $r.Enabled) {
            $gridRules.Rows[$i].DefaultCellStyle.ForeColor = (Get-ToolStatusColour -Kind Muted)
        }
    }

    $irNote.Text = "$($gridRules.Rows.Count) rule(s)$(if ($risky) { ". $risky enabled rule(s) forward, redirect or delete -- shown in red. A rule that forwards externally and deletes the evidence is the classic compromise footprint." })"
    $irNote.ForeColor = $(if ($risky) { (Get-ToolStatusColour -Kind Bad) } else { (Get-ToolStatusColour -Kind Muted) })
}

function Show-RuleDetail {
    if ($gridRules.SelectedRows.Count -eq 0) { return }
    $t = $gridRules.SelectedRows[0].Tag
    if (-not $t) { return }
    $r = $t.Raw
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Rule        : $($r.Name)")
    [void]$sb.AppendLine("Enabled     : $($r.Enabled)")
    [void]$sb.AppendLine("Priority    : $($r.Priority)")
    # Description is Exchange's own rendering of the whole rule, conditions
    # and all -- better than anything reassembled property by property here.
    if ("$($r.Description)".Trim()) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('As Exchange describes it:')
        [void]$sb.AppendLine("$($r.Description)")
    }
    $irDetail.Text = $sb.ToString()
}

$gridRules.Add_SelectionChanged({ Show-RuleDetail })
$btnIrAll.Add_Click({  Set-GridPicked -Grid $gridRules -On $true })
$btnIrNone.Add_Click({ Set-GridPicked -Grid $gridRules -On $false })
$btnIrRefresh.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction { Update-RulesTab; Set-EditorStatus -Text 'Inbox rules re-read.' } 'Refresh inbox rules'
})
$chkIrHidden.Add_CheckedChanged({
    if ($script:Busy) { return }
    if (-not $script:Mbx -or -not $script:Mbx.Ok) { return }
    Invoke-EditorAction { Update-RulesTab } 'Read inbox rules'
})

function Invoke-RuleVerb {
    param([ValidateSet('Enable', 'Disable', 'Remove')][string]$Verb)
    if (-not (Test-EditorCloudOnly -What 'Inbox rules')) { return }
    $picked = Get-GridPicked -Grid $gridRules
    if ($picked.Count -eq 0) { Set-EditorStatus -Text 'Tick one or more rules first.' -Bad; return }

    $word = @{ Enable = 'Turn ON'; Disable = 'Turn OFF'; Remove = 'DELETE' }[$Verb]
    $extra = $(if ($Verb -eq 'Remove') { "`r`n`r`nDeleting a rule cannot be undone from here, and the user will not be told. Turning it off is reversible and usually enough." } else { '' })
    $a = Show-ToolDialog(("$word $($picked.Count) rule(s) on $($script:Mbx.Display)?" + "`r`n`r`n  " +
                          (@($picked | ForEach-Object { $_.Name }) -join "`r`n  ") + $extra),
                         "$word inbox rules", 'OKCancel', 'Warning')
    if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $lines = @(); $ok = 0; $bad = 0
    foreach ($e in $picked) {
        try {
            # LITERAL CMDLET NAMES, not "$Verb-InboxRule". The capability
            # scanner and the audit-coverage test read this file for the names.
            $p = @{ Identity = $e.Identity; Confirm = $false; ErrorAction = 'Stop' }
            switch ($Verb) {
                'Enable'  { Invoke-EXOCommand -Name 'Enable-InboxRule'  -Parameters $p | Out-Null }
                'Disable' { Invoke-EXOCommand -Name 'Disable-InboxRule' -Parameters $p | Out-Null }
                'Remove'  { Invoke-EXOCommand -Name 'Remove-InboxRule'  -Parameters $p | Out-Null }
            }
            $lines += "OK      $($e.Name)"; $ok++
        } catch { $lines += "FAILED  $($e.Name): $($_.Exception.Message)"; $bad++ }
    }
    Write-EditorAudit -Action "MailboxInboxRule$Verb" `
        -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
        -Detail "${word}: $(($picked | ForEach-Object { $_.Name }) -join '; ')"
    if ($bad) { Show-ResultListDialog -Title "$word rules -- $ok done, $bad failed" -Lines $lines }
    Update-RulesTab
    $script:Loaded['Health'] = $false
    Set-EditorStatus -Text "$word $ok rule(s), failed $bad." -Bad:($bad -gt 0)
}

$btnIrDisable.Add_Click({ if (Test-EditorLoaded) { Invoke-EditorAction { Invoke-RuleVerb -Verb 'Disable' } 'Turn off inbox rules' } })
$btnIrEnable.Add_Click({  if (Test-EditorLoaded) { Invoke-EditorAction { Invoke-RuleVerb -Verb 'Enable' }  'Turn on inbox rules' } })
$btnIrRemove.Add_Click({  if (Test-EditorLoaded) { Invoke-EditorAction { Invoke-RuleVerb -Verb 'Remove' }  'Delete inbox rules' } })
#endregion

#region Recover Items tab -------------------------------------------------------
<#
    THE EAC'S 'RECOVER DELETED ITEMS', HERE.

    Items the user deleted are in Recoverable Items until the retention window
    closes. This searches that and puts things back -- which is the request
    that arrives as 'I emptied Deleted Items and I need that message'.

    THREE PLACES, NAMED. An item is in the Deleted Items folder, or in
    Recoverable Items after that folder was emptied, or in Purges after
    Recoverable Items was cleared out too. All three are searched by default
    and every row says WHICH ONE it came from, because that is the difference
    between 'the user can drag this back themselves' and 'a hold is the only
    reason this still exists'. A fourth, DiscoveryHolds, is searched only when
    it is asked for -- see the picker.

    IT DOES NOT LOAD ON TAB SELECTION. There is no sensible default search
    over a mailbox's whole deleted history, and running one on the way past
    would spend seconds fetching something nobody asked for. It loads on
    Search.

    RESTORE IS BY ENTRY ID, ONE ITEM AT A TIME, and never by filter.
    Restore-RecoverableItems can take the same filters as the search, which
    means a slightly different filter at restore time than at search time
    restores a different set than the one on screen. Restoring exactly the
    rows that are ticked is the only version where what you see is what you
    get.
#>
$rsY = 8
[void](New-TabLabel $tabRestore 'RECOVER DELETED ITEMS' 8 $rsY -Bold)
$rsY += 22
$rsNote = New-TabLabel $tabRestore '' 8 $rsY 1040 -Muted
$rsNote.Size = New-Object System.Drawing.Size(1040, 32)
$rsNote.Text = ('Searches Deleted Items, Recoverable Items and Purges together. The Stage column says which of the three each item is in. Items age out when the retention window closes; that window is on the Recovery tab.' + "`r`n" +
                'Search by subject, type and date -- the service does not return or filter on sender. Restored items go back to the folder they came from.')
$rsY += 40

[void](New-TabLabel $tabRestore 'Subject contains:' 8 ($rsY + 3) 110)
$txtRsSubject          = New-Object System.Windows.Forms.TextBox
$txtRsSubject.Location = New-Object System.Drawing.Point(124, $rsY)
$txtRsSubject.Size     = New-Object System.Drawing.Size(280, 24)
$tabRestore.Controls.Add($txtRsSubject)
[void](New-TabLabel $tabRestore 'Type:' 420 ($rsY + 3) 44)
$cmbRsType               = New-Object System.Windows.Forms.ComboBox
$cmbRsType.DropDownStyle = 'DropDownList'
$cmbRsType.Location      = New-Object System.Drawing.Point(468, $rsY)
$cmbRsType.Size          = New-Object System.Drawing.Size(150, 24)
foreach ($t in @('(any)', 'IPM.Note', 'IPM.Appointment', 'IPM.Contact', 'IPM.Task', 'IPM.StickyNote')) { [void]$cmbRsType.Items.Add($t) }
$cmbRsType.SelectedIndex = 0
$tabRestore.Controls.Add($cmbRsType)

# WHERE TO LOOK -- and why this picker earns its space.
#
# Omitting -SourceFolder searches Deleted Items, Recoverable Items and Purged
# Items together, which is what this tab has always done and what it still
# does by default. What it could not do was SAY WHICH of the three an item is
# sitting in, and an admin reads far more out of that one word than out of the
# subject line: still in Deleted Items means the user could have done this
# themselves, and Purged means a hold is the only reason there is anything
# left to restore.
#
# SO 'EVERYWHERE' NOW RUNS ONE SEARCH PER FOLDER and stamps each row with the
# folder that answered. Three calls instead of one, for a column that changes
# what the operator does next.
#
# ON HOLD (DiscoveryHolds) IS NOT PART OF 'EVERYWHERE', deliberately. The
# service does not include that folder when -SourceFolder is omitted, so
# folding it into the default would quietly widen what a plain search means
# and surface hold-protected items to somebody who only asked where a message
# went. It is one pick away when it is genuinely wanted.
[void](New-TabLabel $tabRestore 'Look in:' 634 ($rsY + 3) 60)
$cmbRsWhere               = New-Object System.Windows.Forms.ComboBox
$cmbRsWhere.DropDownStyle = 'DropDownList'
$cmbRsWhere.Location      = New-Object System.Drawing.Point(698, $rsY)
$cmbRsWhere.Size          = New-Object System.Drawing.Size(212, 24)
foreach ($w in @('Everywhere', 'Deleted Items only', 'Recoverable only', 'Purged only', 'On hold only')) { [void]$cmbRsWhere.Items.Add($w) }
$cmbRsWhere.SelectedIndex = 0
$tabRestore.Controls.Add($cmbRsWhere)
$rsY += 28

$chkRsFrom = New-Object System.Windows.Forms.CheckBox
$chkRsFrom.Location = New-Object System.Drawing.Point(8, ($rsY + 2))
$chkRsFrom.Size     = New-Object System.Drawing.Size(70, 20)
$chkRsFrom.Text     = 'From:'
$tabRestore.Controls.Add($chkRsFrom)
$dtRsFrom          = New-Object System.Windows.Forms.DateTimePicker
$dtRsFrom.Location = New-Object System.Drawing.Point(84, $rsY)
$dtRsFrom.Size     = New-Object System.Drawing.Size(200, 24)
$dtRsFrom.Value    = (Get-Date).AddDays(-30)
$tabRestore.Controls.Add($dtRsFrom)
$chkRsTo = New-Object System.Windows.Forms.CheckBox
$chkRsTo.Location = New-Object System.Drawing.Point(300, ($rsY + 2))
$chkRsTo.Size     = New-Object System.Drawing.Size(60, 20)
$chkRsTo.Text     = 'To:'
$tabRestore.Controls.Add($chkRsTo)
$dtRsTo          = New-Object System.Windows.Forms.DateTimePicker
$dtRsTo.Location = New-Object System.Drawing.Point(366, $rsY)
$dtRsTo.Size     = New-Object System.Drawing.Size(200, 24)
$tabRestore.Controls.Add($dtRsTo)
$btnRsSearch = New-TabButton $tabRestore 'Search' 584 ($rsY - 2) 120
$rsY += 32

function Get-RsStageFolders {
    <#
        The -SourceFolder values one pick means.

        'Everywhere' returns the same three the service searches when the
        parameter is left off, so the default coverage is unchanged and only
        the labelling is new. DiscoveryHolds is reachable ONLY by asking for
        it, which is the point of listing it separately.
    #>
    param([string]$Choice)
    switch ("$Choice") {
        'Deleted Items only' { return @('DeletedItems') }
        'Recoverable only'   { return @('RecoverableItems') }
        'Purged only'        { return @('PurgedItems') }
        'On hold only'       { return @('DiscoveryHoldsItems') }
        default              { return @('DeletedItems', 'RecoverableItems', 'PurgedItems') }
    }
}

function Get-RsStageLabel {
    # The cmdlet's folder name in the words the operator uses. Anything
    # unrecognised passes through rather than becoming a blank cell.
    param([string]$Folder)
    switch ("$Folder") {
        'DeletedItems'        { return 'Deleted Items' }
        'RecoverableItems'    { return 'Recoverable' }
        'PurgedItems'         { return 'Purged' }
        'DiscoveryHoldsItems' { return 'Holds' }
        default               { return "$Folder" }
    }
}

function Get-RsDeletedStamp {
    <#
        The deletion time as LOCAL time, in a shape that sorts.

        THE SERVICE ANSWERS IN UTC and the grid used to print whatever
        stringifying that produced -- a UTC stamp sitting next to a 'days until
        purge' figure worked out against local now. Two clocks in one row is
        how somebody reads 'this morning' as 'yesterday' and stops looking.

        A VALUE THAT WILL NOT CONVERT IS PASSED THROUGH, not blanked. An odd
        stamp is still evidence about the item; an empty cell is not.
    #>
    param($Value)
    if ($null -eq $Value) { return '' }
    try { return ([datetime]$Value).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } catch { return "$Value" }
}

function Get-RsPurgeVerdict {
    <#
        How long this item has before the service purges it for good.

        THE SERVICE DOES NOT RETURN THIS, so it is worked out: deleted-items
        retention (RetainDeletedItemsFor on the mailbox) minus how long the
        item has already sat there. LastModifiedTime is the deletion stamp the
        cmdlet itself filters on -- FilterStartTime and FilterEndTime are
        documented as using it -- so it is the right clock to count from.

        A HOLD STOPS THE CLOCK ENTIRELY, and saying '3 days' about an item
        that is never going to be purged would send somebody rushing to
        restore something that was not going anywhere. Litigation hold,
        retention hold, a compliance tag hold and a delay hold all have that
        effect, so any of them reports 'on hold' instead of a number.

        AN HONEST BLANK BEATS A CONFIDENT GUESS. No retention value, no
        deletion stamp, or a mailbox that has not loaded yet all return '?'.
    #>
    param($DeletedOn)

    $mb = $(if ($script:Mbx) { $script:Mbx.Raw } else { $null })
    if (-not $mb) { return '?' }

    foreach ($h in @('LitigationHoldEnabled', 'RetentionHoldEnabled', 'ComplianceTagHoldApplied', 'DelayHoldApplied', 'DelayReleaseHoldApplied')) {
        if ($mb.PSObject.Properties[$h] -and [bool]$mb.$h) { return 'on hold' }
    }
    if ($mb.PSObject.Properties['InPlaceHolds'] -and @($mb.InPlaceHolds).Count -gt 0) { return 'on hold' }

    $days = $null
    try {
        $r = $mb.RetainDeletedItemsFor
        if ($r -is [timespan]) { $days = $r.TotalDays }
        elseif ("$r".Trim()) { $days = ([timespan]::Parse("$r")).TotalDays }
    } catch { $days = $null }
    if ($null -eq $days) { return '?' }

    $when = $null
    try { if ($DeletedOn) { $when = [datetime]$DeletedOn } } catch { $when = $null }
    if (-not $when) { return '?' }

    $left = [int][Math]::Floor($days - ((Get-Date) - $when).TotalDays)
    if ($left -lt 0)  { return 'due' }
    if ($left -eq 0)  { return 'today' }
    return "$left d"
}

# THERE IS NO FROM OR TO COLUMN, AND THAT IS NOT AN OVERSIGHT.
#
# An earlier build carried both, filled by probing the returned item for any
# of several plausible property names. On a live tenant every cell came back
# empty, and the reason is not a wrong spelling: Get-RecoverableItems has no
# sender or recipient PARAMETER either, so there is nothing to filter on and
# nothing to show. The published reference has no Outputs section at all, so
# the documentation can neither confirm nor deny it -- but a cmdlet whose
# whole filter surface is subject, item type and date is not one that carries
# addresses, and the live result agrees.
#
# A COLUMN THAT IS ALWAYS EMPTY IS WORSE THAN NO COLUMN. Blank under 'From'
# reads as 'this message had no sender' rather than 'the service does not tell
# us', and it costs grid width that Subject actually needs. If a future build
# starts returning addresses, add the column back then.
#
# PURGE IN is computed, not returned. See Get-RsPurgeVerdict.
# 'WAS IN' AND 'STAGE' ARE NOT THE SAME COLUMN and sit side by side so that is
# obvious. Was in is the folder the item lived in before it was deleted, which
# is where a restore puts it back. Stage is how far down the disposal chain it
# has travelled since.
$gridRestore = New-TabGrid $tabRestore @(@('Deleted', 120), @('Subject', 340), @('Type', 90), @('Was in', 145), @('Stage', 95), @('Purge in', 70)) 8 $rsY 890 330
Add-GridPickColumn -Grid $gridRestore
$rsBtnY = $rsY
$btnRsRestore = New-TabButton $tabRestore 'Restore ticked' 910 $rsBtnY 140
$rsBtnY += 34
$btnRsAll     = New-TabButton $tabRestore 'Tick all'       910 $rsBtnY 66
$btnRsNone    = New-TabButton $tabRestore 'None'           982 $rsBtnY 68
$rsY += 342
$rsHint = New-TabLabel $tabRestore '' 8 $rsY 1040 -Muted
$rsHint.Size = New-Object System.Drawing.Size(1040, 46)
$rsHint.Text = ('Leave the boxes empty to see everything still recoverable. A date box only filters when its tick box is on -- an unticked one is not "today".' + "`r`n" +
                'Stage: Deleted Items, then Recoverable once that folder is emptied, then Purged. On hold searches DiscoveryHolds, which a plain search does not. Nothing here purges.')

$btnRsAll.Add_Click({  Set-GridPicked -Grid $gridRestore -On $true })
$btnRsNone.Add_Click({ Set-GridPicked -Grid $gridRestore -On $false })

$btnRsSearch.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Recovering deleted items')) { return }
        $gridRestore.Rows.Clear()
        Set-EditorStatus -Text 'Searching recoverable items...'

        # ONE CALL PER FOLDER, so every row can say which folder answered.
        # 'Everywhere' asks the same three the service searches when
        # -SourceFolder is left off, so the coverage of a default search is
        # exactly what it was before the picker existed.
        $folders = Get-RsStageFolders -Choice "$($cmbRsWhere.SelectedItem)"

        $base = @{
            Identity     = $script:Mbx.Smtp
            ResultSize   = 5000
            ErrorAction  = 'Stop'
        }
        if ("$($txtRsSubject.Text)".Trim()) { $base['SubjectContains'] = "$($txtRsSubject.Text)".Trim() }
        if ("$($cmbRsType.SelectedItem)" -ne '(any)') { $base['FilterItemType'] = "$($cmbRsType.SelectedItem)" }
        # A DATE BOX ONLY FILTERS WHEN ITS TICK BOX IS ON. An unticked picker
        # still holds a date, and sending it silently narrows every search to
        # a window nobody chose.
        if ($chkRsFrom.Checked) { $base['FilterStartTime'] = $dtRsFrom.Value }
        if ($chkRsTo.Checked)   { $base['FilterEndTime']   = $dtRsTo.Value }

        $counts     = @()
        $capped     = @()
        $refused    = @()

        foreach ($f in $folders) {
            $stage = Get-RsStageLabel -Folder $f
            $p = $base.Clone()
            $p['SourceFolder'] = $f

            # ONE FOLDER FAILING IS NOT THE SEARCH FAILING. A tenant that will
            # not answer for DiscoveryHolds should still show what is sitting
            # in Deleted Items, so each folder is asked in its own try and a
            # refusal is reported under the grid rather than thrown.
            $items = @()
            try { $items = @(Invoke-EXOCommand -Name 'Get-RecoverableItems' -Parameters $p) }
            catch { $refused += "$stage ($($_.Exception.Message))"; continue }

            foreach ($it in $items) {
                $i = $gridRestore.Rows.Add($false, (Get-RsDeletedStamp -Value $it.LastModifiedTime), "$($it.Subject)",
                                           "$($it.ItemClass)", "$($it.LastParentPath)", $stage,
                                           (Get-RsPurgeVerdict -DeletedOn $it.LastModifiedTime))
                $gridRestore.Rows[$i].Tag = [PSCustomObject]@{
                    EntryId = "$($it.EntryID)"; Subject = "$($it.Subject)"; Folder = $f
                }
            }
            $counts += "$stage $($items.Count)"
            # THE CAP IS PER FOLDER, so it is reported per folder. 'Some results
            # were hidden' without saying where is not something anyone can act
            # on.
            if ($items.Count -ge 5000) { $capped += $stage }
        }

        $rsNote.Text = "$($gridRestore.Rows.Count) recoverable item(s) -- $($counts -join ', '). Tick what you want back and press Restore."
        if ($capped.Count -gt 0) {
            $rsNote.Text += "  Stopped at the 5,000 cap in: $($capped -join ', ') -- narrow the search, there is more than this."
        }
        if ($refused.Count -gt 0) {
            $rsNote.Text += "  Not searched: $($refused -join '; ')."
        }
        $rsNote.ForeColor = (Get-ToolStatusColour -Kind $(if ($capped.Count -or $refused.Count) { 'Warn' } else { 'Muted' }))
        Set-EditorStatus -Text "$($gridRestore.Rows.Count) recoverable item(s) found."
    } 'Search recoverable items'
})

$btnRsRestore.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Recovering deleted items')) { return }
        $picked = Get-GridPicked -Grid $gridRestore
        if ($picked.Count -eq 0) { Set-EditorStatus -Text 'Tick one or more items to restore.' -Bad; return }
        $shown = @($picked | Select-Object -First 15 | ForEach-Object { "  $($_.Subject)" }) -join "`r`n"
        if ($picked.Count -gt 15) { $shown += "`r`n  ... and $($picked.Count - 15) more" }
        $a = Show-ToolDialog(("Restore $($picked.Count) item(s) to $($script:Mbx.Display)?" + "`r`n`r`n" + $shown + "`r`n`r`n" +
                              'They go back to the folder they were deleted from.'),
                             'Restore deleted items', 'OKCancel', 'Question')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $lines = @(); $ok = 0; $bad = 0
        foreach ($e in $picked) {
            try {
                $rp = @{ Identity = $script:Mbx.Smtp; EntryID = $e.EntryId; ErrorAction = 'Stop' }
                # -SourceFolder IS NAMED FOR THE HOLDS FOLDER AND NOTHING ELSE.
                # With the parameter omitted, Restore-RecoverableItems looks in
                # the same three folders a plain search covers -- and an item
                # the picker found in DiscoveryHolds is not in any of them, so
                # for that one bucket the folder has to be said out loud or the
                # restore cannot find what the grid just showed. For the other
                # three, omitting is the proven path and stays: EntryID is
                # unique in the mailbox and naming a folder only adds a way to
                # disagree with the search.
                if ("$($e.Folder)" -eq 'DiscoveryHoldsItems') { $rp['SourceFolder'] = 'DiscoveryHoldsItems' }
                Invoke-EXOCommand -Name 'Restore-RecoverableItems' -Parameters $rp | Out-Null
                $lines += "OK      $($e.Subject)"; $ok++
            } catch { $lines += "FAILED  $($e.Subject): $($_.Exception.Message)"; $bad++ }
        }
        Write-EditorAudit -Action 'MailboxRestoreItems' `
            -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
            -Detail "$ok restored, $bad failed"
        Show-ResultListDialog -Title "Restore items -- $ok done, $bad failed" -Lines $lines
        Set-EditorStatus -Text "Restored $ok, failed $bad." -Bad:($bad -gt 0)
    } 'Restore deleted items'
})
#endregion

#region Quarantine tab ----------------------------------------------------------
<#
    WHAT THE SERVICE HELD BACK FOR THIS MAILBOX.

    Same question as the quarantine tool asks tenant-wide, narrowed to the
    person in front of you -- which is how it actually arrives: 'did we block
    something from them?'

    RELEASING PHISH IS NOT THE SAME AS RELEASING SPAM, and the confirmation
    says which is which. A message quarantined as high-confidence phish was
    not a close call, and releasing one puts a live credential-harvesting mail
    into somebody's inbox with the service's own verdict overruled.

    THE CMDLETS ARE PROBED, NOT ASSUMED. Get-QuarantineMessage pages rather
    than taking -ResultSize on some builds, and a parameter it does not
    declare fails the whole call.
#>
$qtY = 8
[void](New-TabLabel $tabQuar 'QUARANTINE' 8 $qtY -Bold)
$qtY += 22
$qtNote = New-TabLabel $tabQuar '' 8 $qtY 1040 -Muted
$qtNote.Size = New-Object System.Drawing.Size(1040, 32)
$qtNote.Text = 'Messages the service held back for this mailbox. Quarantine keeps items for a limited number of days and then discards them, so nothing here is permanent.'
$qtY += 40

[void](New-TabLabel $tabQuar 'Last' 8 ($qtY + 3) 34)
$numQtDays          = New-Object System.Windows.Forms.NumericUpDown
$numQtDays.Location = New-Object System.Drawing.Point(46, $qtY)
$numQtDays.Size     = New-Object System.Drawing.Size(60, 24)
$numQtDays.Minimum  = 1; $numQtDays.Maximum = 30; $numQtDays.Value = 7
$tabQuar.Controls.Add($numQtDays)
[void](New-TabLabel $tabQuar 'days' 112 ($qtY + 3) 40)
[void](New-TabLabel $tabQuar 'Type:' 158 ($qtY + 3) 44)
$cmbQtType               = New-Object System.Windows.Forms.ComboBox
$cmbQtType.DropDownStyle = 'DropDownList'
$cmbQtType.Location      = New-Object System.Drawing.Point(206, $qtY)
$cmbQtType.Size          = New-Object System.Drawing.Size(180, 24)
foreach ($t in @('(any)', 'Spam', 'HighConfPhish', 'Phish', 'Malware', 'Bulk', 'TransportRule')) { [void]$cmbQtType.Items.Add($t) }
$cmbQtType.SelectedIndex = 0
$tabQuar.Controls.Add($cmbQtType)
$btnQtSearch = New-TabButton $tabQuar 'Search' 400 ($qtY - 2) 120
$qtY += 32

$gridQuar = New-TabGrid $tabQuar @(@('Received', 140), @('Sender', 250), @('Subject', 330), @('Type', 120), @('Released', 80)) 8 $qtY 890 330
Add-GridPickColumn -Grid $gridQuar
$qtBtnY = $qtY
$btnQtRelease = New-TabButton $tabQuar 'Release ticked' 910 $qtBtnY 140
$qtBtnY += 34
$btnQtAll     = New-TabButton $tabQuar 'Tick all'       910 $qtBtnY 66
$btnQtNone    = New-TabButton $tabQuar 'None'           982 $qtBtnY 68
$qtY += 342
$qtHint = New-TabLabel $tabQuar '' 8 $qtY 1040 -Muted
$qtHint.Size = New-Object System.Drawing.Size(1040, 46)
$qtHint.Text = ('Release sends the message to ALL of its original recipients, not only this mailbox -- that is what the service supports, and it is worth knowing before you press it.' + "`r`n" +
                'Nothing here deletes from quarantine; use the Quarantine Manager tool for that.')

$btnQtAll.Add_Click({  Set-GridPicked -Grid $gridQuar -On $true })
$btnQtNone.Add_Click({ Set-GridPicked -Grid $gridQuar -On $false })

$btnQtSearch.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Quarantine')) { return }
        $gridQuar.Rows.Clear()
        $script:QuarCapped = $false
        Set-EditorStatus -Text 'Searching quarantine...'

        # PROBED, NOT ASSUMED. Get-QuarantineMessage pages on some builds and
        # takes -ResultSize on others, and sending a parameter the cmdlet does
        # not declare fails the whole call rather than being ignored.
        $declared = @()
        try { $declared = @(Get-CmdletParameterNames -Name 'Get-QuarantineMessage') } catch { }

        $p = @{
            RecipientAddress  = $script:Mbx.Smtp
            StartReceivedDate = (Get-Date).AddDays(-[int]$numQtDays.Value)
            EndReceivedDate   = (Get-Date)
            ErrorAction       = 'Stop'
        }
        if ($declared -contains 'ResultSize') { $p['ResultSize'] = 'Unlimited' }
        if ("$($cmbQtType.SelectedItem)" -ne '(any)') { $p['QuarantineTypes'] = "$($cmbQtType.SelectedItem)" }

        $msgs = @()
        if (($declared -contains 'Page') -and ($declared -contains 'PageSize')) {
            $page = 1
            while ($true) {
                $q = @{} + $p
                $q['Page'] = $page; $q['PageSize'] = 1000
                $batch = @(Invoke-EXOCommand -Name 'Get-QuarantineMessage' -Parameters $q)
                $msgs += $batch
                if (@($batch).Count -lt 1000) { break }
                # A CAP THAT IS REACHED IS SAID OUT LOUD. A silent truncation
                # presents a partial list as the whole answer, which is the
                # one thing a search must never do.
                if ($page -ge 20) { $script:QuarCapped = $true; break }
                $page++
            }
        } else {
            $msgs = @(Invoke-EXOCommand -Name 'Get-QuarantineMessage' -Parameters $p)
        }

        foreach ($m in $msgs) {
            $i = $gridQuar.Rows.Add($false, "$($m.ReceivedTime)", "$($m.SenderAddress)", "$($m.Subject)",
                                    "$($m.QuarantineTypes)", $(if ($m.Released) { 'yes' } else { '' }))
            $gridQuar.Rows[$i].Tag = [PSCustomObject]@{
                Identity = "$($m.Identity)"; Subject = "$($m.Subject)"; Sender = "$($m.SenderAddress)"
                Type = "$($m.QuarantineTypes)"; Released = [bool]$m.Released
            }
            if ("$($m.QuarantineTypes)" -match '(?i)phish|malware') {
                $gridQuar.Rows[$i].DefaultCellStyle.ForeColor = (Get-ToolStatusColour -Kind Bad)
            }
        }
        $qtNote.Text = "$($gridQuar.Rows.Count) quarantined message(s) in the last $([int]$numQtDays.Value) day(s). Phish and malware are shown in red.$(if ($script:QuarCapped) { '  THIS IS NOT ALL OF THEM -- the search stopped at 20,000. Narrow the window or the type.' })"
        if ($script:QuarCapped) { $qtNote.ForeColor = (Get-ToolStatusColour -Kind Bad) }
        $qtNote.ForeColor = (Get-ToolStatusColour -Kind Muted)
        Set-EditorStatus -Text "$($gridQuar.Rows.Count) quarantined message(s)."
    } 'Search quarantine'
})

$btnQtRelease.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Quarantine')) { return }
        $picked = Get-GridPicked -Grid $gridQuar
        if ($picked.Count -eq 0) { Set-EditorStatus -Text 'Tick one or more messages to release.' -Bad; return }

        # PHISH AND MALWARE GET THEIR OWN SENTENCE. A message held as
        # high-confidence phish was not a marginal call, and releasing one puts
        # a live credential-harvesting mail in front of somebody with the
        # service's verdict overruled by hand.
        $nasty = @($picked | Where-Object { $_.Type -match '(?i)phish|malware' })
        $extra = ''
        if ($nasty.Count) {
            $extra = "`r`n`r`n$($nasty.Count) of these were quarantined as PHISH or MALWARE. That is not a close call the filter got wrong -- releasing one delivers a live attack to everybody it was addressed to. Be sure."
        }
        $shown = @($picked | Select-Object -First 12 | ForEach-Object { "  $($_.Sender)  --  $($_.Subject)  [$($_.Type)]" }) -join "`r`n"
        if ($picked.Count -gt 12) { $shown += "`r`n  ... and $($picked.Count - 12) more" }

        $a = Show-ToolDialog(
            ("Release $($picked.Count) message(s) to ALL of their original recipients?" + "`r`n`r`n" + $shown + $extra),
            'Release from quarantine', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $lines = @(); $ok = 0; $bad = 0
        foreach ($e in $picked) {
            try {
                $q = @{ Identity = $e.Identity; ReleaseToAll = $true; Confirm = $false; ErrorAction = 'Stop' }
                # -Force is what lets an already-released message be released
                # again; without it the cmdlet refuses and the operator reads
                # that as a failure.
                if ($e.Released) { $q['Force'] = $true }
                Invoke-EXOCommand -Name 'Release-QuarantineMessage' -Parameters $q | Out-Null
                $lines += "OK      $($e.Sender) -- $($e.Subject)"; $ok++
            } catch { $lines += "FAILED  $($e.Sender) -- $($e.Subject): $($_.Exception.Message)"; $bad++ }
        }
        Write-EditorAudit -Action 'MailboxQuarantineRelease' `
            -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
            -Detail "$ok released, $bad failed$(if ($nasty.Count) { "; $($nasty.Count) were phish or malware" })"
        Show-ResultListDialog -Title "Release from quarantine -- $ok done, $bad failed" -Lines $lines
        Set-EditorStatus -Text "Released $ok, failed $bad." -Bad:($bad -gt 0)
    } 'Release from quarantine'
})
#endregion

#region Load and tab switching --------------------------------------------------
function Invoke-EditorLoad {
    $t = "$($txtMbx.Text)".Trim()
    if (-not $t) { Set-EditorStatus -Text 'Enter a mailbox address.' -Bad; return }
    Set-EditorStatus -Text "Resolving $t..."
    $m = Resolve-MailboxTarget -Smtp $t
    if (-not $m.Ok) {
        $script:Mbx = $null
        $lblIdent.Text = ''; $lblSync.Text = ''
        # A MAILBOX THAT NEVER MIGRATED lands here, and the bare cloud reason
        # ('not found') is misleading -- the mailbox exists, it is simply not
        # in Exchange Online. Say which it is rather than letting somebody
        # conclude the mailbox is gone.
        $hint = ''
        if ($cmbMode.Enabled) {
            try {
                $probe = Resolve-OnPremMailboxTarget -Identity $t
                if ($probe.Ok -and $probe.Kind -eq 'Mailbox') {
                    $hint = ' It DOES exist on-premises as a mailbox that has not been migrated -- this window cannot open one of those yet.'
                }
            } catch { }
        }
        Set-EditorStatus -Text ("'$t' $($m.Reason)$hint") -Bad
        return
    }
    $script:Mbx = $m
    $script:Loaded = @{}
    $script:OnPrem = $null
    # CLEARED ON EVERY LOAD. The MBX Stats tab deliberately has no entry in
    # Invoke-EditorTabLoad, so nothing would ever refresh it -- mailbox A's
    # folder table would sit there under mailbox B's name. The licence is
    # cleared for the same reason: the Health tab's findings read it, and a
    # stale one makes a finding about the wrong person.
    $script:Licence = $null
    $gridStats.Rows.Clear()
    $stsSummary.Text = 'Not loaded.'

    # RESOLVED ON LOAD, NOT ON THE SWITCH. Finding the on-premises object is
    # two or three cheap reads, and doing it here means flipping the switch is
    # instant and, more to the point, that the switch can be greyed out with a
    # reason BEFORE somebody flips it and finds out the object is not there.
    if ($script:Mode -eq 'OnPrem' -or $cmbMode.Enabled) {
        Set-EditorStatus -Text "Resolving $t on-premises..."
        try { $script:OnPrem = Resolve-OnPremMailboxTarget -Identity $t } catch { $script:OnPrem = $null }
    }

    Update-EditorIdentity

    $script:TabReadFailed = $false
    Update-OverviewTab
    if (-not $script:TabReadFailed) { $script:Loaded['Overview'] = $true }
    # The tab that is already showing has to be loaded too -- the operator may
    # have been on Access when they pressed Load.
    Invoke-EditorTabLoad -Key "$($tabs.SelectedTab.Name)"
    Set-EditorStatus -Text "Loaded $($m.Smtp)."
}


$cmbMode.Add_SelectedIndexChanged({
    # A SILENT EARLY RETURN WOULD DESYNC THE BOX FROM THE WINDOW. Show-ToolDialog
    # pumps messages while Busy is set, so this handler can genuinely arrive
    # mid-action; putting the selection back is what stops the drop-down
    # saying one directory while every Apply writes to the other.
    $want = $(if ($cmbMode.SelectedIndex -eq 1) { 'OnPrem' } else { 'Cloud' })
    if ($script:Busy) {
        $cmbMode.SelectedIndex = $(if ($script:Mode -eq 'OnPrem') { 1 } else { 0 })
        return
    }
    if ($want -eq $script:Mode) { return }
    Invoke-EditorAction { Set-EditorMode -Mode $want } 'Switch editing mode'
})

$btnLoad.Add_Click({ Invoke-EditorAction { Invoke-EditorLoad } 'Load mailbox' })
$txtMbx.Add_KeyDown({ if ($_.KeyCode -eq 'Return') { $btnLoad.PerformClick(); $_.SuppressKeyPress = $true } })

# TAB CHANGE AND WINDOW CLOSE ARE WIRED BY THE SHARED RUNTIME, guards included.
Register-EditorTabHandler
#endregion

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Mailbox Editor' | Out-Null

$form.Add_Shown({
    # THE SWITCH IS PROVED BEFORE IT IS OFFERED. A drop-down that lets you
    # pick On-premises and then fails on every Apply is worse than one that is
    # greyed out with the reason sitting next to it.
    #
    # NOTHING TO PROVE IN A CLOUD-ONLY TENANT. The switch is not on screen
    # there -- New-EditorModeSwitch replaced it with the words 'Cloud Tenant'
    # -- so announcing that on-premises editing is unavailable would be a
    # standing complaint about a directory this tenant never had.
    if (-not (Test-TenantIsCloudOnly)) {
        $ready = Test-OnPremMailboxReady
        $cmbMode.Enabled = $ready.Ok
        $script:OnPremWhy = $(if ($ready.Ok) { '' } else { $ready.Reason })
        if (-not $ready.Ok) {
            # FOUND BY NAME, not held in a variable. The label belongs to the
            # shared builder now; looking it up means this window cannot go on
            # referring to one it no longer creates.
            $lm = @($form.Controls.Find('lblEditorMode', $true))
            if ($lm.Count) { $lm[0].ForeColor = (Get-ToolStatusColour -Kind Muted) }
            Set-EditorStatus -Text "On-premises editing is unavailable: $($ready.Reason)"
        }
    }
    if ("$($txtMbx.Text)".Trim()) { $btnLoad.PerformClick() }
})


[void]$form.ShowDialog()

# DISPOSE IS WRAPPED, AND THAT IS NOT LAZINESS.
#
# ShowDialog has already returned, so the window is shut and every write this
# tool was going to make is made. Dispose here is returning handles to the OS.
# If .NET's own TabControl teardown indexes past the end of its page
# collection -- which it does, and which the guards above make far less likely
# but cannot make impossible -- the operator should not be shown a stack trace
# for a window they already closed. The process is ending; the handles go back
# either way.
try { $form.Dispose() } catch { }
