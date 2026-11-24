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
#   Overview     Who it is, what type, size and quota, archive, last logon,
#                and the directory fields (title, department, manager).
#   Addresses    Every SMTP address, which is primary, and the ability to
#                change it.
#   Access       Full Access, Send As, Send on Behalf -- who has them, add and
#                remove -- plus the two 'keep a copy in Sent Items' settings
#                and the automap repair.
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
# Replies, Recovery -- are greyed out.
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
# Set by any tab read that CATCHES a failure instead of throwing. Without it a
# tab whose data could not be read is cached as loaded and stays empty.
$script:TabReadFailed = $false

# WHICH DIRECTORY THIS WINDOW IS WRITING TO. 'Cloud' or 'OnPrem'. Every gate,
# every Apply and the identity banner read it, so there is one answer and it
# is on screen.
$script:Mode       = 'Cloud'
$script:OnPrem     = $null   # Resolve-OnPremMailboxTarget result
$script:OnPremWhy  = ''      # why the switch is unavailable, if it is
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

# THE EDITING SWITCH. A drop-down rather than a toggle button, because a
# button that says 'Cloud' is ambiguous -- is that what you are in, or what
# you will get if you press it? A list shows the current mode and the
# alternative at the same time and cannot be misread.
$lblMode          = New-Object System.Windows.Forms.Label
$lblMode.Text     = 'Editing:'
$lblMode.Location = New-Object System.Drawing.Point(660, ($y + 3))
$lblMode.Size     = New-Object System.Drawing.Size(55, 18)
$form.Controls.Add($lblMode)

$cmbMode          = New-Object System.Windows.Forms.ComboBox
$cmbMode.DropDownStyle = 'DropDownList'
$cmbMode.Location = New-Object System.Drawing.Point(718, $y)
$cmbMode.Size     = New-Object System.Drawing.Size(220, 24)
[void]$cmbMode.Items.Add('Cloud (Exchange Online)')
[void]$cmbMode.Items.Add('On-premises (AD / Exchange)')
$cmbMode.SelectedIndex = 0
$form.Controls.Add($cmbMode)

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
$lblIdent.ForeColor = [System.Drawing.Color]::DimGray
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
$lblSync.ForeColor = [System.Drawing.Color]::DimGray
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
$lblStatus.ForeColor = [System.Drawing.Color]::DimGray
$lblStatus.Text      = 'Enter a mailbox and press Load.'
$form.Controls.Add($lblStatus)

$form.ClientSize = New-Object System.Drawing.Size(1120, ($lblStatus.Bottom + 12))

function New-EditorTab {
    param([string]$Key, [string]$Text)
    $p = New-Object System.Windows.Forms.TabPage
    $p.Text    = $Text
    $p.Name    = $Key
    $p.Padding = New-Object System.Windows.Forms.Padding(8)
    $p.AutoScroll = $true
    [void]$tabs.TabPages.Add($p)
    return $p
}

$tabOverview  = New-EditorTab -Key 'Overview'  -Text 'Overview'
# SECOND, because the question after 'who is this' is 'is anything wrong with
# it'. Health carries only what no other tab shows.
$tabHealth    = New-EditorTab -Key 'Health'    -Text 'Health'
$tabAddresses = New-EditorTab -Key 'Addresses' -Text 'Addresses'
$tabAccess    = New-EditorTab -Key 'Access'    -Text 'Access'
$tabFolders   = New-EditorTab -Key 'Folders'   -Text 'Folder Permissions'
$tabMailFlow  = New-EditorTab -Key 'MailFlow'  -Text 'Mail Flow'
$tabSettings  = New-EditorTab -Key 'Settings'  -Text 'Settings'
$tabRetention = New-EditorTab -Key 'Retention' -Text 'Retention'
$tabReplies   = New-EditorTab -Key 'Replies'   -Text 'Automatic Replies'
$tabRecovery  = New-EditorTab -Key 'Recovery'  -Text 'Recovery'
# LAST, AND IT LOADS ON A BUTTON. Per-folder statistics are the slowest read in
# this window by a wide margin, and nothing else should ever wait behind them.
$tabStats     = New-EditorTab -Key 'Stats'     -Text 'Statistics'
#endregion

#region Shell helpers -----------------------------------------------------------

function Set-EditorStatus {
    param([string]$Text, [switch]$Bad)
    $lblStatus.Text = $Text
    $lblStatus.ForeColor = $(if ($Bad) { [System.Drawing.Color]::FromArgb(180, 0, 0) } else { [System.Drawing.Color]::DimGray })
    $lblStatus.Refresh()
}

function Update-EditorIdentity {
    <#
        Repaint the identity and sync banners from $script:Mbx.

        A FUNCTION rather than two lines inside Invoke-EditorLoad, because
        several Applies re-resolve the mailbox -- making an alias primary
        changes the address in the banner, converting the type changes the
        type in it. Left inline, the banner kept showing what was true when
        Load was pressed while every button acted on something else.
    #>
    $m = $script:Mbx
    if (-not $m -or -not $m.Ok) { $lblIdent.Text = 'No mailbox loaded.'; $lblSync.Text = ''; return }

    # ON-PREMISES MODE GETS ITS OWN BANNER, in its own colour. The whole point
    # of the switch is that you can tell at a glance which directory an Apply
    # is about to write to; a banner that looked the same in both modes would
    # give that away.
    if ($script:Mode -eq 'OnPrem') {
        $o = $script:OnPrem
        if (-not $o -or -not $o.Ok) {
            $lblIdent.Text = "ON-PREMISES -- $($m.Smtp) was not found on-premises."
            $lblIdent.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
            $lblSync.Text = "$(if ($o) { $o.Reason } else { '' })"
            $lblSync.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
            return
        }
        $lblIdent.Text = "ON-PREMISES:  $($o.Display)   <$($o.Smtp)>   $($o.Kind)"
        $lblIdent.ForeColor = [System.Drawing.Color]::FromArgb(150, 80, 0)
        $lblSync.Text = "Writing to Active Directory via $($o.SetCmdlet). $($script:OnPremSyncNotice)"
        $lblSync.ForeColor = [System.Drawing.Color]::FromArgb(150, 80, 0)
        return
    }

    $lblIdent.Text = "$($m.Display)   <$($m.Smtp)>   $($m.Type)"
    $lblIdent.ForeColor = [System.Drawing.Color]::FromArgb(0, 110, 60)
    if ($m.IsSynced) {
        $lblSync.Text = $(if ($script:OnPremWhy) {
            "SYNCED FROM ACTIVE DIRECTORY -- addresses, display name, alias and hide-from-GAL are mastered on-premises. On-premises editing is unavailable here: $($script:OnPremWhy)"
        } else {
            'SYNCED FROM ACTIVE DIRECTORY -- addresses, display name, alias and hide-from-GAL are mastered on-premises. Set Editing (top right) to On-premises to change them.'
        })
        $lblSync.ForeColor = [System.Drawing.Color]::FromArgb(190, 110, 0)
    } else {
        $lblSync.Text = 'Cloud-only mailbox -- every field on every tab is editable here.'
        $lblSync.ForeColor = [System.Drawing.Color]::DimGray
    }
}

function Invoke-EditorAction {
    <#
        THE WRAPPER EVERY APPLY AND EVERY LOAD GOES THROUGH.

        One busy flag for the whole window, not one per tab: a second action
        started while the first is mid-write is how a grid gets rebuilt under
        a loop that is still reading it.
    #>
    param([scriptblock]$Do, [string]$Context)
    if ($script:Busy) { return }
    $script:Busy = $true
    try {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        & $Do
    } catch {
        Set-EditorStatus -Text "$Context failed: $($_.Exception.Message)" -Bad
        # AUDITED HERE, not only by Show-HandlerError. That logs the Context
        # as the target ('Set forwarding'), which cannot be tied back to a
        # mailbox six weeks later -- the exact question an audit trail exists
        # to answer. This line carries the address.
        Write-EditorAudit -Action 'MailboxEditorAction' -Result 'Failed' `
            -Detail "$Context failed: $($_.Exception.Message)"
        Show-HandlerError -ErrorRecord $_ -Context $Context
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $script:Busy = $false
    }
}

function Test-EditorLoaded {
    # Every Apply calls this first. Without it, a button pressed before Load
    # acts on $null and the error names a property rather than the problem.
    if (-not $script:Mbx -or -not $script:Mbx.Ok) {
        Set-EditorStatus -Text 'Load a mailbox first.' -Bad
        return $false
    }
    return $true
}

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
        Set-EditorStatus -Text "$What is an Exchange Online operation. Set Editing back to Cloud." -Bad
        return $false
    }
    return $true
}

function Add-GridPickColumn {
    <#
        A tick box in front of every row.

        WHY NOT JUST USE THE SELECTION. Row selection is invisible the moment
        the grid loses focus, it is lost when the grid is re-read, and it has
        no way to say "these four, spread down a list of thirty". A tick
        survives all three, and it is what the operator means by "selected".

        THE GRID HAS TO STOP BEING ReadOnly for a checkbox column to be
        clickable at all -- so every OTHER column is made read-only
        individually. Leaving the whole grid writable would let somebody type
        over a user name and think they had renamed something.

        THE DIRTY-STATE COMMIT IS NOT OPTIONAL. Without it a DataGridView
        checkbox does not commit until the cell loses focus, so ticking a box
        and immediately pressing a button acts on the state BEFORE the tick.
        That reads as the button ignoring the row you just ticked.
    #>
    param($Grid)
    $c = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $c.Name       = 'Pick'
    $c.HeaderText = ''
    $c.Width      = 30
    $Grid.Columns.Insert(0, $c)
    $Grid.ReadOnly = $false
    foreach ($col in $Grid.Columns) { if ($col.Name -ne 'Pick') { $col.ReadOnly = $true } }
    $Grid.Add_CurrentCellDirtyStateChanged({
        if ($this.IsCurrentCellDirty) { $this.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) }
    })
}

function Get-GridPicked {
    # The Tag of every ticked row. One definition, so 'selected' means the same
    # thing on every tab and in every button.
    param($Grid)
    $out = @()
    foreach ($r in $Grid.Rows) {
        if ($r.Tag -and [bool]$r.Cells['Pick'].Value) { $out += $r.Tag }
    }
    return @($out)
}

function Set-GridPicked {
    param($Grid, [bool]$On)
    foreach ($r in $Grid.Rows) { $r.Cells['Pick'].Value = $On }
}

function New-DialogCheck {
    # One checkbox on a dialog. Text, position, and nothing else -- the callers
    # differ only in what they ask.
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W = 260, [bool]$Checked = $false)
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Text     = $Text
    $c.Location = New-Object System.Drawing.Point($X, $Y)
    $c.Size     = New-Object System.Drawing.Size($W, 20)
    $c.Checked  = $Checked
    $Parent.Controls.Add($c)
    return $c
}

function New-DialogList {
    # A multiline box for a list of addresses, one per line.
    param($Parent, [int]$X, [int]$Y, [int]$W, [int]$H, [string]$Initial = '')
    $t = New-Object System.Windows.Forms.TextBox
    $t.Multiline  = $true
    $t.ScrollBars = 'Vertical'
    $t.WordWrap   = $false
    $t.Font       = New-Object System.Drawing.Font('Consolas', 9)
    $t.Location   = New-Object System.Drawing.Point($X, $Y)
    $t.Size       = New-Object System.Drawing.Size($W, $H)
    $t.Text       = $Initial
    $Parent.Controls.Add($t)
    return $t
}

function ConvertFrom-DialogList {
    # Addresses out of a multiline box. Commas and semicolons are accepted as
    # well as newlines, because half of what gets pasted in comes from a
    # ticket, not from a list.
    param([string]$Text)
    $out = @()
    foreach ($line in ("$Text" -split "[`r`n,;]")) {
        $t = "$line".Trim()
        if ($t) { $out += $t }
    }
    return @($out | Select-Object -Unique)
}

function Show-ResultListDialog {
    <#
        What happened, per target.

        A MULTI-TARGET ACTION CANNOT REPORT 'done'. Four mailboxes where one
        failed is not a success and not a failure, and a status line saying
        'Applied' would be a lie about the one that did not. This lists every
        target with its own outcome.
    #>
    param([string]$Title, [string[]]$Lines)
    $f = New-Object System.Windows.Forms.Form
    $f.Text          = $Title
    $f.ClientSize    = New-Object System.Drawing.Size(720, 420)
    $f.StartPosition = 'CenterParent'
    $f.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $box = New-Object System.Windows.Forms.TextBox
    $box.Multiline  = $true
    $box.ReadOnly   = $true
    $box.ScrollBars = 'Both'
    $box.WordWrap   = $false
    $box.Font       = New-Object System.Drawing.Font('Consolas', 9)
    $box.Location   = New-Object System.Drawing.Point(12, 12)
    $box.Size       = New-Object System.Drawing.Size(696, 356)
    $box.Text       = (@($Lines) -join "`r`n")
    $box.Tag        = 'NoTheme'
    $box.BackColor  = [System.Drawing.Color]::WhiteSmoke
    $box.ForeColor  = [System.Drawing.Color]::Black
    $f.Controls.Add($box)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text         = 'Close'
    $ok.Location     = New-Object System.Drawing.Point(608, 378)
    $ok.Size         = New-Object System.Drawing.Size(100, 28)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $f.Controls.Add($ok)
    $f.AcceptButton = $ok
    Set-ConsoleTheme -Form $f
    [void]$f.ShowDialog()
    $f.Dispose()
}

function Write-EditorAudit {
    # One line per action, one shape. Result is Success unless the caller says
    # otherwise -- a failed write throws before it reaches here.
    param([string]$Action, [string]$Detail, [string]$Result = 'Success')
    try {
        # Environment is the MODE, not a constant. An address change made
        # on-premises and one made in the cloud are different acts with
        # different blast radii, and the log has to tell them apart.
        $env = $(if ($script:Mode -eq 'OnPrem') { 'OnPrem' } else { 'Cloud' })
        Write-CompletedLog -Action $Action -Target (Get-LogTarget $script:Mbx.Smtp) `
            -Result $Result -Detail $Detail -Environment $env
    } catch { }
}

# --- small layout helpers, so every tab looks the same ------------------------
function New-TabLabel {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W = 160, [switch]$Bold, [switch]$Muted)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.Size = New-Object System.Drawing.Size($W, 18)
    if ($Bold)  { $l.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold) }
    if ($Muted) { $l.ForeColor = [System.Drawing.Color]::DimGray }
    $Parent.Controls.Add($l)
    return $l
}

function New-TabValue {
    # A read-only value. A TextBox rather than a Label so it can be selected
    # and copied -- half of what an operator does with this window is paste a
    # value into a ticket.
    param($Parent, [int]$X, [int]$Y, [int]$W = 300)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location  = New-Object System.Drawing.Point($X, ($Y - 3))
    $t.Size      = New-Object System.Drawing.Size($W, 22)
    $t.ReadOnly  = $true
    $t.BorderStyle = 'None'
    $t.Tag       = 'NoTheme'
    $t.BackColor = [System.Drawing.Color]::WhiteSmoke
    $t.ForeColor = [System.Drawing.Color]::Black
    $Parent.Controls.Add($t)
    return $t
}

function New-TabButton {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W = 130)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.Location = New-Object System.Drawing.Point($X, $Y)
    $b.Size = New-Object System.Drawing.Size($W, 28)
    $Parent.Controls.Add($b)
    return $b
}

function New-TabGrid {
    param($Parent, [object[]]$Columns, [int]$X, [int]$Y, [int]$W, [int]$H)
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Location            = New-Object System.Drawing.Point($X, $Y)
    $g.Size                = New-Object System.Drawing.Size($W, $H)
    $g.ReadOnly            = $true
    $g.AllowUserToAddRows  = $false
    $g.SelectionMode       = 'FullRowSelect'
    $g.MultiSelect         = $true
    $g.RowHeadersVisible   = $false
    $g.AutoSizeColumnsMode = 'None'
    $g.ScrollBars          = 'Both'
    $g.Font                = New-Object System.Drawing.Font('Consolas', 9)
    foreach ($c in $Columns) {
        $i = $g.Columns.Add($c[0], $c[0])
        $g.Columns[$i].Width    = $c[1]
        $g.Columns[$i].SortMode = 'Automatic'
    }
    $Parent.Controls.Add($g)
    return $g
}
#endregion

#region Overview tab ------------------------------------------------------------
$ovY = 8
[void](New-TabLabel $tabOverview 'IDENTITY' 8 $ovY -Bold)
$ovY += 22
[void](New-TabLabel $tabOverview 'Display name:' 8 $ovY)
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
[void](New-TabLabel $tabOverview 'Hidden from GAL:' 8 $ovY)
$ovHidden = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Source:' 560 $ovY 135)
$ovSource = New-TabValue $tabOverview 700 $ovY 320
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
$ovY += 30

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
$ovY += 28

# HIDDEN IN CLOUD MODE. Title, department, office and manager are directory
# properties; in the cloud they are read-only here on purpose, because writing
# them to a synced object is the silent-revert trap this whole window is built
# around. In on-premises mode the same four boxes become editable and this
# appears under them.
$btnDirApply = New-TabButton $tabOverview 'Apply directory fields' 170 $ovY 170
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
[void](New-TabLabel $tabOverview 'Last logon:' 8 $ovY)
$ovLogon = New-TabValue $tabOverview 170 $ovY 360
$ovLogonNote = New-TabLabel $tabOverview '' 560 $ovY 460 -Muted
$ovY += 30

$btnOvRefresh = New-TabButton $tabOverview 'Refresh size' 8 $ovY 130
$ovY += 36

$ovNotes            = New-Object System.Windows.Forms.TextBox
$ovNotes.Multiline  = $true
$ovNotes.ReadOnly   = $true
$ovNotes.ScrollBars = 'Vertical'
$ovNotes.WordWrap   = $true
$ovNotes.Location   = New-Object System.Drawing.Point(8, $ovY)
$ovNotes.Size       = New-Object System.Drawing.Size(1050, 150)
$ovNotes.Font       = New-Object System.Drawing.Font('Consolas', 9)
$ovNotes.Tag        = 'NoTheme'
$ovNotes.BackColor  = [System.Drawing.Color]::WhiteSmoke
$ovNotes.ForeColor  = [System.Drawing.Color]::Black
$tabOverview.Controls.Add($ovNotes)

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
    $r = [PSCustomObject]@{ Title = ''; Department = ''; Manager = ''; Office = ''; Source = ''; Note = '' }
    # Get-User is the Exchange view of the directory object and works for both
    # cloud and synced mailboxes without needing AD or Graph to be connected.
    try {
        $u = Invoke-EXOCommand -Name 'Get-User' -Parameters @{ Identity = $Mbx.Smtp; ErrorAction = 'Stop' }
        $r.Title      = "$($u.Title)"
        $r.Department = "$($u.Department)"
        $r.Office     = "$($u.Office)"
        $r.Manager    = "$($u.Manager)"
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

function Set-EditorDirectoryEditable {
    # The four directory boxes are the same controls in both modes; only their
    # ReadOnly flag and the Apply button move. Two sets of controls would mean
    # two places for the values to disagree.
    param([bool]$On)
    foreach ($t in @($ovTitle, $ovDept, $ovMgr, $ovOffice)) {
        $t.ReadOnly    = -not $On
        $t.BorderStyle = $(if ($On) { 'FixedSingle' } else { 'None' })
        $t.BackColor   = $(if ($On) { [System.Drawing.Color]::White } else { [System.Drawing.Color]::WhiteSmoke })
    }
    $btnDirApply.Visible = $On
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
    $ovUpn.Text    = "$(if ($u) { $u.UserPrincipalName } else { '' })"
    $ovGuid.Text   = "$($o.Raw.Guid)"
    $ovHidden.Text = $(if ($o.Hidden) { 'YES -- hidden' } else { 'no' })
    $ovSource.Text = 'Active Directory (on-premises)'

    $ovTitle.Text  = "$(if ($u) { $u.Title } else { '' })"
    $ovDept.Text   = "$(if ($u) { $u.Department } else { '' })"
    $ovMgr.Text    = "$(if ($u) { $u.Manager } else { '' })"
    $ovOffice.Text = "$(if ($u) { $u.Office } else { '' })"
    $ovDirNote.Text = '-- editable here; these are the Active Directory values.'
    $ovDirNote.ForeColor = [System.Drawing.Color]::DimGray

    $ovLicence.Text = '(switch to Cloud to read the licence)'
    foreach ($t in @($ovSize, $ovItems, $ovQuota, $ovLogon)) { $t.Text = '(Exchange Online only)' }
    $ovArchive.Text = '(Exchange Online only)'
    $ovLogonNote.Text = 'Size, item count and last logon live on the cloud mailbox. Switch Editing to Cloud to see them.'

    $ovNotes.Text = "$($o.Reason)`r`n$($script:OnPremSyncNotice)"
    Set-EditorDirectoryEditable -On ($null -ne $u)
    if (-not $u) {
        $ovDirNote.Text = '-- no on-premises user object was returned, so the directory fields cannot be edited.'
        $ovDirNote.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
    }
}

function Update-OverviewTab {
    if ($script:Mode -eq 'OnPrem') { Update-OverviewOnPrem; return }
    Set-EditorDirectoryEditable -On $false
    $m = $script:Mbx
    $mb = $m.Raw
    $ovName.Text   = "$($mb.DisplayName)"
    $ovType.Text   = "$($mb.RecipientTypeDetails)"
    $ovSmtp.Text   = "$($mb.PrimarySmtpAddress)"
    $ovAlias.Text  = "$($mb.Alias)"
    $ovUpn.Text    = "$($mb.UserPrincipalName)"
    $ovGuid.Text   = "$($mb.ExternalDirectoryObjectId)"
    $ovHidden.Text = $(if ($mb.HiddenFromAddressListsEnabled) { 'YES -- hidden' } else { 'no' })
    $ovSource.Text = $(if ($m.IsSynced) { 'Synced from Active Directory' } else { 'Cloud only' })

    $d = Get-EditorDirectoryFields -Mbx $m
    $ovTitle.Text  = $d.Title
    $ovDept.Text   = $d.Department
    $ovMgr.Text    = $d.Manager
    $ovOffice.Text = $d.Office
    $ovDirNote.Text = $(if ($m.IsSynced) {
        '-- mastered in Active Directory for this mailbox; change them on-premises or in AD Lookup.'
    } else {
        '-- change these in the Attributes editor (Unified Lookup) or AD Lookup.'
    })

    $notes = New-Object System.Text.StringBuilder
    if ($d.Note) { [void]$notes.AppendLine($d.Note) }

    # Statistics are a separate, slower call. Fetched here rather than on open
    # so the identity block appears immediately on a large mailbox.
    $ovSize.Text = '(reading...)'; $ovItems.Text = ''; $ovQuota.Text = ''; $ovLogon.Text = ''
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

    if ($notes.Length -eq 0) { [void]$notes.Append('No problems reading this mailbox.') }
    $ovNotes.Text = $notes.ToString()
}

$btnDirApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if ($script:Mode -ne 'OnPrem') { Set-EditorStatus -Text 'Directory fields are only editable in on-premises mode.' -Bad; return }
        if (-not (Test-EditorFieldWritable -Field 'DisplayName')) { return }
        $u = $script:OnPrem.User
        if (-not $u) { Set-EditorStatus -Text 'There is no on-premises user object to write to.' -Bad; return }

        # ONLY WHAT CHANGED. Set-User with all four every time overwrites a
        # field somebody else edited between this window loading and the
        # button being pressed, and claims four changes in the audit line for
        # the one the operator made.
        $p = @{ Identity = "$($u.Identity)"; ErrorAction = 'Stop' }
        $changed = @()
        foreach ($f in @(
            @{ Name = 'Title';      Box = $ovTitle;  Was = "$($u.Title)" }
            @{ Name = 'Department'; Box = $ovDept;   Was = "$($u.Department)" }
            @{ Name = 'Office';     Box = $ovOffice; Was = "$($u.Office)" }
            @{ Name = 'Manager';    Box = $ovMgr;    Was = "$($u.Manager)" }
        )) {
            $now = "$($f.Box.Text)".Trim()
            if ($now -ne "$($f.Was)".Trim()) {
                $p[$f.Name] = $(if ($now) { $now } else { $null })
                $changed += "$($f.Name)='$now'"
            }
        }
        if ($changed.Count -eq 0) { Set-EditorStatus -Text 'Nothing changed.'; return }

        Invoke-OnPremCommand -Name 'Set-User' -Parameters $p | Out-Null
        Write-EditorAudit -Action 'MailboxDirectoryFields' -Detail ($changed -join '; ')
        Invoke-EditorReload
        Update-OverviewTab
        Set-EditorStatus -Text ("Applied on-premises: $($changed -join '; ')." + (Get-EditorSyncSuffix))
    } 'Apply directory fields'
})

$btnOvRefresh.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction { Update-OverviewTab; Set-EditorStatus -Text 'Overview refreshed.' } 'Refresh overview'
})
#endregion

#region Addresses tab -----------------------------------------------------------
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
            $gridAddr.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::DimGray
        }
    }
    if ($onPrem) {
        $can = [bool]$script:OnPrem.SetCmdlet
        $adNote.Text = "$($gridAddr.Rows.Count) on-premises address(es). The primary is in bold; X500 entries are what let replies to old mail resolve and are never removed automatically."
        $adWarn.Text = $(if ($can) { "Changes are written with $($script:OnPrem.SetCmdlet). $($script:OnPremSyncNotice)" }
                         else { 'This on-premises object has no mailbox, so it has no address list to change.' })
        $adWarn.ForeColor = [System.Drawing.Color]::FromArgb(150, 80, 0)
        $btnAddrPrimary.Enabled = $can; $btnAddrAdd.Enabled = $can; $btnAddrRemove.Enabled = $can
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
        $adWarn.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
        $btnAddrPrimary.Enabled = $false; $btnAddrAdd.Enabled = $false; $btnAddrRemove.Enabled = $false
    }
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

#region Access tab --------------------------------------------------------------
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
$btnAccAutomap = New-TabButton $tabAccess 'Fix automapping...' 720 $acBtnY 170
$acBtnY += 34
$btnAccAll     = New-TabButton $tabAccess 'Tick all'            720 $acBtnY 82
$btnAccNone    = New-TabButton $tabAccess 'None'                808 $acBtnY 82
$acBtnY += 40
$acHint = New-TabLabel $tabAccess '' 720 $acBtnY 340 -Muted
$acHint.Size = New-Object System.Drawing.Size(340, 100)
$acHint.Text = ('Full Access opens the mailbox. Send As sends from it as the mailbox. ' +
                'Send on Behalf sends showing both names.' + "`r`n`r`n" +
                'Change and Remove act on every TICKED row, not on the highlighted one.')
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
        if ($e.Detail -eq 'DENY') { $gridAccess.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0) }
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
    $note.ForeColor = [System.Drawing.Color]::DimGray
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
            Mailboxes = (ConvertFrom-DialogList -Text $txtMbxList.Text)
            Trustees  = (ConvertFrom-DialogList -Text $txtWhoList.Text)
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
    $note.ForeColor = [System.Drawing.Color]::DimGray
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

$btnAccAutomap.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Fixing automapping')) { return }
        $picked = @(Get-GridPicked -Grid $gridAccess | Where-Object { $_.Kind -eq 'FullAccess' })
        if ($picked.Count -eq 0) { Set-EditorStatus -Text 'Tick one or more Full Access rows to repair.' -Bad; return }
        $who = @($picked | ForEach-Object { "$($_.User)" } | Select-Object -Unique)
        $a = Show-ToolDialog(
            ("Repair automapping on $($script:Mbx.Display) for:" + "`r`n  " + ($who -join "`r`n  ") + "`r`n`r`n" +
             'Yes -- make it open automatically in their Outlook.' + "`r`n" +
             'No  -- stop it opening automatically, but KEEP their access.' + "`r`n`r`n" +
             'Either way Full Access is removed and re-added, so there are a few seconds where they have no access.'),
            'Fix automapping', 'YesNoCancel', 'Question')
        if ($a -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
        $auto = ($a -eq [System.Windows.Forms.DialogResult]::Yes)
        $lines = @(); $ok = 0; $bad = 0
        foreach ($u in $who) {
            try { Repair-MailboxAutoMapping -Smtp $script:Mbx.Smtp -Trustee $u -AutoMap $auto; $lines += "OK      $u"; $ok++ }
            catch { $lines += "FAILED  $u`: $($_.Exception.Message)"; $bad++ }
        }
        Write-EditorAudit -Action 'MailboxAutomapFix' `
            -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
            -Detail "Full Access restamped with automapping $auto for $($who -join '; ')"
        if ($bad) { Show-ResultListDialog -Title "Fix automapping -- $ok done, $bad failed" -Lines $lines }
        Update-AccessTab
        Set-EditorStatus -Text "Automapping is now $(if ($auto) { 'on' } else { 'off' }) for $ok, failed $bad. Outlook can take a while to notice." -Bad:($bad -gt 0)
    } 'Fix automapping'
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


#region Folder Permissions tab --------------------------------------------------
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
$fpNote.Text = 'Send on Behalf is the other half of delegation and lives on the Access tab -- it is a mailbox property, not a folder right.'
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
                $gridFolder.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(190, 110, 0)
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
    $note.ForeColor = [System.Drawing.Color]::DimGray
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
            Mailboxes = (ConvertFrom-DialogList -Text $txtMbxList.Text)
            Trustees  = (ConvertFrom-DialogList -Text $txtWhoList.Text)
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
    MAILBOX TYPE, HIDE FROM THE ADDRESS LIST, ARCHIVE.

    The three that used to be three separate tools, and the tab where the
    hybrid rule bites hardest: hide-from-GAL is directory-mastered, so on a
    synced mailbox it is disabled here with the reason rather than accepted
    and silently reverted half an hour later. Hide-Unhide.ps1 learned that the
    hard way and this inherits the lesson.
#>
$stY = 8
[void](New-TabLabel $tabSettings 'MAILBOX TYPE' 8 $stY -Bold)
$stY += 22
$stTypeNow = New-TabLabel $tabSettings '' 8 $stY 500 -Muted
$stY += 24
[void](New-TabLabel $tabSettings 'Change to:' 8 ($stY + 3) 90)
$cmbType               = New-Object System.Windows.Forms.ComboBox
$cmbType.DropDownStyle = 'DropDownList'
$cmbType.Location      = New-Object System.Drawing.Point(104, $stY)
$cmbType.Size          = New-Object System.Drawing.Size(180, 24)
foreach ($t in @('Regular', 'Shared', 'Room', 'Equipment')) { [void]$cmbType.Items.Add($t) }
$cmbType.SelectedIndex = 0
$tabSettings.Controls.Add($cmbType)
$btnTypeApply = New-TabButton $tabSettings 'Convert' 294 ($stY - 2) 110
$stY += 30
$stTypeNote = New-TabLabel $tabSettings '' 8 $stY 1050 -Muted
$stTypeNote.Size = New-Object System.Drawing.Size(1050, 50)
$stTypeNote.Text = ('Converting to Shared does not remove the licence -- a shared mailbox under 50GB does not need one, but taking it away is a separate step.' + "`r`n" +
                    'A shared mailbox cannot have its meetings transferred by Change Meeting Organizer, so do that BEFORE converting a leaver.')
# KEPT, because on-premises mode overwrites this label with its own sentence
# and the cloud path has to be able to put the real one back. A label that
# silently keeps the other mode's text is how a warning goes missing.
$script:StTypeNoteText = $stTypeNote.Text
$stY += 60

[void](New-TabLabel $tabSettings 'ADDRESS LIST' 8 $stY -Bold)
$stY += 22
$chkHidden          = New-Object System.Windows.Forms.CheckBox
$chkHidden.Text     = 'Hidden from the global address list'
$chkHidden.Location = New-Object System.Drawing.Point(8, $stY)
$chkHidden.Size     = New-Object System.Drawing.Size(320, 20)
$tabSettings.Controls.Add($chkHidden)
$btnHiddenApply = New-TabButton $tabSettings 'Apply' 340 ($stY - 5) 100
$stY += 26
$stHideNote = New-TabLabel $tabSettings '' 8 $stY 1050 -Muted
$stHideNote.Size = New-Object System.Drawing.Size(1050, 34)
$stY += 44

[void](New-TabLabel $tabSettings 'ARCHIVE' 8 $stY -Bold)
$stY += 22
$stArchNow = New-TabLabel $tabSettings '' 8 $stY 500 -Muted
$stY += 24
$btnArchEnable = New-TabButton $tabSettings 'Enable archive' 8 $stY 150
$stY += 34
$stArchNote = New-TabLabel $tabSettings '' 8 $stY 1050 -Muted
$stArchNote.Size = New-Object System.Drawing.Size(1050, 50)
$stArchNote.Text = ('Enabling an archive is not reversible from here -- disabling one is a separate, destructive operation that is deliberately not offered.' + "`r`n" +
                    'The archive takes a while to appear in Outlook after it is created.')

function Update-SettingsTab {
    # ON-PREMISES MODE SHOWS THE ON-PREMISES VALUES. Hide-from-GAL is the one
    # thing on this tab that on-premises masters, and it is the one thing the
    # Apply writes there -- so the checkbox has to be drawn from the same
    # object the Apply will write to. Drawn from the cloud copy it could sit
    # unticked while the on-premises object is hidden, and pressing Apply
    # would then unhide something that was already visible and report success.
    if ($script:Mode -eq 'OnPrem') {
        $o = $script:OnPrem
        $ok = ($o -and $o.Ok -and $o.SetCmdlet)
        $stTypeNow.Text = "Currently: $(if ($o -and $o.Ok) { $o.Kind } else { 'not found on-premises' }) -- mailbox type is an Exchange Online setting"
        $chkHidden.Checked = [bool]($o -and $o.Hidden)
        $chkHidden.Enabled = $ok; $btnHiddenApply.Enabled = $ok
        $stHideNote.Text = $(if ($ok) { "Written on-premises with $($o.SetCmdlet). $($script:OnPremSyncNotice)" }
                             else { 'There is no on-premises mailbox object to change.' })
        $stHideNote.ForeColor = [System.Drawing.Color]::FromArgb(150, 80, 0)

        # Type conversion and archive are Exchange Online operations with no
        # on-premises equivalent for a migrated mailbox.
        $cmbType.Enabled = $false; $btnTypeApply.Enabled = $false; $btnArchEnable.Enabled = $false
        $stTypeNote.Text = 'Converting the mailbox type and enabling the archive are Exchange Online operations. Set Editing back to Cloud.'
        $stTypeNote.ForeColor = [System.Drawing.Color]::FromArgb(150, 80, 0)
        $stArchNow.Text = '(Exchange Online only)'
        return
    }
    $cmbType.Enabled = $true; $btnTypeApply.Enabled = $true
    $stTypeNote.Text = $script:StTypeNoteText
    $stTypeNote.ForeColor = [System.Drawing.Color]::DimGray
    $mb = $script:Mbx.Raw
    $stTypeNow.Text = "Currently: $($mb.RecipientTypeDetails)"
    $chkHidden.Checked = [bool]$mb.HiddenFromAddressListsEnabled

    $t = Get-MailboxWriteTarget -Field 'HiddenFromAddressListsEnabled' -IsSynced $script:Mbx.IsSynced
    if ($t.State -eq 'Ok') {
        $chkHidden.Enabled = $true; $btnHiddenApply.Enabled = $true
        $stHideNote.Text = 'Hiding a mailbox removes it from the address book. Mail sent directly to the address still delivers.'
        $stHideNote.ForeColor = [System.Drawing.Color]::DimGray
    } else {
        $chkHidden.Enabled = $false; $btnHiddenApply.Enabled = $false
        $stHideNote.Text = $t.Reason
        $stHideNote.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
    }

    $arcGuid = "$($mb.ArchiveGuid)"
    $hasArchive = ("$($mb.ArchiveStatus)" -eq 'Active') -or
                  ($arcGuid -and $arcGuid -ne '00000000-0000-0000-0000-000000000000')
    $stArchNow.Text = $(if ($hasArchive) { "Currently: enabled ($($mb.ArchiveStatus))" } else { 'Currently: no archive' })
    $btnArchEnable.Enabled = -not $hasArchive
}

$btnTypeApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorFieldWritable -Field 'MailboxType')) { return }
        $to = "$($cmbType.SelectedItem)"
        $from = "$($script:Mbx.Raw.RecipientTypeDetails)"
        # MAPPED, NOT PREFIX-MATCHED. A regular mailbox is 'UserMailbox', which
        # does not start with 'Regular', so the prefix test let a no-op through
        # and Set-Mailbox -Type Regular threw at the operator instead.
        $typeIs = @{ Regular = 'UserMailbox'; Shared = 'SharedMailbox'
                     Room = 'RoomMailbox'; Equipment = 'EquipmentMailbox' }
        if ($from -ieq "$($typeIs[$to])") { Set-EditorStatus -Text "Already a $to mailbox."; return }
        $extra = ''
        if ($to -eq 'Shared') {
            $extra = "`r`n`r`nAFTER THIS: recurring meetings this person organises can no longer be transferred -- the transfer cmdlet refuses a shared mailbox. Do that first if it is needed."
        }
        if ($from -imatch 'Shared' -and $to -eq 'Regular') {
            $extra = "`r`n`r`nA regular mailbox needs a licence. Assign one, or mail will stop after the grace period."
        }
        $a = Show-ToolDialog(
            ("Convert $($script:Mbx.Display) from $from to $to?" + $extra),
            'Convert mailbox type', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
            Identity = $script:Mbx.Smtp; Type = $to; ErrorAction = 'Stop'
        } | Out-Null
        Write-EditorAudit -Action 'MailboxTypeConvert' -Detail "Converted from $from to $to."
        $script:Mbx = Resolve-MailboxTarget -Smtp $script:Mbx.Smtp
        Update-EditorIdentity
        Update-SettingsTab; $script:Loaded['Overview'] = $false
        Set-EditorStatus -Text "Converted to $to."
    } 'Convert mailbox type'
})

$btnHiddenApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorFieldWritable -Field 'HiddenFromAddressListsEnabled')) { return }
        Invoke-EditorSetMailbox -Parameters @{ HiddenFromAddressListsEnabled = [bool]$chkHidden.Checked }
        Write-EditorAudit -Action 'MailboxHideFromGal' -Detail "Hidden from address list = $($chkHidden.Checked)."
        Invoke-EditorReload
        Update-SettingsTab; $script:Loaded['Overview'] = $false
        Set-EditorStatus -Text ($(if ($chkHidden.Checked) { 'Hidden from the address list.' } else { 'Visible in the address list.' }) + (Get-EditorSyncSuffix))
    } 'Apply address list visibility'
})

$btnArchEnable.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
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
        Update-SettingsTab; $script:Loaded['Overview'] = $false
        Set-EditorStatus -Text 'Archive enabled. It takes a while to appear in Outlook.'
    } 'Enable archive'
})
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

#region Automatic Replies tab ---------------------------------------------------
<#
    THE MESSAGES ARE EDITED IN THE SHARED HTML EDITOR, NOT HERE.

    This tab used to hold two plain multiline text boxes. That was wrong in a
    way that did damage rather than merely looking poor: Exchange stores
    auto-reply text as HTML, and a plain box shows the markup as text and
    writes back whatever is in it -- so opening this tab on a message somebody
    had formatted in Outlook and pressing Apply flattened it. Two editors for
    one field, and the simpler one silently destroyed the other's work.

    So the editor that OOFEditor.ps1 already had -- a WebBrowser with
    contenteditable, a real toolbar, templates and a preview -- moved into
    Common.ps1 on 2026-09-18, and both this tab and that tool open the same
    one. The tab shows the current state and hands over.

    WHAT IS LEFT HERE IS A SUMMARY, and deliberately only a summary: state,
    audience, schedule. Everything editable is one button away, in the window
    that can edit it properly.
#>
$arY = 8
[void](New-TabLabel $tabReplies 'AUTOMATIC REPLIES' 8 $arY -Bold)
$arY += 24
$arNow = New-TabLabel $tabReplies '' 8 $arY 1040 -Muted
$arNow.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$arY += 26
[void](New-TabLabel $tabReplies 'External audience:' 8 $arY 150)
$arAud = New-TabValue $tabReplies 170 $arY 360
$arY += 24
[void](New-TabLabel $tabReplies 'Scheduled from:' 8 $arY 150)
$arFrom = New-TabValue $tabReplies 170 $arY 360
$arY += 24
[void](New-TabLabel $tabReplies 'Scheduled to:' 8 $arY 150)
$arTo = New-TabValue $tabReplies 170 $arY 360
$arY += 30

[void](New-TabLabel $tabReplies 'MESSAGE (internal), as plain text' 8 $arY -Bold)
$arY += 22
$arPreview            = New-Object System.Windows.Forms.TextBox
$arPreview.Multiline  = $true
$arPreview.ReadOnly   = $true
$arPreview.ScrollBars = 'Vertical'
$arPreview.WordWrap   = $true
$arPreview.Location   = New-Object System.Drawing.Point(8, $arY)
$arPreview.Size       = New-Object System.Drawing.Size(1040, 200)
$arPreview.Tag        = 'NoTheme'
$arPreview.BackColor  = [System.Drawing.Color]::WhiteSmoke
$arPreview.ForeColor  = [System.Drawing.Color]::Black
$tabReplies.Controls.Add($arPreview)
$arY += 210

$btnOofEdit = New-TabButton $tabReplies 'Edit automatic replies...' 8 $arY 200
$arY += 34
$arNote = New-TabLabel $tabReplies '' 8 $arY 1040 -Muted
$arNote.Size = New-Object System.Drawing.Size(1040, 34)
$arNote.Text = ('The editor handles the state, the schedule, the audience and both messages, with formatting, templates and a preview.' + "`r`n" +
                'The text above is the internal message with its markup stripped, so you can read it at a glance -- it is not what the sender sees.')

function ConvertFrom-OofHtml {
    <#
        HTML to something readable in a box, for DISPLAY ONLY.

        Not a parser and not trying to be. It exists so the summary is legible
        without the operator reading <div style="font-family:Calibri"> at the
        start of every line. Nothing here is ever written back -- the editor
        owns the real value -- so losing fidelity costs nothing.
    #>
    param([string]$Html)
    $t = "$Html"
    if (-not $t) { return '' }
    $t = $t -replace '(?is)<style.*?</style>', ''
    $t = $t -replace '(?i)<br\s*/?>', "`r`n"
    $t = $t -replace '(?i)</p\s*>', "`r`n"
    $t = $t -replace '(?i)</div\s*>', "`r`n"
    $t = $t -replace '(?s)<[^>]+>', ''
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    # Three or more blank lines is what stripping a nest of divs leaves behind.
    $t = $t -replace "(`r`n){3,}", "`r`n`r`n"
    return $t.Trim()
}

function Update-RepliesTab {
    $c = $null
    try {
        $c = Invoke-EXOCommand -Name 'Get-MailboxAutoReplyConfiguration' -Parameters @{
                Identity = $script:Mbx.Smtp; ErrorAction = 'Stop' }
    } catch {
        $script:TabReadFailed = $true
        $arNow.Text = "Automatic replies could not be read: $($_.Exception.Message)"
        $arNow.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0)
        return
    }
    $state = "$($c.AutoReplyState)"
    $arNow.Text = "Currently: $state"
    $arNow.ForeColor = $(if ($state -eq 'Disabled') { [System.Drawing.Color]::DimGray } else { [System.Drawing.Color]::FromArgb(0, 110, 60) })
    $arAud.Text  = "$($c.ExternalAudience)"
    $arFrom.Text = $(if ($state -eq 'Scheduled') { "$($c.StartTime)" } else { '(not scheduled)' })
    $arTo.Text   = $(if ($state -eq 'Scheduled') { "$($c.EndTime)" } else { '(not scheduled)' })
    $arPreview.Text = (ConvertFrom-OofHtml -Html "$($c.InternalMessage)")
    if (-not $arPreview.Text) { $arPreview.Text = '(no internal message set)' }
}

$btnOofEdit.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorCloudOnly -What 'Automatic replies')) { return }
        if (-not (Test-EditorFieldWritable -Field 'AutomaticReplies')) { return }
        # THE SHARED EDITOR, WITH NO LOOK-UP BAR. The mailbox is already chosen
        # here, and a second address box on the dialog is an invitation to edit
        # somebody else's replies while this window's banner names this person.
        # THE HELP PATH IS OOFEditor.ps1's, NOT THIS FILE'S. Passing our own
        # would put the whole Mailbox Editor help text under a window titled
        # 'Out of Office'.
        Show-OutOfOfficeForm -Smtp $script:Mbx.Smtp -NoTargetBar `
            -HelpScriptPath (Join-Path $PSScriptRoot 'OOFEditor.ps1')
        # The editor writes and logs its own change; this just re-reads so the
        # summary is not stale the moment the dialog closes.
        Update-RepliesTab
        Set-EditorStatus -Text 'Automatic replies re-read.'
    } 'Edit automatic replies'
})
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

      PROTOCOL ACCESS   OWA, ActiveSync, POP, IMAP, MAPI, EWS. The answer to
                        "it works in the browser but not on the phone", and
                        the thing a compromise assessment turns off.
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

[void](New-TabLabel $tabHealth 'PROTOCOL ACCESS' 8 $hlY -Bold)
$hlY += 22
[void](New-TabLabel $tabHealth 'Outlook on the web:' 8 $hlY 150)
$hlOwa   = New-TabValue $tabHealth 170 $hlY 120
[void](New-TabLabel $tabHealth 'ActiveSync:' 320 $hlY 110)
$hlEas   = New-TabValue $tabHealth 440 $hlY 120
[void](New-TabLabel $tabHealth 'MAPI (Outlook):' 590 $hlY 130)
$hlMapi  = New-TabValue $tabHealth 730 $hlY 120
$hlY += 24
[void](New-TabLabel $tabHealth 'POP:' 8 $hlY 150)
$hlPop   = New-TabValue $tabHealth 170 $hlY 120
[void](New-TabLabel $tabHealth 'IMAP:' 320 $hlY 110)
$hlImap  = New-TabValue $tabHealth 440 $hlY 120
[void](New-TabLabel $tabHealth 'EWS:' 590 $hlY 130)
$hlEws   = New-TabValue $tabHealth 730 $hlY 120
$hlY += 24
$hlProtoNote = New-TabLabel $tabHealth '' 8 $hlY 1040 -Muted
$hlProtoNote.Text = 'POP and IMAP are off by default in most tenants and are worth questioning when they are on.'
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

[void](New-TabLabel $tabHealth 'INBOX RULES' 8 $hlY -Bold)
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
    param($Value)
    if ($null -eq $Value) { return '(unknown)' }
    if ([bool]$Value) { return 'enabled' }
    return 'DISABLED'
}

function Update-HealthTab {
    foreach ($t in @($hlOwa, $hlEas, $hlMapi, $hlPop, $hlImap, $hlEws)) { $t.Text = '(reading...)' }
    $hlVerdict.Text = 'Checking...'
    $hlVerdict.Refresh()

    $m        = $script:Mbx
    $mb       = $m.Raw
    $findings = New-Object System.Collections.ArrayList
    $notes    = New-Object System.Collections.ArrayList

    # --- protocols -------------------------------------------------------
    try {
        $cas = Invoke-EXOCommand -Name 'Get-CASMailbox' -Parameters @{ Identity = $m.Smtp; ErrorAction = 'Stop' }
        $hlOwa.Text  = (Format-HealthFlag $cas.OWAEnabled)
        $hlEas.Text  = (Format-HealthFlag $cas.ActiveSyncEnabled)
        $hlMapi.Text = (Format-HealthFlag $cas.MAPIEnabled)
        $hlPop.Text  = (Format-HealthFlag $cas.PopEnabled)
        $hlImap.Text = (Format-HealthFlag $cas.ImapEnabled)
        $hlEws.Text  = (Format-HealthFlag $cas.EwsEnabled)
        if (-not $cas.MAPIEnabled)       { [void]$findings.Add('MAPI is disabled -- Outlook cannot connect at all.') }
        if (-not $cas.OWAEnabled)        { [void]$findings.Add('Outlook on the web is disabled.') }
        if (-not $cas.ActiveSyncEnabled) { [void]$findings.Add('ActiveSync is disabled -- phones will not sync.') }
        if ($cas.PopEnabled -or $cas.ImapEnabled) {
            [void]$findings.Add('POP or IMAP is enabled. Both are legacy, neither supports modern authentication well, and both are common in mailbox compromises.')
        }
    } catch {
        $script:TabReadFailed = $true
        foreach ($t in @($hlOwa, $hlEas, $hlMapi, $hlPop, $hlImap, $hlEws)) { $t.Text = '(unavailable)' }
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
        $hlGuidNote.ForeColor = [System.Drawing.Color]::DimGray
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
            $hlGuidNote.ForeColor = [System.Drawing.Color]::DimGray
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
                $hlGuidNote.ForeColor = [System.Drawing.Color]::DimGray
                if ($m.IsSynced) {
                    [void]$findings.Add('This mailbox is synced from Active Directory but has NO on-premises remote mailbox stub. On-premises mail has nowhere to route.')
                }
            } else {
                $hlEgOk.Text = '(not checked)'; $hlAgOk.Text = '(not checked)'
                $hlGuidNote.Text = 'The on-premises stub could not be read -- see the findings box.'
                $hlGuidNote.ForeColor = [System.Drawing.Color]::FromArgb(150, 80, 0)
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
            $sus = @(@($ir.Rules) | Where-Object {
                $_.Enabled -and (
                    $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo -or
                    $_.DeleteMessage -or $_.PermanentDelete -or $_.SoftDeleteMessage
                )
            } | ForEach-Object {
                $acts = @()
                if ($_.ForwardTo -or $_.ForwardAsAttachmentTo) { $acts += 'Forward' }
                if ($_.RedirectTo) { $acts += 'Redirect' }
                if ($_.DeleteMessage -or $_.PermanentDelete -or $_.SoftDeleteMessage) { $acts += 'Delete' }
                "$($_.Name)  [$($acts -join ', ')]"
            })
            if ($sus.Count) {
                $hlRuleList.Text = ($sus -join "`r`n")
                [void]$findings.Add("$($sus.Count) enabled rule(s) forward, redirect or delete mail. A rule that forwards externally and deletes the evidence is the classic compromise footprint -- check who created these.")
            } else {
                $hlRuleList.Text = 'No rule forwards, redirects or deletes.'
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
        [void]$findings.Add('Hidden from the address list. Mail sent directly still delivers. See Settings.')
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

#region Statistics tab ----------------------------------------------------------
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
            $gridStats.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(150, 80, 0)
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
    # CLEARED ON EVERY LOAD. The Statistics tab deliberately has no entry in
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

function Invoke-EditorTabLoad {
    <#
        Load a tab's data the first time it is shown.

        The cache is per LOAD, not per session: Invoke-EditorLoad clears it,
        so switching mailbox re-reads everything rather than showing the
        previous mailbox's permissions under the new name.
    #>
    param([string]$Key)
    if (-not $script:Mbx -or -not $script:Mbx.Ok) { return }

    # SIX TABS ARE CLOUD-ONLY AND ARE NOT DRAWN IN ON-PREMISES MODE. Full
    # Access, folder permissions, forwarding, litigation hold, automatic
    # replies and recoverable items are all properties of the Exchange Online
    # mailbox. Leaving the cloud values on screen under an on-premises banner
    # would be worse than blanking them: it would read as though those were
    # the on-premises settings and that an Apply would change them.
    if ($script:Mode -eq 'OnPrem' -and $script:CloudOnlyTabs -contains $Key) {
        Set-EditorStatus -Text "The $Key tab is Exchange Online only. Set Editing back to Cloud to use it." -Bad
        return
    }
    if ($script:Loaded[$Key]) { return }
    # THE FLAG IS SET AFTER THE UPDATE, AND ONLY IF IT DID NOT THROW. Setting
    # it first caches a tab whose read failed on a transient error, and the
    # tab then stays empty until the mailbox is loaded again -- looking like
    # the mailbox has nothing rather than like a call that needs retrying.
    $script:TabReadFailed = $false
    switch ($Key) {
        'Overview'  { Update-OverviewTab }
        'Addresses' { Update-AddressesTab }
        'Access'    { Update-AccessTab }
        'Folders'   { Update-FoldersTab }
        'MailFlow'  { Update-MailFlowTab }
        'Settings'  { Update-SettingsTab }
        'Retention' { Update-RetentionTab }
        'Replies'   { Update-RepliesTab }
        'Recovery'  { Update-RecoveryTab }
        'Health'    { Update-HealthTab }
        # 'Stats' IS DELIBERATELY ABSENT. Selecting the Statistics tab must not
        # start a folder enumeration -- it loads on its own button. Listing it
        # here with an empty body would look like an oversight; saying why does
        # not.
    }
    if (-not $script:TabReadFailed) { $script:Loaded[$Key] = $true }
}

# The tabs that mean nothing on-premises. One list, used by both the loader
# and the mode switch, so a tab cannot be blocked in one and live in the other.
$script:CloudOnlyTabs = @('Access', 'Folders', 'MailFlow', 'Retention', 'Replies', 'Recovery',
                         'Health', 'Stats')

function Set-EditorCloudTabsEnabled {
    <#
        DISABLED, NOT MERELY SKIPPED.

        Blocking the tab's LOAD left every button on it live: the grid still
        held rows read a moment ago for the cloud mailbox, and 'Remove
        selected' would have removed real permissions from whatever
        $script:Mbx pointed at, under a banner that said on-premises.

        The grids are cleared as well as disabled, because a disabled grid
        still showing the last mailbox's trustees is a screenshot somebody
        pastes into a ticket.
    #>
    param([bool]$On)
    foreach ($k in $script:CloudOnlyTabs) {
        $page = $tabs.TabPages | Where-Object { $_.Name -eq $k }
        if ($page) { $page.Enabled = $On }
    }
    if (-not $On) {
        foreach ($g in @($gridAccess, $gridFolder, $gridStats)) {
            if ($g) { $g.Rows.Clear() }
        }
    }
}

function Set-EditorMode {
    <#
        Point the window at the other directory.

        The tab cache is cleared, not merged: every tab's contents came from
        one directory and none of it is true of the other.
    #>
    param([Parameter(Mandatory)][ValidateSet('Cloud', 'OnPrem')][string]$Mode)
    $script:Mode   = $Mode
    $script:Loaded = @{}
    Set-EditorCloudTabsEnabled -On ($Mode -eq 'Cloud')
    if ($Mode -eq 'OnPrem' -and ($script:CloudOnlyTabs -contains "$($tabs.SelectedTab.Name)")) {
        # Land on a tab that works rather than on a greyed-out one.
        $tabs.SelectedTab = ($tabs.TabPages | Where-Object { $_.Name -eq 'Overview' })
    }

    if ($Mode -eq 'OnPrem' -and $script:Mbx -and $script:Mbx.Ok -and -not $script:OnPrem) {
        try { $script:OnPrem = Resolve-OnPremMailboxTarget -Identity $script:Mbx.Smtp } catch { $script:OnPrem = $null }
    }

    Update-EditorIdentity
    if ($script:Mbx -and $script:Mbx.Ok) {
        $script:TabReadFailed = $false
        Update-OverviewTab
        if (-not $script:TabReadFailed) { $script:Loaded['Overview'] = $true }
        Invoke-EditorTabLoad -Key "$($tabs.SelectedTab.Name)"
    }

    if ($Mode -eq 'OnPrem') {
        $ok = ($script:OnPrem -and $script:OnPrem.Ok)
        Set-EditorStatus -Text $(if ($ok) { "Editing on-premises. $($script:OnPremSyncNotice)" }
                                 else { "$(if ($script:OnPrem) { $script:OnPrem.Reason } else { 'The on-premises object could not be read.' })" }) -Bad:(-not $ok)
    } else {
        Set-EditorStatus -Text 'Editing in Exchange Online.'
    }
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

$tabs.Add_SelectedIndexChanged({
    if ($script:Busy) { return }
    $k = "$($tabs.SelectedTab.Name)"
    Invoke-EditorAction { Invoke-EditorTabLoad -Key $k } "Load the $k tab"
})
#endregion

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Mailbox Editor' | Out-Null

$form.Add_Shown({
    # THE SWITCH IS PROVED BEFORE IT IS OFFERED. A drop-down that lets you
    # pick On-premises and then fails on every Apply is worse than one that is
    # greyed out with the reason sitting next to it.
    $ready = Test-OnPremMailboxReady
    $cmbMode.Enabled = $ready.Ok
    $script:OnPremWhy = $(if ($ready.Ok) { '' } else { $ready.Reason })
    if (-not $ready.Ok) {
        $lblMode.ForeColor = [System.Drawing.Color]::DimGray
        Set-EditorStatus -Text "On-premises editing is unavailable: $($ready.Reason)"
    }
    if ("$($txtMbx.Text)".Trim()) { $btnLoad.PerformClick() }
})

[void]$form.ShowDialog()
$form.Dispose()
