#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
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
# TOOLKIT-HELP-RUNS: Reads one mailbox across several tabs. Each section's Apply changes only that section, and the risky ones ask first.
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

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(1018, ($y - 2))
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
$tabAddresses = New-EditorTab -Key 'Addresses' -Text 'Addresses'
$tabAccess    = New-EditorTab -Key 'Access'    -Text 'Access'
$tabFolders   = New-EditorTab -Key 'Folders'   -Text 'Folder Permissions'
$tabMailFlow  = New-EditorTab -Key 'MailFlow'  -Text 'Mail Flow'
$tabSettings  = New-EditorTab -Key 'Settings'  -Text 'Settings'
$tabRetention = New-EditorTab -Key 'Retention' -Text 'Retention'
$tabReplies   = New-EditorTab -Key 'Replies'   -Text 'Automatic Replies'
$tabRecovery  = New-EditorTab -Key 'Recovery'  -Text 'Recovery'
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
    $lblIdent.Text = "$($m.Display)   <$($m.Smtp)>   $($m.Type)"
    $lblIdent.ForeColor = [System.Drawing.Color]::FromArgb(0, 110, 60)
    if ($m.IsSynced) {
        $lblSync.Text = 'SYNCED FROM ACTIVE DIRECTORY -- addresses, display name, alias and hide-from-GAL are mastered on-premises and cannot be changed here. Everything else can.'
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
    $t = Get-MailboxWriteTarget -Field $Field -IsSynced $script:Mbx.IsSynced
    if ($t.State -eq 'Directory') {
        $null = Show-ToolDialog($t.Reason, 'Mastered in Active Directory', 'OK', 'Warning')
        Set-EditorStatus -Text $t.Reason -Bad
        return $false
    }
    if ($t.State -eq 'Warn') {
        $a = Show-ToolDialog(($t.Reason + "`r`n`r`nGo ahead?"), 'May not persist', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return $false }
    }
    return $true
}

function Write-EditorAudit {
    # One line per action, one shape. Result is Success unless the caller says
    # otherwise -- a failed write throws before it reaches here.
    param([string]$Action, [string]$Detail, [string]$Result = 'Success')
    try {
        Write-CompletedLog -Action $Action -Target (Get-LogTarget $script:Mbx.Smtp) `
            -Result $Result -Detail $Detail -Environment 'Cloud'
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
[void](New-TabLabel $tabOverview 'Mailbox type:' 560 $ovY)
$ovType = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
[void](New-TabLabel $tabOverview 'Primary address:' 8 $ovY)
$ovSmtp = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Alias:' 560 $ovY)
$ovAlias = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
[void](New-TabLabel $tabOverview 'User ID (UPN):' 8 $ovY)
$ovUpn = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Object GUID:' 560 $ovY)
$ovGuid = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
[void](New-TabLabel $tabOverview 'Hidden from GAL:' 8 $ovY)
$ovHidden = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Source:' 560 $ovY)
$ovSource = New-TabValue $tabOverview 700 $ovY 320
$ovY += 30

[void](New-TabLabel $tabOverview 'DIRECTORY' 8 $ovY -Bold)
$ovDirNote = New-TabLabel $tabOverview '' 120 $ovY 900 -Muted
$ovY += 22
[void](New-TabLabel $tabOverview 'Title:' 8 $ovY)
$ovTitle = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Department:' 560 $ovY)
$ovDept = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
[void](New-TabLabel $tabOverview 'Manager:' 8 $ovY)
$ovMgr = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Office:' 560 $ovY)
$ovOffice = New-TabValue $tabOverview 700 $ovY 320
$ovY += 30

[void](New-TabLabel $tabOverview 'SIZE AND ACTIVITY' 8 $ovY -Bold)
$ovY += 22
[void](New-TabLabel $tabOverview 'Mailbox size:' 8 $ovY)
$ovSize = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Items:' 560 $ovY)
$ovItems = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
[void](New-TabLabel $tabOverview 'Quota:' 8 $ovY)
$ovQuota = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Archive:' 560 $ovY)
$ovArchive = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
[void](New-TabLabel $tabOverview 'Last logon:' 8 $ovY)
$ovLogon = New-TabValue $tabOverview 170 $ovY 360
$ovLogonNote = New-TabLabel $tabOverview '' 540 $ovY 540 -Muted
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

function Update-OverviewTab {
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

    if ($notes.Length -eq 0) { [void]$notes.Append('No problems reading this mailbox.') }
    $ovNotes.Text = $notes.ToString()
}

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
    $m  = $script:Mbx
    $mb = $m.Raw
    foreach ($a in @($mb.EmailAddresses)) {
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
        $fresh = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
                    Identity = $script:Mbx.Smtp; ErrorAction = 'Stop' }
        $plan = [PSCustomObject]@{ Cmdlet = 'Set-Mailbox'; PrimaryParam = '' }
        Set-CloudPrimaryAddress -Plan $plan -Identity $script:Mbx.Smtp -NewPrimary $addr `
            -ExistingAddresses @($fresh.EmailAddresses)
        Write-EditorAudit -Action 'MailboxPrimaryAddress' -Detail "Primary address changed to $addr."
        $script:Mbx = Resolve-MailboxTarget -Smtp $addr
        Update-EditorIdentity
        Update-AddressesTab; $script:Loaded['Overview'] = $false
        Set-EditorStatus -Text "Primary address is now $addr."
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
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
            Identity = $script:Mbx.Smtp; EmailAddresses = @{ Add = "smtp:$new" }; ErrorAction = 'Stop'
        } | Out-Null
        Write-EditorAudit -Action 'MailboxAddressAdd' -Detail "Alias $new added."
        $script:Mbx = Resolve-MailboxTarget -Smtp $script:Mbx.Smtp
        Update-AddressesTab; $script:Loaded['Overview'] = $false
        Set-EditorStatus -Text "Added alias $new."
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
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
            Identity = $script:Mbx.Smtp; EmailAddresses = @{ Remove = $full }; ErrorAction = 'Stop'
        } | Out-Null
        Write-EditorAudit -Action 'MailboxAddressRemove' -Detail "Address $full removed."
        $script:Mbx = Resolve-MailboxTarget -Smtp $script:Mbx.Smtp
        Update-AddressesTab; $script:Loaded['Overview'] = $false
        Set-EditorStatus -Text "Removed $addr."
    } 'Remove address'
})
#endregion

#region Access tab --------------------------------------------------------------
$acY = 8
[void](New-TabLabel $tabAccess 'WHO CAN GET AT THIS MAILBOX' 8 $acY -Bold)
$acY += 22
$acNote = New-TabLabel $tabAccess '' 8 $acY 1050 -Muted
$acY += 22

$gridAccess = New-TabGrid $tabAccess @(@('Kind', 120), @('User', 380), @('Note', 120)) 8 $acY 700 300

$acBtnY = $acY
$btnAccAdd    = New-TabButton $tabAccess 'Grant access...'   720 $acBtnY 170
$acBtnY += 34
$btnAccRemove = New-TabButton $tabAccess 'Remove selected'   720 $acBtnY 170
$acBtnY += 34
$btnAccAutomap = New-TabButton $tabAccess 'Fix automapping...' 720 $acBtnY 170
$acBtnY += 40
$acHint = New-TabLabel $tabAccess '' 720 $acBtnY 340 -Muted
$acHint.Size = New-Object System.Drawing.Size(340, 130)
$acHint.Text = ('Full Access lets someone open the mailbox. Send As sends from it as the mailbox. ' +
                'Send on Behalf sends showing both names.' + "`r`n`r`n" +
                'Fix automapping removes and re-adds Full Access to restamp the link Outlook uses to open it.')
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

function Update-AccessTab {
    $gridAccess.Rows.Clear()
    $m = $script:Mbx
    $r = Get-MailboxAccessEntry -Smtp $m.Smtp
    foreach ($e in @($r.Entries)) {
        $i = $gridAccess.Rows.Add($e.Kind, $e.User, $e.Detail)
        $gridAccess.Rows[$i].Tag = $e
        if ($e.Detail -eq 'DENY') { $gridAccess.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 0, 0) }
    }
    $acNote.Text = "$($gridAccess.Rows.Count) grant(s). Inherited and self entries are not shown. $($r.Note)"

    $mb = $m.Raw
    $chkCopySentAs.Checked   = [bool]$mb.MessageCopyForSentAsEnabled
    $chkCopyOnBehalf.Checked = [bool]$mb.MessageCopyForSendOnBehalfEnabled
}

$btnAccAdd.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $who = Show-InputDialog -Title 'Grant access' -Width 560 `
                -Prompt "Who should get access to $($script:Mbx.Display)? Address or alias:"
        $who = "$who".Trim()
        if (-not $who) { return }

        $kind = Show-InputDialog -Title 'Grant access' -Width 560 -Default 'FullAccess' `
                -Prompt 'Which access? FullAccess, SendAs or SendOnBehalf:'
        $kind = "$kind".Trim()
        if ($kind -notmatch '(?i)^(FullAccess|SendAs|SendOnBehalf)$') {
            Set-EditorStatus -Text 'Access must be FullAccess, SendAs or SendOnBehalf.' -Bad; return
        }
        # Normalise the case so the cmdlet and the audit line agree.
        $kind = @{ 'fullaccess' = 'FullAccess'; 'sendas' = 'SendAs'; 'sendonbehalf' = 'SendOnBehalf' }[$kind.ToLowerInvariant()]

        if ($kind -eq 'SendOnBehalf' -and -not (Test-EditorFieldWritable -Field 'GrantSendOnBehalfTo')) { return }

        $auto = $true
        if ($kind -eq 'FullAccess') {
            # ASKED, NOT ASSUMED. A colleague covering a mailbox wants it to
            # appear in Outlook; a shared mailbox thirty people can open does
            # not want to land in thirty profiles.
            $a = Show-ToolDialog(
                ("Should $who's Outlook open this mailbox automatically (automapping)?" + "`r`n`r`n" +
                 'Yes  -- it appears in their Outlook on its own. Right for covering one person.' + "`r`n" +
                 'No   -- they have access but must add it themselves. Right for a busy shared mailbox.'),
                'Automapping', 'YesNo', 'Question')
            $auto = ($a -eq [System.Windows.Forms.DialogResult]::Yes)
        }

        Add-MailboxAccessEntry -Smtp $script:Mbx.Smtp -Kind $kind -Trustee $who -AutoMap $auto
        Write-EditorAudit -Action "MailboxAccessAdd" -Detail "$kind granted to $who$(if ($kind -eq 'FullAccess') { "; automapping $auto" })."
        Update-AccessTab
        Set-EditorStatus -Text "Granted $kind to $who."
    } 'Grant access'
})

$btnAccRemove.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $picked = @()
        foreach ($r in $gridAccess.SelectedRows) { if ($r.Tag) { $picked += $r.Tag } }
        if ($picked.Count -eq 0) { Set-EditorStatus -Text 'Select one or more grants to remove.' -Bad; return }
        $names = @($picked | ForEach-Object { "  $($_.Kind)  --  $($_.User)" }) -join "`r`n"
        $a = Show-ToolDialog(("Remove $($picked.Count) grant(s) from $($script:Mbx.Display)?" + "`r`n`r`n" + $names),
                             'Remove access', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $ok = 0; $bad = 0; $errs = @()
        foreach ($e in $picked) {
            try { Remove-MailboxAccessEntry -Smtp $script:Mbx.Smtp -Kind $e.Kind -Trustee $e.User; $ok++ }
            catch { $bad++; $errs += "$($e.Kind) $($e.User): $($_.Exception.Message)" }
        }
        Write-EditorAudit -Action 'MailboxAccessRemove' `
            -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
            -Detail "$ok removed, $bad failed: $(($picked | ForEach-Object { "$($_.Kind) $($_.User)" }) -join '; ')"
        Update-AccessTab
        Set-EditorStatus -Text "Removed $ok, failed $bad.$(if ($errs) { '  ' + ($errs -join '  ') })" -Bad:($bad -gt 0)
    } 'Remove access'
})

$btnAccAutomap.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $sel = @($gridAccess.SelectedRows | Where-Object { $_.Tag -and $_.Tag.Kind -eq 'FullAccess' })
        if ($sel.Count -ne 1) { Set-EditorStatus -Text 'Select exactly one Full Access row to repair.' -Bad; return }
        $who = "$($sel[0].Tag.User)"
        $a = Show-ToolDialog(
            ("Repair automapping for $who on $($script:Mbx.Display)?" + "`r`n`r`n" +
             'Yes -- make it open automatically in their Outlook.' + "`r`n" +
             'No  -- stop it opening automatically, but KEEP their access.' + "`r`n`r`n" +
             'Either way the Full Access permission is removed and re-added, so there are a few seconds where they have no access.'),
            'Fix automapping', 'YesNoCancel', 'Question')
        if ($a -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
        $auto = ($a -eq [System.Windows.Forms.DialogResult]::Yes)
        Repair-MailboxAutoMapping -Smtp $script:Mbx.Smtp -Trustee $who -AutoMap $auto
        Write-EditorAudit -Action 'MailboxAutomapFix' -Detail "Full Access for $who restamped with automapping $auto."
        Update-AccessTab
        Set-EditorStatus -Text "Automapping for $who is now $(if ($auto) { 'on' } else { 'off' }). Outlook can take a while to notice."
    } 'Fix automapping'
})

$btnCopyApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
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
    CALENDAR, INBOX, CONTACTS, TASKS AND NOTES IN ONE PLACE.

    Three tools used to cover this badly: CalendarPermissions did the calendar,
    UserDefinedDelegates showed the primary folders read-mostly, and NOTHING
    did the inbox -- which is the folder people ask about second. The cmdlets
    are identical for all five; only the folder in the Identity changes. So one
    grid, one dropdown, and the gap closes for free.

    'ALL FOLDERS' IS THE DELEGATE VIEW. Outlook's idea of a delegate is not a
    single object you can read -- it is folder rights across these folders plus
    Send on Behalf. Listing every folder at once and sorting by user IS that
    question answered, without inventing a 'delegate' abstraction that Exchange
    does not have and that would go stale the moment somebody set a permission
    outside this tool.
#>
$fpY = 8
[void](New-TabLabel $tabFolders 'FOLDER PERMISSIONS' 8 $fpY -Bold)
$fpY += 22
$fpNote = New-TabLabel $tabFolders '' 8 $fpY 1050 -Muted
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

$gridFolder = New-TabGrid $tabFolders @(@('Folder', 110), @('User', 330), @('Rights', 250)) 8 $fpY 700 340

$fpBtnY = $fpY
$btnFolderAdd    = New-TabButton $tabFolders 'Grant / change...' 720 $fpBtnY 170
$fpBtnY += 34
$btnFolderRemove = New-TabButton $tabFolders 'Remove selected'   720 $fpBtnY 170
$fpBtnY += 44
$fpHint = New-TabLabel $tabFolders '' 720 $fpBtnY 340 -Muted
$fpHint.Size = New-Object System.Drawing.Size(340, 220)
$fpHint.Text = ('Default is what everyone in the organisation gets. Anonymous is what people' +
                ' outside get. Those two are the usual answer to "why can everyone see my' +
                ' calendar".' + "`r`n`r`n" +
                'Reviewer reads. Editor reads, creates and changes. AvailabilityOnly shows' +
                ' free/busy and nothing else -- calendar only.' + "`r`n`r`n" +
                'Grant / change works either way: it adds the right if the person has none' +
                ' and updates it if they already do.')
$fpY += 352

function Get-EditorSelectedFolderType {
    # Which folder a button should act on. On 'All folders' the grid row
    # decides, because acting on 'all five at once' is a request nobody makes
    # and an easy way to grant far more than intended.
    if ($cmbFolder.SelectedIndex -gt 0) {
        return "$($script:MailboxWellKnownFolder[$cmbFolder.SelectedIndex - 1].Type)"
    }
    if ($gridFolder.SelectedRows.Count -gt 0) {
        return "$($gridFolder.SelectedRows[0].Cells['Folder'].Value)"
    }
    return ''
}

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
            $i = $gridFolder.Rows.Add($f.Type, $e.User, $e.Rights)
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
    Set-EditorStatus -Text $msg
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

$btnFolderAdd.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $ft = Get-EditorSelectedFolderType
        if (-not $ft) { Set-EditorStatus -Text 'Pick a folder, or select a row to say which folder.' -Bad; return }

        # Seeded from the selected row, because 'change Jane from Reviewer to
        # Editor' is the common case and retyping her address invites a typo
        # that silently creates a second entry for the wrong person.
        $seedUser = ''
        if ($gridFolder.SelectedRows.Count -gt 0) { $seedUser = "$($gridFolder.SelectedRows[0].Cells['User'].Value)" }

        $who = Show-InputDialog -Title "Folder permission -- $ft" -Width 560 -Default $seedUser `
                -Prompt "Who? An address or alias, or the words Default or Anonymous:"
        $who = "$who".Trim()
        if (-not $who) { return }

        $role = Show-InputDialog -Title "Folder permission -- $ft" -Width 620 -Default 'Reviewer' `
                -Prompt ("Which role?  " + ($script:MailboxFolderRole -join ', '))
        $role = "$role".Trim()
        if (-not $role) { return }
        if ($script:MailboxFolderRole -notcontains $role) {
            # Case-insensitive match, then use the canonical spelling: Exchange
            # accepts 'reviewer' but the audit line should not record it that way.
            $hit = @($script:MailboxFolderRole | Where-Object { $_ -ieq $role })
            if ($hit.Count -ne 1) {
                Set-EditorStatus -Text "'$role' is not a role. Use one of: $($script:MailboxFolderRole -join ', ')" -Bad
                return
            }
            $role = $hit[0]
        }

        # DEFAULT AND ANONYMOUS CHANGE WHAT EVERYONE SEES, so they are the one
        # case here that asks before acting.
        if ($who -imatch '^(Default|Anonymous)$' -and $role -notmatch '(?i)^(None|AvailabilityOnly)$') {
            $scope = $(if ($who -imatch '^Default$') { 'EVERYONE in the organisation' } else { 'ANYONE outside the organisation' })
            $a = Show-ToolDialog(
                ("Give $scope the $role role on $($script:Mbx.Display)'s $ft folder?" + "`r`n`r`n" +
                 'This is not one person -- it is the baseline for everybody.'),
                'Change the default permission', 'OKCancel', 'Warning')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        }

        $what = Set-MailboxFolderAccess -Smtp $script:Mbx.Smtp -FolderType $ft -User $who -Role $role
        Write-EditorAudit -Action 'MailboxFolderPermission' -Detail "$what $role for $who on the $ft folder."
        Update-FoldersTab
        Set-EditorStatus -Text "$what $role for $who on $ft."
    } 'Grant folder permission'
})

$btnFolderRemove.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $picked = @()
        foreach ($r in $gridFolder.SelectedRows) {
            if (-not $r.Tag) { continue }
            $picked += [PSCustomObject]@{
                Folder = "$($r.Cells['Folder'].Value)"
                User   = "$($r.Cells['User'].Value)"
                Rights = "$($r.Cells['Rights'].Value)"
            }
        }
        if ($picked.Count -eq 0) { Set-EditorStatus -Text 'Select one or more permissions to remove.' -Bad; return }

        # DEFAULT AND ANONYMOUS CANNOT BE REMOVED, only set to None. Exchange
        # refuses the remove with a message about the entry being required,
        # which reads as a fault rather than as 'you meant something else'.
        $builtin = @($picked | Where-Object { $_.User -imatch '^(Default|Anonymous)$' })
        if ($builtin.Count -gt 0) {
            $null = Show-ToolDialog(
                ("Default and Anonymous cannot be removed -- every folder always has them." + "`r`n`r`n" +
                 'To take the access away, use Grant / change and set the role to None.'),
                'Built-in entries', 'OK', 'Information')
            return
        }

        $names = @($picked | ForEach-Object { "  $($_.Folder)  $($_.User)  ($($_.Rights))" }) -join "`r`n"
        $a = Show-ToolDialog(("Remove $($picked.Count) folder permission(s) from $($script:Mbx.Display)?" + "`r`n`r`n" + $names),
                             'Remove folder permission', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $ok = 0; $bad = 0; $errs = @()
        foreach ($p in $picked) {
            try { Remove-MailboxFolderAccess -Smtp $script:Mbx.Smtp -FolderType $p.Folder -User $p.User; $ok++ }
            catch { $bad++; $errs += "$($p.Folder)/$($p.User): $($_.Exception.Message)" }
        }
        Write-EditorAudit -Action 'MailboxFolderPermissionRemove' `
            -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
            -Detail "$ok removed, $bad failed: $(($picked | ForEach-Object { "$($_.Folder)/$($_.User)" }) -join '; ')"
        Update-FoldersTab
        Set-EditorStatus -Text "Removed $ok, failed $bad.$(if ($errs) { '  ' + ($errs -join '  ') })" -Bad:($bad -gt 0)
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
        Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters @{
            Identity = $script:Mbx.Smtp
            HiddenFromAddressListsEnabled = [bool]$chkHidden.Checked
            ErrorAction = 'Stop'
        } | Out-Null
        Write-EditorAudit -Action 'MailboxHideFromGal' -Detail "Hidden from address list = $($chkHidden.Checked)."
        $script:Mbx = Resolve-MailboxTarget -Smtp $script:Mbx.Smtp
        Update-SettingsTab; $script:Loaded['Overview'] = $false
        Set-EditorStatus -Text $(if ($chkHidden.Checked) { 'Hidden from the address list.' } else { 'Visible in the address list.' })
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
    OUT OF OFFICE, as a tab.

    THE INTERNAL AND EXTERNAL MESSAGES ARE SEPARATE AND BOTH ARE SENT. A
    common mistake is filling in only the internal one and assuming customers
    see it; they see nothing. Both boxes are shown side by side for that
    reason, and the external audience setting is on screen rather than buried.

    The stored messages are HTML. They are shown here as text, and what is
    typed is sent as typed -- Exchange wraps plain text acceptably. A message
    written in Outlook with formatting will show its tags here; editing it
    then replaces the formatting with what is in the box, which is said on
    the tab rather than discovered.
#>
$arY = 8
[void](New-TabLabel $tabReplies 'AUTOMATIC REPLIES' 8 $arY -Bold)
$arY += 22
$arNow = New-TabLabel $tabReplies '' 8 $arY 1050 -Muted
$arY += 24

[void](New-TabLabel $tabReplies 'State:' 8 ($arY + 3) 60)
$cmbOof               = New-Object System.Windows.Forms.ComboBox
$cmbOof.DropDownStyle = 'DropDownList'
$cmbOof.Location      = New-Object System.Drawing.Point(70, $arY)
$cmbOof.Size          = New-Object System.Drawing.Size(160, 24)
foreach ($s in @('Disabled', 'Enabled', 'Scheduled')) { [void]$cmbOof.Items.Add($s) }
$cmbOof.SelectedIndex = 0
$tabReplies.Controls.Add($cmbOof)

[void](New-TabLabel $tabReplies 'From:' 244 ($arY + 3) 44)
$dtOofStart          = New-Object System.Windows.Forms.DateTimePicker
$dtOofStart.Location = New-Object System.Drawing.Point(290, $arY)
$dtOofStart.Size     = New-Object System.Drawing.Size(180, 24)
$dtOofStart.Format   = 'Short'
$tabReplies.Controls.Add($dtOofStart)
[void](New-TabLabel $tabReplies 'To:' 480 ($arY + 3) 30)
$dtOofEnd          = New-Object System.Windows.Forms.DateTimePicker
$dtOofEnd.Location = New-Object System.Drawing.Point(512, $arY)
$dtOofEnd.Size     = New-Object System.Drawing.Size(180, 24)
$dtOofEnd.Format   = 'Short'
$tabReplies.Controls.Add($dtOofEnd)
$arY += 30

[void](New-TabLabel $tabReplies 'External replies go to:' 8 ($arY + 3) 150)
$cmbOofAudience               = New-Object System.Windows.Forms.ComboBox
$cmbOofAudience.DropDownStyle = 'DropDownList'
$cmbOofAudience.Location      = New-Object System.Drawing.Point(164, $arY)
$cmbOofAudience.Size          = New-Object System.Drawing.Size(240, 24)
foreach ($s in @('None', 'Known', 'All')) { [void]$cmbOofAudience.Items.Add($s) }
$cmbOofAudience.SelectedIndex = 0
$tabReplies.Controls.Add($cmbOofAudience)
$arAudNote = New-TabLabel $tabReplies '' 414 ($arY + 3) 640 -Muted
$arAudNote.Text = 'None = nobody outside gets a reply. Known = only contacts. All = anyone who writes in.'
$arY += 30

[void](New-TabLabel $tabReplies 'Inside the organisation:' 8 $arY 300 -Bold)
[void](New-TabLabel $tabReplies 'Outside the organisation:' 540 $arY 300 -Bold)
$arY += 20

$txtOofInt            = New-Object System.Windows.Forms.TextBox
$txtOofInt.Multiline  = $true
$txtOofInt.ScrollBars = 'Vertical'
$txtOofInt.Location   = New-Object System.Drawing.Point(8, $arY)
$txtOofInt.Size       = New-Object System.Drawing.Size(510, 300)
$tabReplies.Controls.Add($txtOofInt)

$txtOofExt            = New-Object System.Windows.Forms.TextBox
$txtOofExt.Multiline  = $true
$txtOofExt.ScrollBars = 'Vertical'
$txtOofExt.Location   = New-Object System.Drawing.Point(540, $arY)
$txtOofExt.Size       = New-Object System.Drawing.Size(510, 300)
$tabReplies.Controls.Add($txtOofExt)
$arY += 308

$btnOofApply = New-TabButton $tabReplies 'Apply automatic replies' 8 $arY 200
$btnOofOff   = New-TabButton $tabReplies 'Turn off'                218 $arY 120
$arY += 34
$arNote = New-TabLabel $tabReplies '' 8 $arY 1050 -Muted
$arNote.Size = New-Object System.Drawing.Size(1050, 34)
$arNote.Text = 'Messages are stored as HTML. Anything typed here is sent as typed and replaces whatever formatting was there before.'

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
    if ($cmbOof.Items.Contains($state)) { $cmbOof.SelectedItem = $state }
    $arNow.Text = "Currently: $state$(if ($state -eq 'Scheduled') { " -- $($c.StartTime) to $($c.EndTime)" })"
    $arNow.ForeColor = $(if ($state -eq 'Disabled') { [System.Drawing.Color]::DimGray } else { [System.Drawing.Color]::FromArgb(0, 110, 60) })

    try { if ($c.StartTime) { $dtOofStart.Value = [datetime]$c.StartTime } } catch { }
    try { if ($c.EndTime)   { $dtOofEnd.Value   = [datetime]$c.EndTime } } catch { }
    $aud = "$($c.ExternalAudience)"
    if ($cmbOofAudience.Items.Contains($aud)) { $cmbOofAudience.SelectedItem = $aud }
    $txtOofInt.Text = "$($c.InternalMessage)"
    $txtOofExt.Text = "$($c.ExternalMessage)"
}

$btnOofApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorFieldWritable -Field 'AutomaticReplies')) { return }
        $state = "$($cmbOof.SelectedItem)"
        $p = @{
            Identity         = $script:Mbx.Smtp
            AutoReplyState   = $state
            ExternalAudience = "$($cmbOofAudience.SelectedItem)"
            InternalMessage  = "$($txtOofInt.Text)"
            ExternalMessage  = "$($txtOofExt.Text)"
            ErrorAction      = 'Stop'
        }
        if ($state -eq 'Scheduled') {
            if ($dtOofEnd.Value -le $dtOofStart.Value) {
                Set-EditorStatus -Text 'The end of the schedule must be after the start.' -Bad; return
            }
            $p['StartTime'] = $dtOofStart.Value
            $p['EndTime']   = $dtOofEnd.Value
        }
        # THE ONE THAT CATCHES PEOPLE OUT: replies are on, but the external
        # audience is None, so customers get silence. Said before it happens
        # rather than discovered from a complaint.
        if ($state -ne 'Disabled' -and "$($cmbOofAudience.SelectedItem)" -eq 'None' -and "$($txtOofExt.Text)".Trim()) {
            $a = Show-ToolDialog(
                ("There is an external message, but external replies are set to None -- nobody outside the organisation will receive it." + "`r`n`r`n" +
                 'Apply anyway?'),
                'External replies are off', 'OKCancel', 'Warning')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        }
        Invoke-EXOCommand -Name 'Set-MailboxAutoReplyConfiguration' -Parameters $p | Out-Null
        Write-EditorAudit -Action 'MailboxAutoReply' -Detail "State $state; external audience $($cmbOofAudience.SelectedItem)$(if ($state -eq 'Scheduled') { "; $($dtOofStart.Value.ToShortDateString()) to $($dtOofEnd.Value.ToShortDateString())" })."
        Update-RepliesTab
        Set-EditorStatus -Text "Automatic replies: $state."
    } 'Apply automatic replies'
})

$btnOofOff.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        if (-not (Test-EditorFieldWritable -Field 'AutomaticReplies')) { return }
        Invoke-EXOCommand -Name 'Set-MailboxAutoReplyConfiguration' -Parameters @{
            Identity = $script:Mbx.Smtp; AutoReplyState = 'Disabled'; ErrorAction = 'Stop'
        } | Out-Null
        Write-EditorAudit -Action 'MailboxAutoReply' -Detail 'Automatic replies turned off.'
        Update-RepliesTab
        Set-EditorStatus -Text 'Automatic replies turned off. The messages are kept.'
    } 'Turn off automatic replies'
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

#region Load and tab switching --------------------------------------------------
function Invoke-EditorLoad {
    $t = "$($txtMbx.Text)".Trim()
    if (-not $t) { Set-EditorStatus -Text 'Enter a mailbox address.' -Bad; return }
    Set-EditorStatus -Text "Resolving $t..."
    $m = Resolve-MailboxTarget -Smtp $t
    if (-not $m.Ok) {
        $script:Mbx = $null
        $lblIdent.Text = ''; $lblSync.Text = ''
        Set-EditorStatus -Text "'$t' $($m.Reason)" -Bad
        return
    }
    $script:Mbx = $m
    $script:Loaded = @{}

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
    }
    if (-not $script:TabReadFailed) { $script:Loaded[$Key] = $true }
}

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
    if ("$($txtMbx.Text)".Trim()) { $btnLoad.PerformClick() }
})

[void]$form.ShowDialog()
$form.Dispose()
