#Requires -Version 5.1
# TOOLKIT-VARIANT: Dual
#
# CLOUD BY DEFAULT, ON-PREMISES ON A SWITCH -- the same Editing drop-down the
# Mailbox Editor has, in the same place, with the same wording, built by the
# same function in Common.ps1. A room's addresses, alias, display name and
# hide-from-address-book are mastered on-premises when it is synced, and the
# Place address block with them.
<#
    RoomEditor.ps1
    One room, every setting. The window you open when somebody says "the room
    is not accepting bookings" and you do not yet know why.

    Author   : Glenn Armbruster
    Created  : 2026-09-18

    WHY THIS EXISTS. Seven tools answered parts of one question -- calendar
    processing, booking policy, room lists, Place metadata, regional and
    working hours, calendar permissions, and the duplicator's capture list.
    Finding out why a room misbehaves meant opening four of them and holding
    the answers in your head, because no window could see what another had
    found. A room is one object and it now has one window.

    THE SAME SHAPE AS MailboxEditor.ps1, deliberately: tabs across one target,
    per-section Apply, no global Save, every write audited against the room.
#>
# TOOLKIT-HELP-RUNS: Reads one room or equipment mailbox across several tabs, in Exchange Online or -- on the Editing switch -- against the on-premises object. Each section's Apply changes only that section, and the risky ones ask first.
# TOOLKIT-HELP-BEGIN
# Everything about one room, in tabs. Type an address, press Load.
#
#   Overview      What it is, capacity, whether it is hidden, and a plain-
#                 language verdict on whether it will actually take bookings.
#   Booking       How the room answers requests -- AutomateProcessing and the
#                 whole automatic-acceptance policy.
#   Who Can Book  The three in-policy lists and the delegates, as resolved
#                 people rather than raw legacy DNs.
#   Replies       What the room strips from a request and what it says back.
#   Place         Building, floor, capacity, devices, accessibility -- the
#                 record Room Finder actually reads.
#   Regional      Timezone, language, and the working hours that
#                 ScheduleOnlyDuringWorkHours depends on entirely.
#   Room Lists    Which lists this room appears in for Room Finder.
#   Calendar      Who can see and edit the room's calendar.
#
# THE ORDER OF THE TABS IS THE ORDER OF THE QUESTIONS. 'Why will it not
# accept' is Booking; 'why will it not accept for HER' is Who Can Book; 'why
# can nobody find it' is Room Lists and Place.
#
# EACH BLOCK HAS ITS OWN APPLY. There is no global Save. Whatever you press
# does that one thing and logs it; nothing else on the tab is touched.
#
# THE EDITING SWITCH, TOP RIGHT. The same switch the Mailbox Editor has, and
# it means the same thing. Cloud is the default. Set it to On-premises when
# the room is synced from Active Directory and you need a field the cloud will
# not keep -- addresses, alias, hide-from-address-list, or the directory
# fields. The banner turns amber and says so, and the tabs that only exist in
# Exchange Online are greyed out.
#
# If the switch is greyed out, this window has no on-premises Exchange
# connection (or the tenant has none); the reason is on the sync banner.
#
# AN ON-PREMISES CHANGE IS NOT INSTANT -- allow a directory sync cycle, about
# half an hour, before deciding it did not work.
#
# A synced room has some Place fields mastered on-premises. Those are shown
# and refused here rather than accepted and silently reverted.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not (Test-ConsoleLaunch)) { return }

#region State -------------------------------------------------------------------
# ONE PLACE FOR 'WHICH ROOM AND WHAT DO WE KNOW ABOUT IT'. Every tab reads
# $script:Room; nothing else holds a copy, so a reload cannot leave one tab
# showing the previous room.
$script:Room   = $null    # Resolve-MailboxTarget result
$script:Cal    = $null    # Get-CalendarProcessing result
$script:Loaded = @{}
$script:Busy   = $false
$script:Closing = $false   # set by FormClosing; see the tab handler at the foot of the file
$script:TabReadFailed = $false

# THE AUTO-ACCEPT-ONLY SETTINGS, carried over from CalendarProcessing.ps1.
#
# An operator who sets AllowConflicts while AutomateProcessing is None gets a
# successful write, no error, and no change in behaviour -- then reports that
# the tool does not work. The list is deliberately conservative: a warning that
# is sometimes wrong is worse than no warning.
$script:RoomAutoAcceptOnly = @(
    'AllowConflicts', 'AllowRecurringMeetings', 'BookingWindowInDays',
    'MaximumDurationInMinutes', 'MinimumDurationInMinutes',
    'ConflictPercentageAllowed', 'MaximumConflictInstances',
    'EnforceCapacity', 'EnforceSchedulingHorizon',
    'ScheduleOnlyDuringWorkHours', 'ProcessExternalMeetingMessages'
)

# THE PLACE FIELDS A SYNCED ROOM MASTERS ON-PREMISES.
#
# Quoted from the Set-Place reference. Refused BEFORE the call, because the
# cmdlet does not always error on these -- it can accept the call and simply
# not persist the value, which is the worst possible outcome: a success
# message and no change.
$script:PlaceHybridBlocked = @(
    'City', 'CountryOrRegion', 'GeoCoordinates', 'Phone',
    'PostalCode', 'State', 'Street'
)

# WHICH DIRECTORY THIS WINDOW IS WRITING TO. Same three variables, same names
# and same meanings as the Mailbox Editor, so the two files read alike.
$script:Mode      = 'Cloud'
$script:OnPrem    = $null   # Resolve-OnPremMailboxTarget result
$script:OnPremWhy = ''      # why the switch is unavailable, if it is
#endregion

#region Form shell --------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text          = 'Room Editor'
$form.ClientSize    = New-Object System.Drawing.Size(1120, 760)
$form.StartPosition = 'Manual'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$y = 12
$lblWho          = New-Object System.Windows.Forms.Label
$lblWho.Text     = 'Room:'
$lblWho.Location = New-Object System.Drawing.Point(12, ($y + 3))
$lblWho.Size     = New-Object System.Drawing.Size(48, 18)
$form.Controls.Add($lblWho)

$txtRoom          = New-Object System.Windows.Forms.TextBox
$txtRoom.Location = New-Object System.Drawing.Point(62, $y)
$txtRoom.Size     = New-Object System.Drawing.Size(340, 24)
$txtRoom.Font     = New-Object System.Drawing.Font('Consolas', 9)
$txtRoom.Text     = "$(Get-TargetSmtp -Text $SmtpAddress)"
$form.Controls.Add($txtRoom)

$btnLoad          = New-Object System.Windows.Forms.Button
$btnLoad.Text     = 'Load'
$btnLoad.Location = New-Object System.Drawing.Point(412, ($y - 2))
$btnLoad.Size     = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnLoad)

# CLOSE SITS WELL CLEAR OF THE CORNER. Add-ToolHelpButton drops the '?' in the
# first free spot it finds, and a Close button in the corner leaves it nowhere
# to go but on top -- which is exactly what happened to the Mailbox Editor.
# THE EDITING SWITCH, from Common.ps1 so all three editors get the identical
# control rather than three hand-typed copies of the thing that decides which
# directory a write lands in.
$cmbMode = New-EditorModeSwitch -Form $form -X 718 -Y $y -LabelX 660

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(952, ($y - 2))
$btnClose.Size     = New-Object System.Drawing.Size(90, 28)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)
$y += 34

$lblIdent           = New-Object System.Windows.Forms.Label
$lblIdent.Tag       = 'NoTheme'
$lblIdent.Location  = New-Object System.Drawing.Point(12, $y)
$lblIdent.Size      = New-Object System.Drawing.Size(1096, 18)
$lblIdent.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblIdent.ForeColor = (Get-ToolStatusColour -Kind Muted)
$lblIdent.Text      = 'No room loaded.'
$form.Controls.Add($lblIdent)
$y += 20

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
$lblStatus.Text      = 'Enter a room address and press Load.'
$form.Controls.Add($lblStatus)

$form.ClientSize = New-Object System.Drawing.Size(1120, ($lblStatus.Bottom + 12))


# THE ORDER IS THE ORDER OF THE QUESTIONS, not alphabetical and not the order
# the old tools happened to be written in.
# ---- register with the shared editor runtime ----------------------------------
Initialize-TabbedEditor -Name 'Room' -Noun 'room' -ObjectName 'Room' `
    -Form $form -Tabs $tabs -StatusLabel $lblStatus -IdentityLabel $lblIdent -SyncLabel $lblSync -ModeCombo $cmbMode `
    -CloudOnlyTabs @('Booking', 'Who', 'Replies', 'Place', 'Regional', 'Lists', 'CalPerm') `
    -CloudGridNames @('gridBooking', 'gridPolicy', 'gridDeleg', 'gridPlace', 'gridLists', 'gridCalPerm') `
    -TabLoaders @{
        Overview = { Update-OverviewTab }
        Booking  = { Update-BookingTab }
        Who      = { Update-WhoTab }
        Replies  = { Update-RepliesTab }
        Place    = { Update-PlaceTab }
        Regional = { Update-RegionalTab }
        Lists    = { Update-ListsTab }
        CalPerm  = { Update-CalPermTab }
    } `
    -Identity {
        param($m)
        $cap = ''
        try { if ($null -ne $m.Raw.ResourceCapacity) { $cap = "   capacity $($m.Raw.ResourceCapacity)" } } catch { }
        if ("$($m.Type)" -notmatch '(?i)Room|Equipment') {
            return @{ Text = "$($m.Display)   <$($m.Smtp)>   $($m.Type)$cap"; Colour = 'Good'
                      Sync = "THIS IS NOT A ROOM OR EQUIPMENT MAILBOX ($($m.Type)). Calendar processing answers for any mailbox but only governs a resource -- nothing on the Booking tab will change how this one behaves."
                      SyncColour = 'Bad' }
        }
        $sync = $(if ($m.IsSynced) {
                if ($script:OnPremWhy) { "SYNCED FROM ACTIVE DIRECTORY -- addresses, alias, hide-from-address-book and the Place address block are mastered on-premises. On-premises editing is unavailable here: $($script:OnPremWhy)" }
                else { 'SYNCED FROM ACTIVE DIRECTORY -- addresses, alias, hide-from-address-book and the Place address block are mastered on-premises. Set Editing (top right) to On-premises to change them.' }
            } else { 'Cloud-only room -- every field on every tab is editable here.' })
        @{ Text = "$($m.Display)   <$($m.Smtp)>   $($m.Type)$cap"; Colour = 'Good'
           Sync = $sync; SyncColour = $(if ($m.IsSynced) { 'Warn' } else { 'Muted' }) }
    }

$tabOverview = New-EditorTab -Key 'Overview' -Text 'Overview'
$tabBooking  = New-EditorTab -Key 'Booking'  -Text 'Booking'
$tabWho      = New-EditorTab -Key 'Who'      -Text 'Who Can Book'
$tabReplies  = New-EditorTab -Key 'Replies'  -Text 'Replies'
$tabPlace    = New-EditorTab -Key 'Place'    -Text 'Place'
$tabRegional = New-EditorTab -Key 'Regional' -Text 'Regional'
$tabLists    = New-EditorTab -Key 'Lists'    -Text 'Room Lists'
$tabCalPerm  = New-EditorTab -Key 'CalPerm'  -Text 'Calendar'
#endregion

#region Shell helpers -----------------------------------------------------------





function Get-RoomBool {
    # Booleans are a dropdown everywhere in this window, never a text box:
    # 'false' cast to [bool] is $true, so a typo would enable the thing it was
    # meant to disable, silently.
    param($Parent, [int]$X, [int]$Y, [int]$W = 110)
    $c = New-Object System.Windows.Forms.ComboBox
    $c.DropDownStyle = 'DropDownList'
    $c.Location = New-Object System.Drawing.Point($X, $Y)
    $c.Size     = New-Object System.Drawing.Size($W, 24)
    [void]$c.Items.Add('True'); [void]$c.Items.Add('False')
    $Parent.Controls.Add($c)
    return $c
}

function Set-RoomBool {
    param($Combo, $Value)
    $Combo.SelectedItem = $(if ([bool]$Value) { 'True' } else { 'False' })
}

function Get-RoomBoolValue {
    param($Combo)
    return ("$($Combo.SelectedItem)" -eq 'True')
}
#endregion

#region Overview tab ------------------------------------------------------------
$ovY = 8
[void](New-TabLabel $tabOverview 'IDENTITY' 8 $ovY -Bold)
$ovY += 22
[void](New-TabLabel $tabOverview 'Display name:' 8 $ovY)
$ovName = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Type:' 560 $ovY 135)
$ovType = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
[void](New-TabLabel $tabOverview 'Address:' 8 $ovY)
$ovSmtp = New-TabValue $tabOverview 170 $ovY 360
[void](New-TabLabel $tabOverview 'Alias:' 560 $ovY 135)
$ovAlias = New-TabValue $tabOverview 700 $ovY 320
$ovY += 24
[void](New-TabLabel $tabOverview 'Source:' 8 $ovY)
$ovSource = New-TabValue $tabOverview 170 $ovY 360
$ovY += 30

[void](New-TabLabel $tabOverview 'CAPACITY AND VISIBILITY' 8 $ovY -Bold)
$ovY += 22
[void](New-TabLabel $tabOverview 'Seats:' 8 ($ovY + 3) 110)
$txtCapacity          = New-Object System.Windows.Forms.TextBox
$txtCapacity.Location = New-Object System.Drawing.Point(170, $ovY)
$txtCapacity.Size     = New-Object System.Drawing.Size(80, 24)
$tabOverview.Controls.Add($txtCapacity)
[void](New-TabLabel $tabOverview 'Hide from the address book:' 300 ($ovY + 3) 180)
$cmbHide = Get-RoomBool $tabOverview 486 $ovY
$btnOvApply = New-TabButton $tabOverview 'Apply' 620 ($ovY - 2) 100
$ovY += 28
$ovCapNote = New-TabLabel $tabOverview '' 8 $ovY 1040 -Muted
$ovCapNote.Size = New-Object System.Drawing.Size(1040, 32)
$ovCapNote.Text = ('Seats is ResourceCapacity on the mailbox. It is what EnforceCapacity measures against, and it is a different field from Capacity on the Place tab -- ' + "`r`n" +
                   'Room Finder reads the Place one. They should agree; nothing makes them.')
$ovY += 38

# THE SAME BOX EVERY EDITOR HAS. The question it answers for a room is
# still 'will it take a booking' -- that is what 'anything wrong' means here.
$ovVerdict = Add-EditorVerdictBox -Tab $tabOverview -Y $ovY -Height 230
$ovY += 262
$btnOvRecheck = New-TabButton $tabOverview 'Re-check' 8 $ovY 130

function Update-OverviewOnPrem {
    <#
        The Overview tab, read from the on-premises object.

        SEATS AND THE VERDICT ARE BLANKED rather than left showing the cloud
        numbers: ResourceCapacity and every calendar-processing setting are
        Exchange Online properties, and a capacity beside an on-premises banner
        reads as the on-premises one.
    #>
    $o = $script:OnPrem
    if (-not $o -or -not $o.Ok) {
        # BLANKED, NOT LEFT AS IT WAS. An early return here used to leave the
        # CLOUD name, address and alias on screen under a red 'not found
        # on-premises' banner -- values that say the opposite of the banner
        # above them.
        foreach ($t in @($ovName, $ovType, $ovSmtp, $ovAlias, $ovSource)) { $t.Text = '' }
        $txtCapacity.Text = ''; $txtCapacity.Enabled = $false
        $cmbHide.Enabled = $false; $btnOvApply.Enabled = $false
        $ovVerdict.Text = $(if ($o) { $o.Reason } else { 'No on-premises object loaded.' })
        return
    }
    $cmbHide.Enabled = $true; $btnOvApply.Enabled = $true
    $ovName.Text   = "$($o.Display)"
    $ovType.Text   = "$($o.Kind)$(if ($o.Raw.RecipientTypeDetails) { " ($($o.Raw.RecipientTypeDetails))" })"
    $ovSmtp.Text   = "$($o.Smtp)"
    $ovAlias.Text  = "$($o.Alias)"
    $ovSource.Text = 'Active Directory (on-premises)'

    $txtCapacity.Text = '(Exchange Online only)'
    $txtCapacity.Enabled = $false
    Set-RoomBool $cmbHide $o.Hidden
    $ovCapNote.Text = 'Seats is a property of the cloud mailbox and cannot be changed here. Hide-from-address-book is mastered on-premises and CAN.'

    $ovVerdict.Text = ("$($o.Reason)" + "`r`n" + $script:OnPremSyncNotice + "`r`n`r`n" +
                       'Booking, the policy lists, replies, Place, regional settings, room lists and calendar permissions are all Exchange Online properties -- those tabs are greyed while Editing is set to On-premises.')
}

function Update-OverviewTab {
    <#
        THE VERDICT IS THE REASON THIS TAB EXISTS.

        'The room is not accepting bookings' has about six causes and they
        live on five different tabs. Reading them one tab at a time is what
        the old seven tools made you do. This states the finding and names the
        tab that holds the detail -- it does not repeat the values, because
        two places showing the same fact is how they come to disagree.
    #>
    # ON-PREMISES MODE READS THE ON-PREMISES OBJECT. Showing the cloud values
    # under an amber banner would say the opposite of the truth about what an
    # Apply is going to change.
    if ($script:Mode -eq 'OnPrem') { Update-OverviewOnPrem; return }

    $m  = $script:Room
    $mb = $m.Raw
    $ovName.Text   = "$($mb.DisplayName)"
    $ovType.Text   = "$($mb.RecipientTypeDetails)"
    $ovSmtp.Text   = "$($mb.PrimarySmtpAddress)"
    $ovAlias.Text  = "$($mb.Alias)"
    $ovSource.Text = $(if ($m.IsSynced) { 'Synced from Active Directory' } else { 'Cloud only' })
    $txtCapacity.Text = "$($mb.ResourceCapacity)"
    $txtCapacity.Enabled = $true
    $cmbHide.Enabled = $true; $btnOvApply.Enabled = $true
    $ovCapNote.Text = ('Seats is ResourceCapacity on the mailbox. It is what EnforceCapacity measures against, and it is a different field from Capacity on the Place tab -- ' + "`r`n" +
                       'Room Finder reads the Place one. They should agree; nothing makes them.')
    Set-RoomBool $cmbHide $mb.HiddenFromAddressListsEnabled

    $f = New-Object System.Collections.ArrayList
    $c = $script:Cal

    if ("$($mb.RecipientTypeDetails)" -notmatch '(?i)Room|Equipment') {
        [void]$f.Add("This is a $($mb.RecipientTypeDetails), not a resource mailbox. Calendar processing settings are stored but never act on it.")
    }
    if ($c) {
        $auto = "$($c.AutomateProcessing)"
        if ($auto -ne 'AutoAccept') {
            [void]$f.Add("AutomateProcessing is $auto, not AutoAccept. $(if ($auto -eq 'None') { 'Requests are left for a delegate to answer by hand -- nothing is automatic.' } else { 'Only the free/busy calendar is updated; requests are not answered.' }) See Booking.")
        }
        if ([bool]$c.AllBookInPolicy -and @($c.BookInPolicy).Count -gt 0) {
            # The bug Set-BookingPolicy.ps1 was written to fix, now stated
            # rather than left to be discovered: the room carries a list of
            # names that nothing reads.
            [void]$f.Add("AllBookInPolicy is True AND there are $(@($c.BookInPolicy).Count) name(s) in BookInPolicy. The list is not being enforced -- everyone can book, while the room LOOKS restricted. See Who Can Book.")
        }
        if (-not [bool]$c.AllBookInPolicy -and @($c.BookInPolicy).Count -eq 0 -and $auto -eq 'AutoAccept') {
            [void]$f.Add('Booking is restricted to a list and the list is EMPTY, so nobody can book automatically. See Who Can Book.')
        }
        if ([bool]$c.ScheduleOnlyDuringWorkHours) {
            [void]$f.Add('ScheduleOnlyDuringWorkHours is on, so requests outside the working hours are declined. Those hours and their timezone are on Regional -- and if the timezone is wrong, the hours are wrong.')
        }
        if ([bool]$c.EnforceCapacity) {
            [void]$f.Add("EnforceCapacity is on, so a request for more attendees than the $($mb.ResourceCapacity) seats set here is declined.")
        }
        if (@($c.ResourceDelegates).Count -eq 0 -and $auto -eq 'None') {
            [void]$f.Add('AutomateProcessing is None and there are no delegates, so nothing and nobody answers a request for this room.')
        }
    } else {
        [void]$f.Add('Calendar processing could not be read -- see the Booking tab.')
    }
    if ($mb.HiddenFromAddressListsEnabled) {
        [void]$f.Add('Hidden from the address book. It can still be booked by typing the address, but it will not be found by browsing.')
    }

    Set-EditorVerdict -Findings @($f) -Box $ovVerdict `
        -Clean 'Nothing obviously wrong. The room auto-accepts, the policy lists agree with each other, and it is visible in the address book.' `
        -Notes @($(if ($f.Count -eq 0) { 'If it is still not being found, check Room Lists -- a room missing from every list does not appear in Outlook''s Room Finder even though it books perfectly well by address.' }))
}

$btnOvRecheck.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $script:Cal = Get-RoomCalendarProcessing -Smtp $script:Room.Smtp
        Update-OverviewTab
        Set-EditorStatus -Text 'Re-checked.'
    } 'Re-check the room'
})

$btnOvApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $onPremMode = ($script:Mode -eq 'OnPrem')
        $p = @{ Identity = $script:Room.Smtp; ErrorAction = 'Stop' }
        $changed = @()

        # SEATS IS NOT EVEN LOOKED AT IN ON-PREMISES MODE. The box shows
        # '(Exchange Online only)' there, and comparing that string to the
        # cloud value made it look like a change, failed the digits test, and
        # returned before the mode routing below was ever reached -- so the
        # hide change, the one thing this Apply exists to do on-premises,
        # could never be written.
        if (-not $onPremMode) {
            $capText = "$($txtCapacity.Text)".Trim()
            if ($capText -ne "$($script:Room.Raw.ResourceCapacity)") {
                if ($capText -and $capText -notmatch '^\d+$') {
                    Set-EditorStatus -Text 'Seats must be a whole number, or empty to clear it.' -Bad; return
                }
                $p['ResourceCapacity'] = $(if ($capText) { [int]$capText } else { $null })
                $changed += "ResourceCapacity=$(if ($capText) { $capText } else { '(cleared)' })"
            }
        }

        # THE BASELINE IS WHICHEVER OBJECT THE BOX WAS FILLED FROM. Comparing
        # an on-premises tick to the CLOUD value discards the change in
        # exactly the case this mode exists for -- the two directories
        # disagreeing.
        $wantHide = Get-RoomBoolValue $cmbHide
        $nowHide  = $(if ($onPremMode) { [bool]($script:OnPrem -and $script:OnPrem.Hidden) }
                      else { [bool]$script:Room.Raw.HiddenFromAddressListsEnabled })
        if ($wantHide -ne $nowHide) {
            # HIDE-FROM-GAL IS DIRECTORY-MASTERED ON A SYNCED ROOM. The same
            # trap the Mailbox Editor documents: Exchange Online accepts the
            # write and directory sync puts it back, with no error anywhere.
            if ($script:Room.IsSynced -and -not $onPremMode) {
                $extra = $(if ($script:OnPremWhy) { " On-premises editing is unavailable here: $($script:OnPremWhy)" }
                           else { ' Set Editing (top right) to On-premises to change it.' })
                $null = Show-ToolDialog(
                    ('Hide-from-address-book is mastered in Active Directory for this room. Exchange Online would accept the change and directory sync would put it back within about half an hour, with no error.' + $extra),
                    'Mastered in Active Directory', 'OK', 'Warning')
                Set-EditorStatus -Text 'Hide-from-address-book is mastered on-premises for this room.' -Bad
                return
            }
            $p['HiddenFromAddressListsEnabled'] = $wantHide
            $changed += "HiddenFromAddressLists=$wantHide"
        }
        if ($changed.Count -eq 0) { Set-EditorStatus -Text 'Nothing changed.'; return }

        # ROUTED BY MODE. Seats is a cloud-only property, so a change to it in
        # on-premises mode is refused rather than sent to a cmdlet that has no
        # such parameter.
        if ($onPremMode) {
            if (-not $script:OnPrem -or -not $script:OnPrem.Ok -or -not $script:OnPrem.SetCmdlet) {
                Set-EditorStatus -Text 'No on-premises object is loaded for this room.' -Bad
                return
            }
            $p['Identity'] = $script:OnPrem.Smtp
            Invoke-OnPremCommand -Name $script:OnPrem.SetCmdlet -Parameters $p | Out-Null
            Write-EditorAudit -Action 'RoomCapacityVisibility' -Detail ($changed -join '; ')
            $script:OnPrem = Resolve-OnPremMailboxTarget -Identity $script:OnPrem.Smtp
        } else {
            Invoke-EXOCommand -Name 'Set-Mailbox' -Parameters $p | Out-Null
            Write-EditorAudit -Action 'RoomCapacityVisibility' -Detail ($changed -join '; ')
            $script:Room = Resolve-MailboxTarget -Smtp $script:Room.Smtp
        }
        Update-EditorIdentity
        Update-OverviewTab
        Set-EditorStatus -Text ("Applied: $($changed -join '; ')" + $(if ($script:Mode -eq 'OnPrem') { " $($script:OnPremSyncNotice)" }))
    } 'Apply capacity and visibility'
})
#endregion

#region Booking tab -------------------------------------------------------------
<#
    HOW THE ROOM ANSWERS A REQUEST.

    A GRID OVER Set-CalendarProcessing, NOT A HAND-PICKED FORM. The writable
    parameter set is discovered at run time from the cmdlet itself, so a
    setting Microsoft adds next year appears here without this file changing.
    Hand-picking would have meant a form that is quietly a year out of date.

    THE NOTE COLUMN IS THE POINT. Most of these settings do nothing unless
    AutomateProcessing is AutoAccept, and Exchange accepts the write either
    way -- no error, no change in behaviour, and a ticket saying the tool does
    not work. The note says which ones are inert right now.
#>
$bkY = 8
[void](New-TabLabel $tabBooking 'AUTOMATIC PROCESSING' 8 $bkY -Bold)
$bkY += 24
[void](New-TabLabel $tabBooking 'AutomateProcessing:' 8 ($bkY + 3) 150)
$cmbAuto               = New-Object System.Windows.Forms.ComboBox
$cmbAuto.DropDownStyle = 'DropDownList'
$cmbAuto.Location      = New-Object System.Drawing.Point(170, $bkY)
$cmbAuto.Size          = New-Object System.Drawing.Size(200, 24)
foreach ($v in @('None', 'AutoUpdate', 'AutoAccept')) { [void]$cmbAuto.Items.Add($v) }
$tabBooking.Controls.Add($cmbAuto)
$btnAutoApply = New-TabButton $tabBooking 'Apply' 386 ($bkY - 2) 100
$bkY += 28
$bkAutoNote = New-TabLabel $tabBooking '' 8 $bkY 1040 -Muted
$bkAutoNote.Size = New-Object System.Drawing.Size(1040, 46)
$bkAutoNote.Text = ('None       -- nothing is automatic. A delegate answers every request by hand, and if there is no delegate nothing answers at all.' + "`r`n" +
                    'AutoUpdate -- the calendar is kept up to date but requests are still answered by a person.' + "`r`n" +
                    'AutoAccept -- the room answers for itself, using every setting below. This is what a bookable room normally is.')
$bkY += 54

[void](New-TabLabel $tabBooking 'EVERY OTHER BOOKING SETTING' 8 $bkY -Bold)
$bkY += 22
$gridBooking = New-TabGrid $tabBooking @(@('Setting', 260), @('Value', 200), @('Note', 420)) 8 $bkY 900 330
$bkBtnY = $bkY
$btnBkEdit    = New-TabButton $tabBooking 'Edit...'  920 $bkBtnY 130
$bkBtnY += 34
$btnBkRefresh = New-TabButton $tabBooking 'Refresh'  920 $bkBtnY 130
$bkY += 342
$bkNote = New-TabLabel $tabBooking '' 8 $bkY 1040 -Muted
$bkNote.Size = New-Object System.Drawing.Size(1040, 34)
$bkNote.Text = ('Double-click a row to edit it. One setting is written per Apply, so a value Exchange rejects cannot take the others down with it.' + "`r`n" +
                'The lists -- who may book, who must ask -- are on the Who Can Book tab, where the names are resolved to people.')

# The settings that belong to other tabs and would be two editors for one
# field if they were also editable here.
$script:BookingElsewhere = @(
    'BookInPolicy', 'AllBookInPolicy', 'RequestInPolicy', 'AllRequestInPolicy',
    'RequestOutOfPolicy', 'AllRequestOutOfPolicy', 'ResourceDelegates',
    'DeleteSubject', 'AddOrganizerToSubject', 'DeleteComments', 'DeleteAttachments',
    'DeleteNonCalendarItems', 'RemovePrivateProperty', 'EnableResponseDetails',
    'OrganizerInfo', 'AddAdditionalResponse', 'AdditionalResponse',
    'AutomateProcessing'
)
# Noise that is never worth a row.
$script:BookingExcluded = @(
    'PSComputerName', 'PSShowComputerName', 'RunspaceId', 'ObjectState',
    'Identity', 'MailboxOwnerId', 'OrganizationId', 'ExchangeVersion',
    'DistinguishedName', 'ObjectCategory', 'ObjectClass', 'WhenChanged',
    'WhenCreated', 'WhenChangedUTC', 'WhenCreatedUTC', 'Guid', 'IsValid'
)

function Get-RoomCalendarProcessing {
    # One read, shared by four tabs. Returns $null rather than throwing, so a
    # room whose policy cannot be read still opens and says so.
    param([string]$Smtp)
    try {
        return Invoke-EXOCommand -Name 'Get-CalendarProcessing' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
    } catch {
        $script:TabReadFailed = $true
        return $null
    }
}

function Get-RoomBookingNote {
    <#
        Why this row may be doing nothing.

        DELIBERATELY CONSERVATIVE. A warning that is sometimes wrong is worse
        than no warning -- people stop reading the column, and then miss the
        one that mattered.
    #>
    param([string]$Name, $Cal)
    if (-not $Cal) { return '' }
    if ($script:RoomAutoAcceptOnly -contains $Name -and "$($Cal.AutomateProcessing)" -ne 'AutoAccept') {
        return "INERT -- only applies while AutomateProcessing is AutoAccept."
    }
    if ($Name -eq 'BookingWindowInDays' -and [int]"$($Cal.BookingWindowInDays)" -eq 0) {
        return 'Zero means same-day only.'
    }
    if ($Name -eq 'MaximumDurationInMinutes' -and [int]"$($Cal.MaximumDurationInMinutes)" -eq 0) {
        return 'Zero means no limit.'
    }
    if ($Name -eq 'ScheduleOnlyDuringWorkHours' -and [bool]$Cal.ScheduleOnlyDuringWorkHours) {
        return 'Working hours and their timezone are on the Regional tab.'
    }
    if ($Name -eq 'EnforceCapacity' -and [bool]$Cal.EnforceCapacity) {
        return 'Measured against Seats on the Overview tab.'
    }
    return ''
}

function Update-BookingTab {
    $gridBooking.Rows.Clear()
    $c = $script:Cal
    if (-not $c) {
        $script:TabReadFailed = $true
        $bkAutoNote.Text = 'Calendar processing could not be read for this room.'
        return
    }
    $cmbAuto.SelectedItem = "$($c.AutomateProcessing)"

    # A HASHTABLE OF NAME -> TYPE, NOT A LIST OF NAMES. Wrapping it in @()
    # makes a one-element array holding the hashtable, so -contains never
    # matches and every row renders read-only -- which is to say the whole tab
    # silently stops being an editor.
    $writable = @{}
    try { $writable = Get-EXOWritableParameter -Name 'Set-CalendarProcessing' } catch { }

    foreach ($prop in @($c.PSObject.Properties | Sort-Object Name)) {
        $n = "$($prop.Name)"
        if ($script:BookingExcluded -contains $n) { continue }
        if ($script:BookingElsewhere -contains $n) { continue }
        $v = $prop.Value
        # A list here is a policy list and belongs on Who Can Book; showing it
        # as a flattened string would invite editing it as one.
        if ($v -is [System.Array] -and @($v).Count -gt 0) { continue }
        $i = $gridBooking.Rows.Add($n, "$v", (Get-RoomBookingNote -Name $n -Cal $c))
        $gridBooking.Rows[$i].Tag = [PSCustomObject]@{
            Name = $n; Value = $v; Writable = $writable.ContainsKey($n)
            Type = $(if ($writable.ContainsKey($n)) { $writable[$n] } else { $null })
        }
        if (-not $writable.ContainsKey($n)) {
            $gridBooking.Rows[$i].DefaultCellStyle.ForeColor = (Get-ToolStatusColour -Kind Muted)
        } elseif ("$($gridBooking.Rows[$i].Cells['Note'].Value)" -match 'INERT') {
            $gridBooking.Rows[$i].DefaultCellStyle.ForeColor = (Get-ToolStatusColour -Kind Warn)
        }
    }
    if ($writable.Keys.Count -eq 0) {
        $bkNote.Text = 'The writable parameter set came back empty, which means there is no Exchange Online session in this window rather than a cmdlet with no parameters. Nothing here can be edited until it connects.'
    }
}

function Invoke-BookingEdit {
    if ($gridBooking.SelectedRows.Count -eq 0) { Set-EditorStatus -Text 'Select a setting to edit.' -Bad; return }
    $t = $gridBooking.SelectedRows[0].Tag
    if (-not $t) { return }
    if (-not $t.Writable) {
        Set-EditorStatus -Text "'$($t.Name)' is not something Set-CalendarProcessing can change." -Bad
        return
    }

    $cur = "$($t.Value)"
    $new = $null
    if ($t.Value -is [bool]) {
        $a = Show-ToolDialog("Set $($t.Name) to True or False?`r`n`r`nIt is currently $cur.",
                             "Edit $($t.Name)", 'YesNoCancel', 'Question')
        if ($a -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
        $new = ($a -eq [System.Windows.Forms.DialogResult]::Yes)
    } elseif (($choices = Get-ParameterChoices -Type $t.Type -Name $t.Name)) {
        # A LIST, NOT A TEXT BOX, when the parameter accepts a known few. It is
        # 'AutoAccept', not 'Auto Accept' or 'autoaccept', and asking somebody
        # to type an enum value is asking them to get it wrong.
        $pick = Show-ChoiceDialog -Title "Edit $($t.Name)" -Prompt "New value for $($t.Name):" -Choices $choices -Current $cur
        if ($null -eq $pick) { return }
        $new = "$pick"
        if ("$new" -eq $cur) { Set-EditorStatus -Text 'Nothing changed.'; return }
    } else {
        # CANCEL COMES BACK AS AN EMPTY STRING, not $null. Testing for $null
        # let Cancel fall through as 'set it to empty', which on most of these
        # parameters means CLEAR IT -- pressing Cancel would have changed the
        # room. -Distinguish is how the two are told apart when clearing is a
        # legitimate answer; here it is not, so empty simply means no.
        $typed = Show-InputDialog -Title "Edit $($t.Name)" -Width 560 -Default $cur `
                    -Prompt "New value for $($t.Name):"
        if ("$typed" -eq '') { return }
        $new = "$typed".Trim()
        if ("$new" -eq $cur) { Set-EditorStatus -Text 'Nothing changed.'; return }
    }

    $warn = ''
    if ($script:RoomAutoAcceptOnly -contains $t.Name -and "$($script:Cal.AutomateProcessing)" -ne 'AutoAccept') {
        $warn = "`r`n`r`nNOTE: AutomateProcessing is $($script:Cal.AutomateProcessing), so this setting is stored but does nothing. Exchange will accept the change and the room's behaviour will not alter."
    }
    $a = Show-ToolDialog(("Set $($t.Name) on $($script:Room.Display)?" + "`r`n`r`nFrom: $cur" + "`r`n" + "To:   $new" + $warn),
                         'Change booking setting', 'OKCancel', 'Warning')
    if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

    # TYPED THROUGH THE CMDLET'S OWN VIEW OF THE PARAMETER, not guessed from
    # the string. It is 'AutoAccept', not 'Auto Accept', and an enum given the
    # wrong spelling fails in a way that names the type and not the value.
    #
    # -Value AND -Type ARE THE PARAMETERS. The type comes from the writable
    # map captured with the row; calling this with -Cmdlet/-Parameter throws a
    # binding error that a trailing catch swallows, and the raw string goes to
    # Exchange with the boolean guard defeated -- which is the single most
    # dangerous coercion here.
    $val = $new
    if ($t.Type) { $val = ConvertTo-CmdletParameterValue -Value $new -Type $t.Type }

    Invoke-EXOCommand -Name 'Set-CalendarProcessing' -Parameters @{
        Identity = $script:Room.Smtp; "$($t.Name)" = $val; ErrorAction = 'Stop'
    } | Out-Null
    Write-EditorAudit -Action 'RoomBookingSetting' -Detail "$($t.Name): '$cur' -> '$new'"
    $script:Cal = Get-RoomCalendarProcessing -Smtp $script:Room.Smtp
    Update-BookingTab
    $script:Loaded['Overview'] = $false
    Set-EditorStatus -Text "$($t.Name) is now $new."
}

$btnBkEdit.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction { Invoke-BookingEdit } 'Edit a booking setting'
})
$gridBooking.Add_CellDoubleClick({
    if ($_.RowIndex -lt 0) { return }
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction { Invoke-BookingEdit } 'Edit a booking setting'
})
$btnBkRefresh.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $script:Cal = Get-RoomCalendarProcessing -Smtp $script:Room.Smtp
        Update-BookingTab
        Set-EditorStatus -Text 'Booking settings re-read.'
    } 'Refresh booking settings'
})

$btnAutoApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $want = "$($cmbAuto.SelectedItem)"
        if (-not $want) { return }
        if ($want -eq "$($script:Cal.AutomateProcessing)") { Set-EditorStatus -Text "Already $want."; return }
        $extra = ''
        if ($want -eq 'None') {
            $extra = "`r`n`r`nWith None, nothing answers a request automatically. If this room has no delegates, requests will sit unanswered."
        }
        $a = Show-ToolDialog(("Set AutomateProcessing on $($script:Room.Display) from $($script:Cal.AutomateProcessing) to $want?" + $extra),
                             'Change automatic processing', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Invoke-EXOCommand -Name 'Set-CalendarProcessing' -Parameters @{
            Identity = $script:Room.Smtp; AutomateProcessing = $want; ErrorAction = 'Stop'
        } | Out-Null
        Write-EditorAudit -Action 'RoomAutomateProcessing' -Detail "AutomateProcessing -> $want"
        $script:Cal = Get-RoomCalendarProcessing -Smtp $script:Room.Smtp
        Update-BookingTab
        $script:Loaded['Overview'] = $false
        Set-EditorStatus -Text "AutomateProcessing is now $want."
    } 'Apply automatic processing'
})
#endregion

#region Who Can Book tab --------------------------------------------------------
<#
    THE FOUR LISTS, AND THE SWITCH THAT DECIDES WHETHER ANYONE READS THEM.

    THIS TAB EXISTS BECAUSE OF A REAL BUG. The original tool set BookInPolicy
    and never touched AllBookInPolicy, so the room carried on accepting
    bookings from everyone while holding a list of four names that nothing
    read. It confirmed the value was STORED; it could not confirm the value
    was USED. The paired switch is therefore on the same tab, next to the
    list, and the Apply writes both in one call.

    NAMES ARE RESOLVED, NOT TYPED BACK. Get-CalendarProcessing returns X500
    legacy distinguished names, not addresses. Shown raw they are unreadable;
    written back from a text box they are a coin toss. Every entry is resolved
    through Get-Recipient -- and an entry that will not resolve is KEPT
    VERBATIM and written back unchanged, so Apply can never silently drop
    somebody the tool merely failed to look up.
#>
$whY = 8
$script:RoomPolicyList = @(
    [PSCustomObject]@{ Key = 'BookInPolicy';       All = 'AllBookInPolicy';       Label = 'MAY BOOK DIRECTLY (in-policy requests are accepted)' }
    [PSCustomObject]@{ Key = 'RequestInPolicy';    All = 'AllRequestInPolicy';    Label = 'MUST ASK, even in-policy (the request goes to a delegate)' }
    [PSCustomObject]@{ Key = 'RequestOutOfPolicy'; All = 'AllRequestOutOfPolicy'; Label = 'MAY ASK for out-of-policy bookings' }
)

[void](New-TabLabel $tabWho 'WHO CAN BOOK THIS ROOM' 8 $whY -Bold)
$whY += 22
$whNote = New-TabLabel $tabWho '' 8 $whY 1040 -Muted
$whNote.Size = New-Object System.Drawing.Size(1040, 32)
$whNote.Text = ('Each list has a partner switch. While the switch is ticked the list is IGNORED and the rule applies to everybody -- that is Exchange''s design, not a bug here.' + "`r`n" +
                'A ticked switch with names in the list is the single most common reason a room "ignores" its booking restrictions.')
$whY += 40

[void](New-TabLabel $tabWho 'List:' 8 ($whY + 3) 40)
$cmbPolicy               = New-Object System.Windows.Forms.ComboBox
$cmbPolicy.DropDownStyle = 'DropDownList'
$cmbPolicy.Location      = New-Object System.Drawing.Point(52, $whY)
$cmbPolicy.Size          = New-Object System.Drawing.Size(420, 24)
foreach ($l in $script:RoomPolicyList) { [void]$cmbPolicy.Items.Add($l.Label) }
$cmbPolicy.SelectedIndex = 0
$tabWho.Controls.Add($cmbPolicy)

$chkAllPolicy = New-Object System.Windows.Forms.CheckBox
$chkAllPolicy.Location = New-Object System.Drawing.Point(490, ($whY + 2))
$chkAllPolicy.Size     = New-Object System.Drawing.Size(560, 20)
$chkAllPolicy.Text     = 'Applies to EVERYONE (the list below is ignored)'
$tabWho.Controls.Add($chkAllPolicy)
$whY += 32

$gridPolicy = New-TabGrid $tabWho @(@('Name', 380), @('Address', 300), @('State', 180)) 8 $whY 700 230
Add-GridPickColumn -Grid $gridPolicy

$whBtnY = $whY
$btnPolAdd    = New-TabButton $tabWho 'Add people...'  720 $whBtnY 170
$whBtnY += 34
$btnPolRemove = New-TabButton $tabWho 'Remove ticked'  720 $whBtnY 170
$whBtnY += 34
$btnPolAll    = New-TabButton $tabWho 'Tick all'       720 $whBtnY 82
$btnPolNone   = New-TabButton $tabWho 'None'           808 $whBtnY 82
$whBtnY += 42
$btnPolApply  = New-TabButton $tabWho 'Apply this list' 720 $whBtnY 170
$whY += 242

$whPolNote = New-TabLabel $tabWho '' 8 $whY 700 -Muted
$whPolNote.Size = New-Object System.Drawing.Size(700, 46)
$whY += 54

[void](New-TabLabel $tabWho 'RESOURCE DELEGATES' 8 $whY -Bold)
$whY += 22
$gridDeleg = New-TabGrid $tabWho @(@('Name', 380), @('Address', 300)) 8 $whY 700 130
Add-GridPickColumn -Grid $gridDeleg
$dgBtnY = $whY
$btnDelAdd    = New-TabButton $tabWho 'Add delegate...' 720 $dgBtnY 170
$dgBtnY += 34
$btnDelRemove = New-TabButton $tabWho 'Remove ticked'   720 $dgBtnY 170
$whY += 140
$whDelNote = New-TabLabel $tabWho '' 8 $whY 1040 -Muted
$whDelNote.Size = New-Object System.Drawing.Size(1040, 32)
$whDelNote.Text = ('Delegates answer anything the room will not answer for itself, and they are the only thing standing between "AutomateProcessing is None" and nobody answering at all.' + "`r`n" +
                   'ForwardRequestsToDelegates on the Booking tab decides whether requests actually reach them.')

function Get-RoomPolicySelected {
    return $script:RoomPolicyList[[Math]::Max(0, $cmbPolicy.SelectedIndex)]
}

function Resolve-RoomBooker {
    <#
        One entry from a policy list, turned into something readable.

        Get-Recipient, NOT Get-Mailbox: a policy list may legitimately hold a
        distribution group, and Get-Mailbox does not see one.

        A bare alias is retried against the tenant's mail domain. Two matches
        is refused rather than guessed -- picking the first would be a coin
        toss written into a permissions list.

        AN ENTRY THAT WILL NOT RESOLVE IS NOT AN ERROR AND IS NEVER DROPPED.
        It is carried through as-is with its state marked, and written back
        unchanged, because the alternative is a tool that quietly removes
        people it merely failed to look up.
    #>
    param([string]$Entry)
    $r = [PSCustomObject]@{ Raw = "$Entry"; Name = "$Entry"; Address = ''; State = 'unresolved' }
    if (-not "$Entry".Trim()) { return $r }

    foreach ($try in @("$Entry", "$Entry@$(try { Get-MailDomain } catch { '' })")) {
        if (-not "$try".Trim() -or "$try" -match '@$') { continue }
        try {
            $hit = @(Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{ Identity = $try; ErrorAction = 'Stop' })
            if (@($hit).Count -gt 1) {
                $r.State = 'AMBIGUOUS -- more than one recipient matches; left as it is'
                return $r
            }
            if (@($hit).Count -eq 1) {
                $r.Name    = "$($hit[0].DisplayName)"
                $r.Address = "$($hit[0].PrimarySmtpAddress)"
                $r.State   = 'ok'
                return $r
            }
        } catch { }
    }
    $r.State = 'could not be resolved -- kept as it is'
    return $r
}

function Update-WhoTab {
    $gridPolicy.Rows.Clear()
    $gridDeleg.Rows.Clear()
    $c = $script:Cal
    if (-not $c) { $script:TabReadFailed = $true; $whPolNote.Text = 'Calendar processing could not be read.'; return }

    $sel = Get-RoomPolicySelected
    $chkAllPolicy.Checked = [bool]$c."$($sel.All)"

    foreach ($e in @($c."$($sel.Key)")) {
        $res = Resolve-RoomBooker -Entry "$e"
        $i = $gridPolicy.Rows.Add($false, $res.Name, $res.Address, $res.State)
        $gridPolicy.Rows[$i].Tag = $res
        if ($res.State -ne 'ok') {
            $gridPolicy.Rows[$i].DefaultCellStyle.ForeColor = (Get-ToolStatusColour -Kind Warn)
        }
    }
    foreach ($e in @($c.ResourceDelegates)) {
        $res = Resolve-RoomBooker -Entry "$e"
        $i = $gridDeleg.Rows.Add($false, $res.Name, $res.Address)
        $gridDeleg.Rows[$i].Tag = $res
    }

    # THE FINDING, SAID OUT LOUD, every time the tab is drawn.
    if ($chkAllPolicy.Checked -and $gridPolicy.Rows.Count -gt 0) {
        $whPolNote.Text = "THE $($gridPolicy.Rows.Count) NAME(S) ABOVE ARE NOT BEING ENFORCED. '$($sel.All)' is True, so this rule applies to everybody and the list is ignored. Untick the box and Apply to make the list mean something."
        $whPolNote.ForeColor = (Get-ToolStatusColour -Kind Bad)
    } elseif (-not $chkAllPolicy.Checked -and $gridPolicy.Rows.Count -eq 0) {
        $whPolNote.Text = "The list is empty and '$($sel.All)' is False, so this rule currently applies to NOBODY."
        $whPolNote.ForeColor = (Get-ToolStatusColour -Kind Warn)
    } else {
        $whPolNote.Text = "$($gridPolicy.Rows.Count) name(s). Apply writes the list and the switch together -- writing one without the other is what made a room look restricted while accepting everything."
        $whPolNote.ForeColor = (Get-ToolStatusColour -Kind Muted)
    }
}

function Invoke-PolicyApply {
    $sel = Get-RoomPolicySelected
    # WRITTEN BACK FROM THE ROW TAGS, NOT FROM THE DISPLAYED TEXT. The Name
    # column holds a display name; the value Exchange stores is the raw entry
    # it gave us, and that is what goes back.
    $entries = @()
    foreach ($row in $gridPolicy.Rows) {
        if ($row.Tag) { $entries += "$(if ($row.Tag.Address) { $row.Tag.Address } else { $row.Tag.Raw })" }
    }
    $restrict = -not $chkAllPolicy.Checked

    # AN UNRESOLVED ENTRY IS CARRIED THROUGH, AND THAT IS ALL-OR-NOTHING.
    # Set-CalendarProcessing validates the whole -BookInPolicy array: if one
    # entry names a recipient that no longer exists, the entire write is
    # rejected and nothing changes. Better to say so before the attempt than
    # to have 'never drops anyone' turn into 'cannot Apply' with no reason.
    $unresolved = @()
    foreach ($row in $gridPolicy.Rows) {
        if ($row.Tag -and $row.Tag.State -ne 'ok' -and $row.Tag.State -notmatch 'not yet applied') {
            $unresolved += "$($row.Tag.Raw)"
        }
    }
    $extra = ''
    if ($unresolved.Count) {
        $extra += "`r`n`r`n$($unresolved.Count) entry/entries could not be resolved and are being written back unchanged:" + "`r`n  " + ($unresolved -join "`r`n  ") +
                  "`r`n`r`nExchange validates the whole list. If one of these names a recipient that no longer exists the ENTIRE write is rejected and nothing changes -- remove them and try again if that happens."
    }
    if ($restrict -and $entries.Count -eq 0) {
        $extra += "`r`n`r`nTHE LIST IS EMPTY AND YOU ARE RESTRICTING TO IT. Nobody will be able to book this room under this rule."
    }
    if (-not $restrict -and $entries.Count -gt 0) {
        $extra += "`r`n`r`nYou are leaving the rule open to everybody while keeping $($entries.Count) name(s) in the list. The names will be stored and ignored."
    }
    $a = Show-ToolDialog(
        ("On $($script:Room.Display):" + "`r`n`r`n" +
         "$($sel.Key): $(if ($entries.Count) { ($entries -join '; ') } else { '(empty)' })" + "`r`n" +
         "$($sel.All): $(-not $restrict)" + $extra),
        'Apply booking policy', 'OKCancel', 'Warning')
    if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

    # ONE CALL, BOTH PARAMETERS. This is the whole lesson of the tool this
    # replaced.
    $p = @{ Identity = $script:Room.Smtp; ErrorAction = 'Stop' }
    $p["$($sel.Key)"] = @($entries)
    $p["$($sel.All)"] = (-not $restrict)
    Invoke-EXOCommand -Name 'Set-CalendarProcessing' -Parameters $p | Out-Null

    Write-EditorAudit -Action 'RoomBookingPolicy' `
        -Detail "$($sel.Key) = $(if ($entries) { $entries -join '; ' } else { '(empty)' }); $($sel.All) = $(-not $restrict)"
    $script:Cal = Get-RoomCalendarProcessing -Smtp $script:Room.Smtp
    Update-WhoTab
    $script:Loaded['Overview'] = $false
    Set-EditorStatus -Text "$($sel.Key) applied, with $($sel.All) set to $(-not $restrict)."
}

$cmbPolicy.Add_SelectedIndexChanged({
    if ($script:Busy) { return }
    if (-not $script:Room -or -not $script:Room.Ok) { return }
    Invoke-EditorAction { Update-WhoTab } 'Read the booking policy list'
})
$btnPolAll.Add_Click({  Set-GridPicked -Grid $gridPolicy -On $true })
$btnPolNone.Add_Click({ Set-GridPicked -Grid $gridPolicy -On $false })

$btnPolAdd.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $who = Show-InputDialog -Title 'Add to the list' -Width 620 `
                -Prompt 'Who may book? Addresses or aliases, separated by semicolons. A distribution group is allowed.'
        $list = ConvertFrom-DialogList -Text "$who"
        if ($list.Count -eq 0) { return }
        $added = 0
        foreach ($w in $list) {
            $res = Resolve-RoomBooker -Entry $w
            if ($res.State -ne 'ok') {
                Set-EditorStatus -Text "'$w' $($res.State). Nothing was added." -Bad
                return
            }
            $dupe = $false
            foreach ($row in $gridPolicy.Rows) { if ($row.Tag -and $row.Tag.Address -ieq $res.Address) { $dupe = $true } }
            if ($dupe) { continue }
            $i = $gridPolicy.Rows.Add($false, $res.Name, $res.Address, 'ok (not yet applied)')
            $gridPolicy.Rows[$i].Tag = $res
            $added++
        }
        Set-EditorStatus -Text "$added added to the list on screen. Press 'Apply this list' to write it."
    } 'Add to the booking policy list'
})

$btnPolRemove.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $picked = Get-GridPicked -Grid $gridPolicy
        if ($picked.Count -eq 0) { Set-EditorStatus -Text 'Tick one or more names to remove.' -Bad; return }
        for ($i = $gridPolicy.Rows.Count - 1; $i -ge 0; $i--) {
            if ([bool]$gridPolicy.Rows[$i].Cells['Pick'].Value) { $gridPolicy.Rows.RemoveAt($i) }
        }
        Set-EditorStatus -Text "$($picked.Count) removed from the list on screen. Press 'Apply this list' to write it."
    } 'Remove from the booking policy list'
})

$btnPolApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction { Invoke-PolicyApply } 'Apply the booking policy'
})

$btnDelAdd.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $who = Show-InputDialog -Title 'Add delegates' -Width 620 `
                -Prompt 'Who should answer requests for this room? Addresses or aliases, separated by semicolons.'
        $list = ConvertFrom-DialogList -Text "$who"
        if ($list.Count -eq 0) { return }
        $current = @()
        foreach ($row in $gridDeleg.Rows) { if ($row.Tag) { $current += "$(if ($row.Tag.Address) { $row.Tag.Address } else { $row.Tag.Raw })" } }
        foreach ($w in $list) {
            $res = Resolve-RoomBooker -Entry $w
            if ($res.State -ne 'ok') { Set-EditorStatus -Text "'$w' $($res.State). Nothing was added." -Bad; return }
            if ($current -notcontains $res.Address) { $current += $res.Address }
        }
        $a = Show-ToolDialog(("Set the delegates on $($script:Room.Display) to:" + "`r`n`r`n  " + ($current -join "`r`n  ")),
                             'Resource delegates', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        # ResourceDelegates is replaced whole, like the policy lists -- so the
        # write is always every delegate, not the new one.
        Invoke-EXOCommand -Name 'Set-CalendarProcessing' -Parameters @{
            Identity = $script:Room.Smtp; ResourceDelegates = @($current); ErrorAction = 'Stop'
        } | Out-Null
        Write-EditorAudit -Action 'RoomDelegates' -Detail "ResourceDelegates = $($current -join '; ')"
        $script:Cal = Get-RoomCalendarProcessing -Smtp $script:Room.Smtp
        Update-WhoTab
        $script:Loaded['Overview'] = $false
        Set-EditorStatus -Text 'Delegates applied.'
    } 'Add resource delegates'
})

$btnDelRemove.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $picked = Get-GridPicked -Grid $gridDeleg
        if ($picked.Count -eq 0) { Set-EditorStatus -Text 'Tick one or more delegates to remove.' -Bad; return }
        $keep = @()
        foreach ($row in $gridDeleg.Rows) {
            if ([bool]$row.Cells['Pick'].Value) { continue }
            if ($row.Tag) { $keep += "$(if ($row.Tag.Address) { $row.Tag.Address } else { $row.Tag.Raw })" }
        }
        $names = @($picked | ForEach-Object { "  $($_.Name)" }) -join "`r`n"
        $warn = ''
        if ($keep.Count -eq 0 -and "$($script:Cal.AutomateProcessing)" -eq 'None') {
            $warn = "`r`n`r`nTHIS REMOVES THE LAST DELEGATE while AutomateProcessing is None. Nothing and nobody will answer a request for this room."
        }
        $a = Show-ToolDialog(("Remove $($picked.Count) delegate(s) from $($script:Room.Display)?" + "`r`n`r`n" + $names + $warn),
                             'Remove delegates', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Invoke-EXOCommand -Name 'Set-CalendarProcessing' -Parameters @{
            Identity = $script:Room.Smtp; ResourceDelegates = @($keep); ErrorAction = 'Stop'
        } | Out-Null
        Write-EditorAudit -Action 'RoomDelegates' -Detail "ResourceDelegates = $(if ($keep) { $keep -join '; ' } else { '(none)' })"
        $script:Cal = Get-RoomCalendarProcessing -Smtp $script:Room.Smtp
        Update-WhoTab
        $script:Loaded['Overview'] = $false
        Set-EditorStatus -Text "Removed $($picked.Count) delegate(s)."
    } 'Remove resource delegates'
})
#endregion

#region Replies tab -------------------------------------------------------------
<#
    WHAT THE ROOM STRIPS FROM A REQUEST, AND WHAT IT SAYS BACK.

    These are the settings behind "why does the meeting in the room calendar
    just say 'Busy'" and "why has the room stopped telling people why it
    declined". They are all Set-CalendarProcessing, but they are a different
    question from the booking policy and they belong on their own tab.
#>
$rpY = 8
[void](New-TabLabel $tabReplies 'WHAT THE ROOM KEEPS' 8 $rpY -Bold)
$rpY += 24

$script:RoomReplyFlag = @(
    [PSCustomObject]@{ Name = 'DeleteSubject';            Label = 'Delete the subject of the request' }
    [PSCustomObject]@{ Name = 'AddOrganizerToSubject';    Label = 'Put the organiser''s name in the subject' }
    [PSCustomObject]@{ Name = 'DeleteComments';           Label = 'Delete the body of the request' }
    [PSCustomObject]@{ Name = 'DeleteAttachments';        Label = 'Delete attachments' }
    [PSCustomObject]@{ Name = 'DeleteNonCalendarItems';   Label = 'Delete anything that is not a meeting request' }
    [PSCustomObject]@{ Name = 'RemovePrivateProperty';    Label = 'Clear the Private flag on accepted meetings' }
    [PSCustomObject]@{ Name = 'EnableResponseDetails';    Label = 'Say WHY a request was declined' }
    [PSCustomObject]@{ Name = 'OrganizerInfo';            Label = 'Tell a declined organiser who already has the room' }
)
$script:RoomReplyCombo = @{}
foreach ($fl in $script:RoomReplyFlag) {
    [void](New-TabLabel $tabReplies $fl.Label 8 ($rpY + 3) 420)
    $script:RoomReplyCombo[$fl.Name] = Get-RoomBool $tabReplies 440 $rpY
    [void](New-TabLabel $tabReplies $fl.Name 560 ($rpY + 3) 300 -Muted)
    $rpY += 26
}
$rpY += 8
$rpNote1 = New-TabLabel $tabReplies '' 8 $rpY 1040 -Muted
$rpNote1.Size = New-Object System.Drawing.Size(1040, 46)
$rpNote1.Text = ('DeleteSubject and AddOrganizerToSubject together are what make a room calendar read "Jane Smith" instead of the real subject -- which is usually what people want and occasionally the thing they are complaining about.' + "`r`n" +
                 'EnableResponseDetails off is why a declined request arrives with no reason. Turning it on is almost always the fix for "the room just says no".')
$rpY += 56

[void](New-TabLabel $tabReplies 'WHAT THE ROOM SAYS BACK' 8 $rpY -Bold)
$rpY += 24
[void](New-TabLabel $tabReplies 'Add an extra message to every reply:' 8 ($rpY + 3) 250)
$cmbAddResp = Get-RoomBool $tabReplies 264 $rpY
$rpY += 28
$txtAddResp            = New-Object System.Windows.Forms.TextBox
$txtAddResp.Multiline  = $true
$txtAddResp.ScrollBars = 'Vertical'
$txtAddResp.Location   = New-Object System.Drawing.Point(8, $rpY)
$txtAddResp.Size       = New-Object System.Drawing.Size(1040, 120)
$tabReplies.Controls.Add($txtAddResp)
$rpY += 128
$rpNote2 = New-TabLabel $tabReplies '' 8 $rpY 1040 -Muted
$rpNote2.Size = New-Object System.Drawing.Size(1040, 32)
$rpNote2.Text = ('The extra message is sent on every reply this room makes -- accept and decline alike. It is the place for "this room has no video kit" or "call extension 240 for the key".' + "`r`n" +
                 'Exchange stores it as HTML; anything typed here is sent as typed.')
$rpY += 40
$btnRpApply = New-TabButton $tabReplies 'Apply replies' 8 $rpY 160

function Update-RepliesTab {
    $c = $script:Cal
    if (-not $c) { $script:TabReadFailed = $true; return }
    foreach ($fl in $script:RoomReplyFlag) {
        Set-RoomBool $script:RoomReplyCombo[$fl.Name] $c."$($fl.Name)"
    }
    Set-RoomBool $cmbAddResp $c.AddAdditionalResponse
    $txtAddResp.Text = "$($c.AdditionalResponse)"
}

$btnRpApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $c = $script:Cal
        $p = @{ Identity = $script:Room.Smtp; ErrorAction = 'Stop' }
        $changed = @()
        foreach ($fl in $script:RoomReplyFlag) {
            $want = Get-RoomBoolValue $script:RoomReplyCombo[$fl.Name]
            if ($want -ne [bool]$c."$($fl.Name)") { $p[$fl.Name] = $want; $changed += "$($fl.Name)=$want" }
        }
        $wantResp = Get-RoomBoolValue $cmbAddResp
        if ($wantResp -ne [bool]$c.AddAdditionalResponse) { $p['AddAdditionalResponse'] = $wantResp; $changed += "AddAdditionalResponse=$wantResp" }
        if ("$($txtAddResp.Text)" -ne "$($c.AdditionalResponse)") {
            $p['AdditionalResponse'] = "$($txtAddResp.Text)"
            $changed += 'AdditionalResponse'
        }
        if ($changed.Count -eq 0) { Set-EditorStatus -Text 'Nothing changed.'; return }

        # EXCHANGE REFUSES AdditionalResponse WITHOUT AddAdditionalResponse.
        # Sending the text while the switch is off fails with a message about
        # the parameter rather than about the switch, which reads as the text
        # being wrong.
        if ($p.ContainsKey('AdditionalResponse') -and -not $p.ContainsKey('AddAdditionalResponse') -and -not [bool]$c.AddAdditionalResponse) {
            Set-EditorStatus -Text 'Turn the extra message on before setting its text -- Exchange will not store the text while it is off.' -Bad
            return
        }
        Invoke-EXOCommand -Name 'Set-CalendarProcessing' -Parameters $p | Out-Null
        Write-EditorAudit -Action 'RoomReplySettings' -Detail ($changed -join '; ')
        $script:Cal = Get-RoomCalendarProcessing -Smtp $script:Room.Smtp
        Update-RepliesTab
        Set-EditorStatus -Text "Applied: $($changed -join '; ')"
    } 'Apply reply settings'
})
#endregion

#region Place tab ---------------------------------------------------------------
<#
    THE RECORD ROOM FINDER ACTUALLY READS.

    Get-Place / Set-Place is a different object from the mailbox, and it is
    the one Outlook's Room Finder uses for building, floor, capacity and
    features. A room can book perfectly by address and still be invisible in
    Room Finder because this is empty.

    TWO SAFEGUARDS, BOTH LEARNED THE HARD WAY.

    ONE -- the address block is refused on a synced room before the call is
    made. Set-Place does not always error on those fields: it can accept the
    call and simply not persist the value, which is the worst possible
    outcome, a success message and no change.

    TWO -- every write is read back and compared. Set-Place reports success
    either way, so asking for the value afterwards is the only way to know it
    landed.

    ONE PROPERTY PER CALL, because Set-Place is documented as failing on
    certain COMBINATIONS of properties in older module versions.
#>
$plY = 8
[void](New-TabLabel $tabPlace 'ROOM FINDER RECORD' 8 $plY -Bold)
$plY += 22
$plNote = New-TabLabel $tabPlace '' 8 $plY 1040 -Muted
$plNote.Size = New-Object System.Drawing.Size(1040, 32)
$plY += 40
$gridPlace = New-TabGrid $tabPlace @(@('Setting', 240), @('Value', 340), @('State', 300)) 8 $plY 900 400
$plBtnY = $plY
$btnPlEdit    = New-TabButton $tabPlace 'Edit...' 920 $plBtnY 130
$plBtnY += 34
$btnPlRefresh = New-TabButton $tabPlace 'Refresh' 920 $plBtnY 130
$plY += 412
$plHint = New-TabLabel $tabPlace '' 8 $plY 1040 -Muted
$plHint.Size = New-Object System.Drawing.Size(1040, 32)
$plHint.Text = ('Double-click a row to edit it. Every change is written on its own and then read back -- a value that did not stick is reported, not assumed.' + "`r`n" +
                'Room Finder caches this; allow up to 24 hours before deciding a change has not worked.')

$script:PlaceExcluded = @(
    'PSComputerName', 'PSShowComputerName', 'RunspaceId', 'ObjectState', 'Identity',
    'OrganizationId', 'ExchangeVersion', 'DistinguishedName', 'ObjectCategory',
    'ObjectClass', 'WhenChanged', 'WhenCreated', 'WhenChangedUTC', 'WhenCreatedUTC',
    'Guid', 'IsValid', 'PlaceId', 'MailboxGuid'
)

function Update-PlaceTab {
    $gridPlace.Rows.Clear()
    $pl = $null
    try {
        $pl = Invoke-EXOCommand -Name 'Get-Place' -Parameters @{ Identity = $script:Room.Smtp; ErrorAction = 'Stop' }
    } catch {
        $script:TabReadFailed = $true
        $plNote.Text = "The Place record could not be read: $($_.Exception.Message)"
        $plNote.ForeColor = (Get-ToolStatusColour -Kind Bad)
        return
    }
    $writable = @{}
    try { $writable = Get-EXOWritableParameter -Name 'Set-Place' } catch { }
    if ($writable.Keys.Count -eq 0) {
        $plNote.Text = 'The writable parameter set came back empty. That means no Exchange Online session in this window, not a cmdlet without parameters -- nothing here can be edited until it connects.'
        $plNote.ForeColor = (Get-ToolStatusColour -Kind Bad)
    } else {
        $plNote.Text = 'Building, floor, capacity and the device flags are what Room Finder filters on. A room with none of these set is bookable by address and hard to find by browsing.'
        $plNote.ForeColor = (Get-ToolStatusColour -Kind Muted)
    }

    foreach ($prop in @($pl.PSObject.Properties | Sort-Object Name)) {
        $n = "$($prop.Name)"
        if ($script:PlaceExcluded -contains $n) { continue }
        $v = $prop.Value
        $shown = $(if ($v -is [System.Array]) { (@($v) -join ', ') } else { "$v" })

        $state = 'read-only'
        if ($writable.ContainsKey($n)) {
            if ($script:Room.IsSynced -and $script:PlaceHybridBlocked -contains $n) {
                $state = 'mastered on-premises -- change it in AD'
            } else {
                $state = 'editable'
            }
        }
        $i = $gridPlace.Rows.Add($n, $shown, $state)
        $gridPlace.Rows[$i].Tag = [PSCustomObject]@{
            Name = $n; Value = $v; State = $state
            Type = $(if ($writable.ContainsKey($n)) { $writable[$n] } else { $null })
        }
        if ($state -eq 'mastered on-premises -- change it in AD') {
            $gridPlace.Rows[$i].DefaultCellStyle.ForeColor = (Get-ToolStatusColour -Kind Warn)
        } elseif ($state -eq 'read-only') {
            $gridPlace.Rows[$i].DefaultCellStyle.ForeColor = (Get-ToolStatusColour -Kind Muted)
        }
    }
}

function Invoke-PlaceEdit {
    if ($gridPlace.SelectedRows.Count -eq 0) { Set-EditorStatus -Text 'Select a setting to edit.' -Bad; return }
    $t = $gridPlace.SelectedRows[0].Tag
    if (-not $t) { return }
    if ($t.State -ne 'editable') {
        $m = $(if ($t.State -match 'on-premises') {
                "'$($t.Name)' is mastered in Active Directory for this synced room. Set-Place can accept the call and silently not persist it, so this window refuses it rather than reporting a success that is not one. Change it on-premises."
              } else {
                "'$($t.Name)' is not something Set-Place can change."
              })
        $null = Show-ToolDialog($m, 'Cannot be changed here', 'OK', 'Warning')
        Set-EditorStatus -Text $m -Bad
        return
    }

    $cur = $(if ($t.Value -is [System.Array]) { (@($t.Value) -join ', ') } else { "$($t.Value)" })
    $new = $null
    if ($t.Value -is [bool]) {
        $a = Show-ToolDialog("Set $($t.Name) to True or False?`r`n`r`nIt is currently $cur.",
                             "Edit $($t.Name)", 'YesNoCancel', 'Question')
        if ($a -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
        $new = ($a -eq [System.Windows.Forms.DialogResult]::Yes)
    } elseif (-not ($t.Value -is [System.Array]) -and ($choices = Get-ParameterChoices -Type $t.Type -Name $t.Name)) {
        $pick = Show-ChoiceDialog -Title "Edit $($t.Name)" -Prompt "New value for $($t.Name):" -Choices $choices -Current $cur
        if ($null -eq $pick) { return }
        $new = "$pick"
        if ("$new" -eq $cur) { Set-EditorStatus -Text 'Nothing changed.'; return }
    } else {
        # Cancel is an empty string, not $null -- see the Booking tab.
        $typed = Show-InputDialog -Title "Edit $($t.Name)" -Width 560 -Default $cur `
                    -Prompt "New value for $($t.Name)$(if ($t.Value -is [System.Array]) { ' (separate several with commas)' }):"
        if ("$typed" -eq '') { return }
        $new = "$typed".Trim()
        if ("$new" -eq $cur) { Set-EditorStatus -Text 'Nothing changed.'; return }
    }

    $val = $new
    if ($t.Value -is [System.Array]) {
        $val = @("$new" -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    } elseif ($t.Type) {
        $val = ConvertTo-CmdletParameterValue -Value $new -Type $t.Type
    }

    Invoke-EXOCommand -Name 'Set-Place' -Parameters @{
        Identity = $script:Room.Smtp; "$($t.Name)" = $val; ErrorAction = 'Stop'
    } | Out-Null

    # READ BACK AND COMPARE. Set-Place reports success either way.
    $landed = $false
    $after  = ''
    try {
        $chk = Invoke-EXOCommand -Name 'Get-Place' -Parameters @{ Identity = $script:Room.Smtp; ErrorAction = 'Stop' }
        $got = $chk."$($t.Name)"
        # Tags and the like are compared sorted and joined, so a different
        # ORDER coming back is not reported as a failure.
        $a1 = $(if ($got -is [System.Array]) { (@($got) | Sort-Object) -join '|' } else { "$got" })
        $b1 = $(if ($val -is [System.Array]) { (@($val) | Sort-Object) -join '|' } else { "$val" })
        $after = $(if ($got -is [System.Array]) { (@($got) -join ', ') } else { "$got" })
        $landed = ($a1 -eq $b1)
    } catch { }

    if ($landed) {
        Write-EditorAudit -Action 'RoomPlaceSetting' -Detail "$($t.Name): '$cur' -> '$new' (verified)"
        Set-EditorStatus -Text "$($t.Name) is now $after."
    } else {
        Write-EditorAudit -Action 'RoomPlaceSetting' -Result 'Failed' `
            -Detail "$($t.Name): asked for '$new', reads back as '$after'"
        Set-EditorStatus -Text "Set-Place reported success but $($t.Name) reads back as '$after'. The value did NOT stick -- that is usually a synced field." -Bad
        $null = Show-ToolDialog(
            ("Exchange accepted the change to $($t.Name) and the value did not stick." + "`r`n`r`n" +
             "Asked for: $new" + "`r`n" + "Reads back: $after" + "`r`n`r`n" +
             'Set-Place reports success either way, which is why this window reads it back. A field that behaves like this is almost always mastered on-premises.'),
            'The change did not persist', 'OK', 'Warning')
    }
    Update-PlaceTab
}

$btnPlEdit.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction { Invoke-PlaceEdit } 'Edit a Place setting'
})
$gridPlace.Add_CellDoubleClick({
    if ($_.RowIndex -lt 0) { return }
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction { Invoke-PlaceEdit } 'Edit a Place setting'
})
$btnPlRefresh.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction { Update-PlaceTab; Set-EditorStatus -Text 'Place record re-read.' } 'Refresh the Place record'
})
#endregion

#region Regional tab ------------------------------------------------------------
<#
    TIMEZONE, LANGUAGE AND THE WORKING HOURS.

    ScheduleOnlyDuringWorkHours on the Booking tab depends on these ENTIRELY,
    and the two halves live in two different cmdlets -- regional configuration
    and calendar configuration -- with a timezone in each. A room whose
    regional timezone and whose working-hours timezone disagree declines
    requests at hours nobody can account for.

    THE WRITABLE MAP IS KEYED 'Source:Name', because both cmdlets have a
    property whose name ends in TimeZone. Keyed on the bare name, setting the
    working-hours timezone would write the regional one.

    TIMEZONES ARE WINDOWS IDS. Exchange takes 'Central Standard Time', not
    'CST' and not an IANA name, so the list is built from the machine rather
    than typed.
#>
$rgY = 8
[void](New-TabLabel $tabRegional 'REGIONAL' 8 $rgY -Bold)
$rgY += 24
[void](New-TabLabel $tabRegional 'Timezone:' 8 ($rgY + 3) 150)
$cmbTz               = New-Object System.Windows.Forms.ComboBox
$cmbTz.DropDownStyle = 'DropDownList'
$cmbTz.Location      = New-Object System.Drawing.Point(170, $rgY)
$cmbTz.Size          = New-Object System.Drawing.Size(420, 24)
$tabRegional.Controls.Add($cmbTz)
$rgY += 28
[void](New-TabLabel $tabRegional 'Language:' 8 ($rgY + 3) 150)
$txtLang          = New-Object System.Windows.Forms.TextBox
$txtLang.Location = New-Object System.Drawing.Point(170, $rgY)
$txtLang.Size     = New-Object System.Drawing.Size(180, 24)
$tabRegional.Controls.Add($txtLang)
[void](New-TabLabel $tabRegional 'Date format:' 380 ($rgY + 3) 90)
$txtDateFmt          = New-Object System.Windows.Forms.TextBox
$txtDateFmt.Location = New-Object System.Drawing.Point(476, $rgY)
$txtDateFmt.Size     = New-Object System.Drawing.Size(140, 24)
$tabRegional.Controls.Add($txtDateFmt)
[void](New-TabLabel $tabRegional 'Time format:' 636 ($rgY + 3) 90)
$txtTimeFmt          = New-Object System.Windows.Forms.TextBox
$txtTimeFmt.Location = New-Object System.Drawing.Point(732, $rgY)
$txtTimeFmt.Size     = New-Object System.Drawing.Size(140, 24)
$tabRegional.Controls.Add($txtTimeFmt)
$rgY += 28
$btnRgApply = New-TabButton $tabRegional 'Apply regional' 170 $rgY 160
$rgY += 34
$rgNote = New-TabLabel $tabRegional '' 8 $rgY 1040 -Muted
$rgNote.Size = New-Object System.Drawing.Size(1040, 32)
$rgNote.Text = 'Language is a culture name such as en-GB or en-US. Leave a box as it is to keep the current value.'
$rgY += 44

[void](New-TabLabel $tabRegional 'WORKING HOURS' 8 $rgY -Bold)
$rgY += 24
# THE ORDER HERE IS LOAD-BEARING. Sorted alphabetically, End comes before
# Start, which reads as a mistake every time somebody opens the tab.
[void](New-TabLabel $tabRegional 'Start:' 8 ($rgY + 3) 150)
$txtWhStart          = New-Object System.Windows.Forms.TextBox
$txtWhStart.Location = New-Object System.Drawing.Point(170, $rgY)
$txtWhStart.Size     = New-Object System.Drawing.Size(140, 24)
$tabRegional.Controls.Add($txtWhStart)
[void](New-TabLabel $tabRegional 'End:' 340 ($rgY + 3) 60)
$txtWhEnd          = New-Object System.Windows.Forms.TextBox
$txtWhEnd.Location = New-Object System.Drawing.Point(406, $rgY)
$txtWhEnd.Size     = New-Object System.Drawing.Size(140, 24)
$tabRegional.Controls.Add($txtWhEnd)
$rgY += 28
[void](New-TabLabel $tabRegional 'Working-hours timezone:' 8 ($rgY + 3) 150)
$cmbWhTz               = New-Object System.Windows.Forms.ComboBox
$cmbWhTz.DropDownStyle = 'DropDownList'
$cmbWhTz.Location      = New-Object System.Drawing.Point(170, $rgY)
$cmbWhTz.Size          = New-Object System.Drawing.Size(420, 24)
$tabRegional.Controls.Add($cmbWhTz)
$rgY += 28
[void](New-TabLabel $tabRegional 'Work days:' 8 ($rgY + 3) 150)
# A DROPDOWN, NOT A LIST: WorkDays is a flags enum, so 'Monday, Tuesday' is a
# legitimate value and has to be typeable.
$cmbWorkDays               = New-Object System.Windows.Forms.ComboBox
$cmbWorkDays.DropDownStyle = 'DropDown'
$cmbWorkDays.Location      = New-Object System.Drawing.Point(170, $rgY)
$cmbWorkDays.Size          = New-Object System.Drawing.Size(280, 24)
foreach ($d in @('Weekdays', 'AllDays', 'WeekEndDays', 'Monday, Tuesday, Wednesday, Thursday', 'None')) { [void]$cmbWorkDays.Items.Add($d) }
$tabRegional.Controls.Add($cmbWorkDays)
[void](New-TabLabel $tabRegional 'Week starts:' 480 ($rgY + 3) 90)
$cmbWeekStart               = New-Object System.Windows.Forms.ComboBox
$cmbWeekStart.DropDownStyle = 'DropDownList'
$cmbWeekStart.Location      = New-Object System.Drawing.Point(576, $rgY)
$cmbWeekStart.Size          = New-Object System.Drawing.Size(160, 24)
foreach ($d in @('Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday')) { [void]$cmbWeekStart.Items.Add($d) }
$tabRegional.Controls.Add($cmbWeekStart)
$rgY += 30
$btnWhApply = New-TabButton $tabRegional 'Apply working hours' 170 $rgY 180
$rgY += 36
$rgWarn = New-TabLabel $tabRegional '' 8 $rgY 1040 -Muted
$rgWarn.Size = New-Object System.Drawing.Size(1040, 64)
$rgY += 72
$rgHint = New-TabLabel $tabRegional '' 8 $rgY 1040 -Muted
$rgHint.Size = New-Object System.Drawing.Size(1040, 46)
$rgHint.Text = ('Times are as Exchange stores them, usually hh:mm:ss. A timezone change re-interprets FUTURE requests only -- meetings already in the calendar do not move.' + "`r`n" +
                'Room Finder and Outlook cache these; a change can take a while to show.')

function Get-RoomTimeZoneList {
    # Built from the machine, once. Exchange wants the Windows id, and typing
    # one is how you get 'Central Standard Time' spelled four different ways.
    if ($script:RoomTzList) { return $script:RoomTzList }
    $script:RoomTzList = @([System.TimeZoneInfo]::GetSystemTimeZones() | ForEach-Object { "$($_.Id)" } | Sort-Object)
    return $script:RoomTzList
}

function Set-RoomTzCombo {
    param($Combo, [string]$Value)
    if ($Combo.Items.Count -eq 0) {
        [void]$Combo.Items.Add('')
        foreach ($z in (Get-RoomTimeZoneList)) { [void]$Combo.Items.Add($z) }
    }
    $i = $Combo.Items.IndexOf("$Value")
    $Combo.SelectedIndex = $(if ($i -ge 0) { $i } else { 0 })
}

function Update-RegionalTab {
    $reg = $null; $cal = $null
    try { $reg = Invoke-EXOCommand -Name 'Get-MailboxRegionalConfiguration' -Parameters @{ Identity = $script:Room.Smtp; ErrorAction = 'Stop' } }
    catch { $script:TabReadFailed = $true }
    try { $cal = Invoke-EXOCommand -Name 'Get-MailboxCalendarConfiguration' -Parameters @{ Identity = $script:Room.Smtp; ErrorAction = 'Stop' } }
    catch { $script:TabReadFailed = $true }
    $script:Reg = $reg
    $script:CalCfg = $cal

    Set-RoomTzCombo $cmbTz   "$(if ($reg) { $reg.TimeZone } else { '' })"
    Set-RoomTzCombo $cmbWhTz "$(if ($cal) { $cal.WorkingHoursTimeZone } else { '' })"
    $txtLang.Text    = "$(if ($reg) { $reg.Language } else { '' })"
    $txtDateFmt.Text = "$(if ($reg) { $reg.DateFormat } else { '' })"
    $txtTimeFmt.Text = "$(if ($reg) { $reg.TimeFormat } else { '' })"
    if ($cal) {
        $txtWhStart.Text = "$($cal.WorkingHoursStartTime)"
        $txtWhEnd.Text   = "$($cal.WorkingHoursEndTime)"
        $cmbWorkDays.Text = "$($cal.WorkDays)"
        $i = $cmbWeekStart.Items.IndexOf("$($cal.WeekStartDay)")
        if ($i -ge 0) { $cmbWeekStart.SelectedIndex = $i }
    }

    # THE TWO WARNINGS THAT MATTER, and they only matter together with
    # ScheduleOnlyDuringWorkHours.
    $msgs = @()
    $rtz = "$(if ($reg) { $reg.TimeZone } else { '' })"
    $wtz = "$(if ($cal) { $cal.WorkingHoursTimeZone } else { '' })"
    if (-not $rtz) {
        $msgs += 'NO REGIONAL TIMEZONE IS SET. Until one is, the room interprets times in the datacentre''s timezone rather than its own.'
    }
    if ($rtz -and $wtz -and $rtz -ne $wtz) {
        $msgs += "THE TWO TIMEZONES DISAGREE: regional is '$rtz', working hours is '$wtz'. Requests will be declined at hours nobody can account for."
    }
    if ($script:Cal -and [bool]$script:Cal.ScheduleOnlyDuringWorkHours) {
        $msgs += 'ScheduleOnlyDuringWorkHours is ON for this room, so everything above is actively deciding what gets declined.'
    }
    $rgWarn.Text = $(if ($msgs.Count) { ($msgs -join "`r`n") } else { 'Regional settings and working hours agree.' })
    $rgWarn.ForeColor = $(if ($msgs.Count -and $msgs[0] -match '^(NO|THE TWO)') { (Get-ToolStatusColour -Kind Bad) } else { (Get-ToolStatusColour -Kind Muted) })
}

$btnRgApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $reg = $script:Reg
        $p = @{ Identity = $script:Room.Smtp; ErrorAction = 'Stop' }
        $changed = @()
        $tz = "$($cmbTz.SelectedItem)"
        if ($tz -and $tz -ne "$(if ($reg) { $reg.TimeZone } else { '' })") { $p['TimeZone'] = $tz; $changed += "TimeZone=$tz" }
        foreach ($f in @(
            @{ N = 'Language';   B = $txtLang;    W = "$(if ($reg) { $reg.Language } else { '' })" }
            @{ N = 'DateFormat'; B = $txtDateFmt; W = "$(if ($reg) { $reg.DateFormat } else { '' })" }
            @{ N = 'TimeFormat'; B = $txtTimeFmt; W = "$(if ($reg) { $reg.TimeFormat } else { '' })" }
        )) {
            $now = "$($f.B.Text)".Trim()
            if ($now -ne "$($f.W)".Trim()) { $p[$f.N] = $now; $changed += "$($f.N)=$now" }
        }
        if ($changed.Count -eq 0) { Set-EditorStatus -Text 'Nothing changed.'; return }
        Invoke-EXOCommand -Name 'Set-MailboxRegionalConfiguration' -Parameters $p | Out-Null
        Write-EditorAudit -Action 'RoomRegional' -Detail ($changed -join '; ')
        Update-RegionalTab
        $script:Loaded['Overview'] = $false
        Set-EditorStatus -Text "Applied: $($changed -join '; ')  A timezone change re-interprets future requests only."
    } 'Apply regional settings'
})

$btnWhApply.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $cal = $script:CalCfg
        $p = @{ Identity = $script:Room.Smtp; ErrorAction = 'Stop' }
        $changed = @()
        foreach ($f in @(
            @{ N = 'WorkingHoursStartTime'; V = "$($txtWhStart.Text)".Trim(); W = "$(if ($cal) { $cal.WorkingHoursStartTime } else { '' })" }
            @{ N = 'WorkingHoursEndTime';   V = "$($txtWhEnd.Text)".Trim();   W = "$(if ($cal) { $cal.WorkingHoursEndTime } else { '' })" }
            @{ N = 'WorkDays';              V = "$($cmbWorkDays.Text)".Trim(); W = "$(if ($cal) { $cal.WorkDays } else { '' })" }
            @{ N = 'WeekStartDay';          V = "$($cmbWeekStart.SelectedItem)"; W = "$(if ($cal) { $cal.WeekStartDay } else { '' })" }
        )) {
            if ($f.V -and $f.V -ne "$($f.W)") { $p[$f.N] = $f.V; $changed += "$($f.N)=$($f.V)" }
        }
        $wtz = "$($cmbWhTz.SelectedItem)"
        if ($wtz -and $wtz -ne "$(if ($cal) { $cal.WorkingHoursTimeZone } else { '' })") {
            $p['WorkingHoursTimeZone'] = $wtz; $changed += "WorkingHoursTimeZone=$wtz"
        }
        if ($changed.Count -eq 0) { Set-EditorStatus -Text 'Nothing changed.'; return }
        Invoke-EXOCommand -Name 'Set-MailboxCalendarConfiguration' -Parameters $p | Out-Null
        Write-EditorAudit -Action 'RoomWorkingHours' -Detail ($changed -join '; ')
        Update-RegionalTab
        $script:Loaded['Overview'] = $false
        Set-EditorStatus -Text "Applied: $($changed -join '; ')"
    } 'Apply working hours'
})
#endregion

#region Room Lists tab ----------------------------------------------------------
<#
    WHICH LISTS THIS ROOM APPEARS IN.

    Room Finder browses ROOM LISTS. A room in no list still books perfectly by
    address and is effectively invisible to anyone looking for a room, which
    is most of "nobody can find the new room".

    MEMBERSHIP IS ASKED PER LIST, and that is not laziness. A room mailbox
    carries no usable 'lists I belong to' property, and reading MemberOfGroup
    on the recipient returns SECURITY group membership, which is a different
    thing entirely.

    A LIST THAT WILL NOT READ IS SHOWN AS '?' AND REFUSES BOTH BUTTONS. Saying
    'not a member' about a list we could not read would invite adding the room
    to it a second time.

    CREATING a room list is a tenant-wide act and is not on this tab -- it
    lives in RoomList.ps1. Set-DistributionGroup has no supported switch that
    turns an existing group into a room list, and groups converted by other
    means are a known source of Room Finder oddities, so there is no convert
    here either.
#>
$rlY = 8
[void](New-TabLabel $tabLists 'ROOM LIST MEMBERSHIP' 8 $rlY -Bold)
$rlY += 22
$rlNote = New-TabLabel $tabLists '' 8 $rlY 1040 -Muted
$rlNote.Size = New-Object System.Drawing.Size(1040, 32)
$rlY += 40
$gridLists = New-TabGrid $tabLists @(@('Member', 90), @('Room list', 320), @('Address', 300), @('Rooms', 80)) 8 $rlY 830 400
$rlBtnY = $rlY
$btnRlAdd    = New-TabButton $tabLists 'Add to list'      850 $rlBtnY 150
$rlBtnY += 34
$btnRlRemove = New-TabButton $tabLists 'Remove from list' 850 $rlBtnY 150
$rlBtnY += 34
$btnRlRefresh = New-TabButton $tabLists 'Refresh'         850 $rlBtnY 150
$rlY += 412
$rlHint = New-TabLabel $tabLists '' 8 $rlY 1040 -Muted
$rlHint.Size = New-Object System.Drawing.Size(1040, 46)
$rlHint.Text = ('Select a list and use the buttons. Room Finder caches membership -- allow up to 24 hours before deciding a change has not worked; this grid is the truth in the meantime.' + "`r`n" +
                'To create a new room list, use the Room Lists tool -- that makes a tenant-wide object and does not belong to any one room.')

function Update-ListsTab {
    $gridLists.Rows.Clear()
    $lists = @()
    try {
        $lists = @(Invoke-EXOCommand -Name 'Get-DistributionGroup' -Parameters @{
                        RecipientTypeDetails = 'RoomList'; ResultSize = 'Unlimited'
                        ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' })
    } catch {
        $script:TabReadFailed = $true
        $rlNote.Text = "Room lists could not be read: $($_.Exception.Message)"
        $rlNote.ForeColor = (Get-ToolStatusColour -Kind Bad)
        return
    }

    $me = "$($script:Room.Smtp)"
    $inAny = 0
    # EVERY LIST IS ASKED FOR ITS MEMBERS, and on a tenant with hundreds of
    # room lists that is hundreds of round trips on the UI thread. There is no
    # cheaper answer -- a room carries no usable 'lists I belong to' property
    # -- so the cost is stated rather than hidden, and the tab loads on demand
    # like every other.
    if (@($lists).Count -gt 60) {
        Set-EditorStatus -Text "Reading membership for $(@($lists).Count) room lists -- this takes a moment."
        $form.Refresh()
    }
    foreach ($l in ($lists | Sort-Object DisplayName)) {
        $member = '?'
        $count  = '?'
        $ok     = $false
        try {
            $mem = @(Invoke-EXOCommand -Name 'Get-DistributionGroupMember' -Parameters @{
                        Identity = "$($l.PrimarySmtpAddress)"; ResultSize = 'Unlimited'
                        ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' })
            $count = "$(@($mem).Count)"
            $isIn = [bool](@($mem | Where-Object { "$($_.PrimarySmtpAddress)" -ieq $me }).Count)
            $member = $(if ($isIn) { 'YES' } else { '' })
            if ($isIn) { $inAny++ }
            $ok = $true
        } catch { }

        $i = $gridLists.Rows.Add($member, "$($l.DisplayName)", "$($l.PrimarySmtpAddress)", $count)
        $gridLists.Rows[$i].Tag = [PSCustomObject]@{
            Address = "$($l.PrimarySmtpAddress)"; Name = "$($l.DisplayName)"
            IsMember = ($member -eq 'YES'); Readable = $ok
        }
        if (-not $ok) {
            $gridLists.Rows[$i].DefaultCellStyle.ForeColor = (Get-ToolStatusColour -Kind Bad)
        } elseif ($member -eq 'YES') {
            $gridLists.Rows[$i].DefaultCellStyle.Font = New-Object System.Drawing.Font('Consolas', 9, [System.Drawing.FontStyle]::Bold)
        }
    }

    if ($inAny -eq 0) {
        $rlNote.Text = "THIS ROOM IS IN NO ROOM LIST. It can be booked by typing its address and it will not appear in Outlook's Room Finder. That is the usual answer to 'nobody can find the new room'."
        $rlNote.ForeColor = (Get-ToolStatusColour -Kind Bad)
    } else {
        $rlNote.Text = "In $inAny of $($gridLists.Rows.Count) room list(s). A list shown in red could not be read -- its membership is unknown and both buttons are refused for it."
        $rlNote.ForeColor = (Get-ToolStatusColour -Kind Muted)
    }
}

function Get-SelectedRoomList {
    if ($gridLists.SelectedRows.Count -eq 0) { return $null }
    return $gridLists.SelectedRows[0].Tag
}

$btnRlRefresh.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction { Update-ListsTab; Set-EditorStatus -Text 'Room lists re-read.' } 'Refresh room lists'
})

$btnRlAdd.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $t = Get-SelectedRoomList
        if (-not $t) { Set-EditorStatus -Text 'Select a room list.' -Bad; return }
        if (-not $t.Readable) { Set-EditorStatus -Text "'$($t.Name)' could not be read, so its membership is unknown. Adding could add the room twice." -Bad; return }
        if ($t.IsMember) { Set-EditorStatus -Text "Already in $($t.Name)."; return }
        $a = Show-ToolDialog(("Add $($script:Room.Display) to the room list '$($t.Name)'?" + "`r`n`r`n" +
                              'It will appear under that list in Outlook''s Room Finder. The cache can take up to 24 hours.'),
                             'Add to room list', 'OKCancel', 'Question')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Invoke-EXOCommand -Name 'Add-DistributionGroupMember' -Parameters @{
            Identity = $t.Address; Member = $script:Room.Smtp; ErrorAction = 'Stop'
        } | Out-Null
        Write-EditorAudit -Action 'RoomListAdd' -Detail "Added to room list $($t.Name) <$($t.Address)>"
        Update-ListsTab
        Set-EditorStatus -Text "Added to $($t.Name). Room Finder can take up to 24 hours to show it."
    } 'Add to a room list'
})

$btnRlRemove.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $t = Get-SelectedRoomList
        if (-not $t) { Set-EditorStatus -Text 'Select a room list.' -Bad; return }
        if (-not $t.Readable) { Set-EditorStatus -Text "'$($t.Name)' could not be read, so its membership is unknown." -Bad; return }
        if (-not $t.IsMember) { Set-EditorStatus -Text "Not in $($t.Name)."; return }
        $a = Show-ToolDialog(("Remove $($script:Room.Display) from the room list '$($t.Name)'?" + "`r`n`r`n" +
                              'It will stop appearing under that list in Room Finder. It remains bookable by address.'),
                             'Remove from room list', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Invoke-EXOCommand -Name 'Remove-DistributionGroupMember' -Parameters @{
            Identity = $t.Address; Member = $script:Room.Smtp; Confirm = $false; ErrorAction = 'Stop'
        } | Out-Null
        Write-EditorAudit -Action 'RoomListRemove' -Detail "Removed from room list $($t.Name) <$($t.Address)>"
        Update-ListsTab
        Set-EditorStatus -Text "Removed from $($t.Name)."
    } 'Remove from a room list'
})
#endregion

#region Calendar tab ------------------------------------------------------------
<#
    WHO CAN SEE AND EDIT THE ROOM'S CALENDAR.

    Folder permissions on the room's own calendar, which is a different
    question from who may BOOK it: a person with Reviewer can see what is
    booked without being allowed to book, and Default governs what everybody
    in the organisation sees.

    THE FOLDER IS FOUND BY TYPE, NEVER BY NAME. Folder names are localised and
    renameable; '<room>:\Calendar' does not resolve against a mailbox created
    in another language. Get-MailboxFolderPath in Common.ps1 does the work.
#>
$cpY = 8
[void](New-TabLabel $tabCalPerm 'CALENDAR PERMISSIONS' 8 $cpY -Bold)
$cpY += 22
$cpNote = New-TabLabel $tabCalPerm '' 8 $cpY 1040 -Muted
$cpNote.Size = New-Object System.Drawing.Size(1040, 32)
$cpNote.Text = ('Default is what everybody in the organisation sees; Anonymous is what people outside it see. Both are shown in amber because changing them changes what EVERYONE sees.' + "`r`n" +
                'AvailabilityOnly shows busy/free with no detail; LimitedDetails adds the subject and location.')
$cpY += 40
$gridCalPerm = New-TabGrid $tabCalPerm @(@('User', 380), @('Rights', 260)) 8 $cpY 700 340
Add-GridPickColumn -Grid $gridCalPerm
$cpBtnY = $cpY
$btnCpGrant  = New-TabButton $tabCalPerm 'Grant...'        720 $cpBtnY 170
$cpBtnY += 34
$btnCpChange = New-TabButton $tabCalPerm 'Change ticked...' 720 $cpBtnY 170
$cpBtnY += 34
$btnCpRemove = New-TabButton $tabCalPerm 'Remove ticked'    720 $cpBtnY 170
$cpBtnY += 34
$btnCpAll    = New-TabButton $tabCalPerm 'Tick all'         720 $cpBtnY 82
$btnCpNone   = New-TabButton $tabCalPerm 'None'             808 $cpBtnY 82
$cpY += 352

function Update-CalPermTab {
    $gridCalPerm.Rows.Clear()
    $r = Get-MailboxFolderAccess -Smtp $script:Room.Smtp -FolderType 'Calendar'
    if (-not $r.Ok) {
        $script:TabReadFailed = $true
        $cpNote.Text = "The calendar permissions could not be read: $($r.Reason)"
        $cpNote.ForeColor = (Get-ToolStatusColour -Kind Bad)
        return
    }
    foreach ($e in @($r.Entries)) {
        if ("$($e.Rights)" -eq 'None') { continue }
        $i = $gridCalPerm.Rows.Add($false, $e.User, $e.Rights)
        $gridCalPerm.Rows[$i].Tag = $e
        if ("$($e.User)" -match '(?i)^(Default|Anonymous)$') {
            $gridCalPerm.Rows[$i].DefaultCellStyle.ForeColor = (Get-ToolStatusColour -Kind Warn)
        }
    }
}

function Show-RoomRoleDialog {
    param([string[]]$Rows, [string]$Title)
    $f = New-Object System.Windows.Forms.Form
    $f.Text          = $Title
    $f.ClientSize    = New-Object System.Drawing.Size(560, 360)
    $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'
    $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $l0 = New-Object System.Windows.Forms.Label
    $l0.Text = "Applies to $(@($Rows).Count) row(s):"
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
    $l1.Text = 'Role:'; $l1.Location = New-Object System.Drawing.Point(348, 34); $l1.Size = New-Object System.Drawing.Size(200, 18)
    $l1.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
    $f.Controls.Add($l1)
    $lst = New-Object System.Windows.Forms.ListBox
    $lst.Location = New-Object System.Drawing.Point(348, 56); $lst.Size = New-Object System.Drawing.Size(200, 238)
    foreach ($r in $script:MailboxFolderRole) { [void]$lst.Items.Add($r) }
    $i = $lst.Items.IndexOf('Reviewer')
    $lst.SelectedIndex = $(if ($i -ge 0) { $i } else { 0 })
    $f.Controls.Add($lst)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Apply'; $ok.Location = New-Object System.Drawing.Point(348, 322); $ok.Size = New-Object System.Drawing.Size(96, 28)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK; $f.Controls.Add($ok)
    $ca = New-Object System.Windows.Forms.Button
    $ca.Text = 'Cancel'; $ca.Location = New-Object System.Drawing.Point(452, 322); $ca.Size = New-Object System.Drawing.Size(96, 28)
    $ca.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $f.Controls.Add($ca)
    $f.AcceptButton = $ok; $f.CancelButton = $ca

    Set-ConsoleTheme -Form $f
    $res = $f.ShowDialog()
    $role = $(if ($res -eq [System.Windows.Forms.DialogResult]::OK) { "$($lst.SelectedItem)" } else { '' })
    $f.Dispose()
    return $role
}

$btnCpAll.Add_Click({  Set-GridPicked -Grid $gridCalPerm -On $true })
$btnCpNone.Add_Click({ Set-GridPicked -Grid $gridCalPerm -On $false })

$btnCpGrant.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $who = Show-InputDialog -Title 'Grant calendar permission' -Width 620 `
                -Prompt 'Who? Addresses or aliases separated by semicolons, or the words Default or Anonymous.'
        $list = ConvertFrom-DialogList -Text "$who"
        if ($list.Count -eq 0) { return }
        $role = Show-RoomRoleDialog -Rows $list -Title 'Grant calendar permission'
        if (-not $role) { return }
        $everyone = @($list | Where-Object { $_ -match '(?i)^(Default|Anonymous)$' })
        $extra = ''
        if ($everyone.Count) { $extra = "`r`n`r`nTHIS INCLUDES $($everyone -join ' AND ') -- that is not one person, it is what EVERYBODY gets on this calendar." }
        $a = Show-ToolDialog(("Give $role on $($script:Room.Display)'s calendar to:" + "`r`n  " + ($list -join "`r`n  ") + $extra),
                             'Grant calendar permission', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $lines = @(); $ok = 0; $bad = 0
        foreach ($u in $list) {
            try { Set-MailboxFolderAccess -Smtp $script:Room.Smtp -FolderType 'Calendar' -User $u -Role $role
                  $lines += "OK      $u -> $role"; $ok++ }
            catch { $lines += "FAILED  $u`: $($_.Exception.Message)"; $bad++ }
        }
        Write-EditorAudit -Action 'RoomCalendarPermission' `
            -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
            -Detail "$role granted to $($list -join '; ')"
        if ($bad) { Show-ResultListDialog -Title "Grant calendar permission -- $ok done, $bad failed" -Lines $lines }
        Update-CalPermTab
        Set-EditorStatus -Text "Granted $ok, failed $bad." -Bad:($bad -gt 0)
    } 'Grant a calendar permission'
})

$btnCpChange.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $picked = Get-GridPicked -Grid $gridCalPerm
        if ($picked.Count -eq 0) { Set-EditorStatus -Text 'Tick one or more rows to change.' -Bad; return }
        $shown = @($picked | ForEach-Object { "$($_.User)  ($($_.Rights))" })
        $role = Show-RoomRoleDialog -Rows $shown -Title 'Change calendar permission'
        if (-not $role) { return }
        $lines = @(); $ok = 0; $bad = 0
        foreach ($e in $picked) {
            try { Set-MailboxFolderAccess -Smtp $script:Room.Smtp -FolderType 'Calendar' -User "$($e.User)" -Role $role
                  $lines += "OK      $($e.User)  $($e.Rights) -> $role"; $ok++ }
            catch { $lines += "FAILED  $($e.User): $($_.Exception.Message)"; $bad++ }
        }
        Write-EditorAudit -Action 'RoomCalendarPermission' `
            -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
            -Detail "set $role on: $(($picked | ForEach-Object { $_.User }) -join '; ')"
        Show-ResultListDialog -Title "Change calendar permission -- $ok done, $bad failed" -Lines $lines
        Update-CalPermTab
        Set-EditorStatus -Text "Changed $ok, failed $bad." -Bad:($bad -gt 0)
    } 'Change a calendar permission'
})

$btnCpRemove.Add_Click({
    if (-not (Test-EditorLoaded)) { return }
    Invoke-EditorAction {
        $picked = Get-GridPicked -Grid $gridCalPerm
        if ($picked.Count -eq 0) { Set-EditorStatus -Text 'Tick one or more rows to remove.' -Bad; return }
        # DEFAULT AND ANONYMOUS CANNOT BE REMOVED, only set to None. Exchange
        # refuses the removal with an error that reads like a permissions
        # problem and is not one.
        $special = @($picked | Where-Object { "$($_.User)" -match '(?i)^(Default|Anonymous)$' })
        $normal  = @($picked | Where-Object { "$($_.User)" -notmatch '(?i)^(Default|Anonymous)$' })
        $extra = $(if ($special.Count) { "`r`n`r`n$($special.Count) of these are Default or Anonymous, which cannot be removed -- they will be set to None instead, which is the same thing in effect and is what EVERYBODY then gets." } else { '' })
        $a = Show-ToolDialog(("Remove $($picked.Count) calendar permission(s) from $($script:Room.Display)?" + "`r`n`r`n  " +
                              (@($picked | ForEach-Object { "$($_.User)  ($($_.Rights))" }) -join "`r`n  ") + $extra),
                             'Remove calendar permission', 'OKCancel', 'Warning')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $lines = @(); $ok = 0; $bad = 0
        foreach ($e in $normal) {
            try { Remove-MailboxFolderAccess -Smtp $script:Room.Smtp -FolderType 'Calendar' -User "$($e.User)"
                  $lines += "OK      removed $($e.User)"; $ok++ }
            catch { $lines += "FAILED  $($e.User): $($_.Exception.Message)"; $bad++ }
        }
        foreach ($e in $special) {
            try { Set-MailboxFolderAccess -Smtp $script:Room.Smtp -FolderType 'Calendar' -User "$($e.User)" -Role 'None'
                  $lines += "OK      set None   $($e.User)"; $ok++ }
            catch { $lines += "FAILED  $($e.User): $($_.Exception.Message)"; $bad++ }
        }
        Write-EditorAudit -Action 'RoomCalendarPermission' `
            -Result $(if ($bad -and $ok) { 'Partial' } elseif ($bad) { 'Failed' } else { 'Success' }) `
            -Detail "$ok removed, $bad failed: $(($picked | ForEach-Object { $_.User }) -join '; ')"
        if ($bad) { Show-ResultListDialog -Title "Remove calendar permission -- $ok done, $bad failed" -Lines $lines }
        Update-CalPermTab
        Set-EditorStatus -Text "Removed $ok, failed $bad." -Bad:($bad -gt 0)
    } 'Remove a calendar permission'
})
#endregion


$cmbMode.Add_SelectedIndexChanged({
    # A SILENT EARLY RETURN WOULD DESYNC THE BOX FROM THE WINDOW, so the
    # selection is put back rather than left showing a mode we are not in.
    $want = $(if ($cmbMode.SelectedIndex -eq 1) { 'OnPrem' } else { 'Cloud' })
    if ($script:Busy) { Sync-EditorModeCombo -Combo $cmbMode -Mode $script:Mode; return }
    if ($want -eq $script:Mode) { return }
    Invoke-EditorAction { Set-EditorMode -Mode $want } 'Switch editing mode'
})

function Invoke-RoomLoad {
    $t = "$($txtRoom.Text)".Trim()
    if (-not $t) { Set-EditorStatus -Text 'Enter a room address.' -Bad; return }
    Set-EditorStatus -Text "Resolving $t..."
    $m = Resolve-MailboxTarget -Smtp $t
    if (-not $m.Ok) {
        $script:Room = $null
        $lblIdent.Text = ''; $lblSync.Text = ''
        Set-EditorStatus -Text "'$t' $($m.Reason)" -Bad
        return
    }
    $script:Room   = $m
    $script:Loaded = @{}
    $script:OnPrem = $null
    # RESOLVED ON LOAD, NOT ON THE SWITCH, so flipping the switch is instant
    # and -- more to the point -- so the switch can be greyed with a reason
    # BEFORE somebody flips it and finds the object is not there.
    if ($script:Mode -eq 'OnPrem' -or $cmbMode.Enabled) {
        Set-EditorStatus -Text "Resolving $t on-premises..."
        try { $script:OnPrem = Resolve-OnPremMailboxTarget -Identity $t } catch { $script:OnPrem = $null }
    }
    # CALENDAR PROCESSING IS READ ONCE, AT LOAD, because four tabs want it and
    # four reads of the same policy is three too many.
    # THE FLAG IS CLEARED BEFORE THE READ THAT CAN SET IT. Clearing it after
    # threw away the one signal that said the policy could not be read, and
    # the Overview tab was then cached as loaded with nothing in it.
    $script:TabReadFailed = $false
    $script:Cal = Get-RoomCalendarProcessing -Smtp $m.Smtp

    Update-EditorIdentity
    Update-OverviewTab
    if (-not $script:TabReadFailed) { $script:Loaded['Overview'] = $true }
    # The tab already on screen has to be loaded too -- the operator may have
    # been on Place when they pressed Load.
    Invoke-EditorTabLoad -Key "$($tabs.SelectedTab.Name)"
    Set-EditorStatus -Text "Loaded $($m.Smtp)."
}


$btnLoad.Add_Click({ Invoke-EditorAction { Invoke-RoomLoad } 'Load room' })
$txtRoom.Add_KeyDown({ if ($_.KeyCode -eq 'Return') { $btnLoad.PerformClick(); $_.SuppressKeyPress = $true } })

# TAB CHANGE AND WINDOW CLOSE ARE WIRED BY THE SHARED RUNTIME, guards included.
Register-EditorTabHandler
#endregion

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Room Editor' | Out-Null

$form.Add_Shown({
    # THE SWITCH IS PROVED BEFORE IT IS OFFERED. A drop-down that lets you pick
    # On-premises and then fails on every Apply is worse than one greyed out
    # with the reason beside it -- and WinForms will not show a tooltip on a
    # disabled control, so the reason goes on the status line.
    # THE REASON GOES ON THE SYNC BANNER, NOT THE STATUS LINE. The status line
    # is overwritten by the load that happens on the very next line, so an
    # explanation put there is gone before anyone can read it -- and WinForms
    # will not show a tooltip on the disabled control either.
    $script:OnPremWhy = Set-EditorModeAvailability -Combo $cmbMode
    if ($script:OnPremWhy) {
        $lblSync.Text = "On-premises editing is unavailable: $($script:OnPremWhy)"
        $lblSync.ForeColor = (Get-ToolStatusColour -Kind Muted)
    }
    if ("$($txtRoom.Text)".Trim()) { $btnLoad.PerformClick() }
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
