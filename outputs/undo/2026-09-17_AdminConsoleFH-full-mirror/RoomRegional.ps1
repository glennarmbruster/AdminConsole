#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# ---------------------------------------------------------------------------
# CLOUD BUILD. Everything goes through Invoke-EXOCommand.
#
# The *-MailboxRegionalConfiguration and *-MailboxCalendarConfiguration cmdlets
# exist in BOTH Exchanges, so the bare names are ambiguous in a process holding
# both sessions.
# ---------------------------------------------------------------------------
<#
    RoomRegional.ps1
    Room timezone, language and working hours -- Layer 2 tool script.

    Author   : Glenn Armbruster
    Revised  : 2026-08-09
#>
# TOOLKIT-HELP-BEGIN
# The timezone a room thinks it is in, and the hours it considers bookable.
#
# A room created without a timezone inherits one, and that is how a room ends up
# accepting a 9am booking that shows as 3am to the people in it. Nothing errors;
# the meeting is simply at the wrong time.
#
# TWO TIMEZONES LIVE HERE and they are not the same setting:
#   Regional TimeZone          -- how the mailbox stores and displays times
#   Working hours TimeZone     -- the zone the working hours below are stated in
# They should almost always match. When they do not, "only during working hours"
# means something nobody intended, and this tool says so.
#
# ScheduleOnlyDuringWorkHours in Calendar Processing depends entirely on the
# working hours set here. Turning that on without checking these is how a room
# starts declining bookings for no visible reason.
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

$script:RequiredCommon = @('Get-EXOWritableParameter', 'ConvertTo-CmdletParameterValue')
$missing = @($script:RequiredCommon | Where-Object {
    -not (Get-Command -Name $_ -ErrorAction SilentlyContinue)
})
if ($missing.Count -gt 0) {
    $null = Show-ToolDialog(
        ("This tool needs a newer Common.ps1 than the one in this folder." + "`r`n`r`n" +
         "Missing: $($missing -join ', ')" + "`r`n`r`n" +
         "Copy Common.ps1 from the same release into" + "`r`n$PSScriptRoot"),
        'Common.ps1 is out of date', 'OK', 'Warning')
    return
}

# NO PROMPT BEFORE THE WINDOW. The address box is ON the form -- see
# Add-ToolTargetBar in Common.ps1. An empty address is fine; the window
# opens and asks for one.
$script:Target = Get-TargetSmtp -Text $SmtpAddress

#region Configuration ----------------------------------------------------------

<#
    Two cmdlet pairs, one grid.

    Regional and Calendar configuration are separate cmdlets that between them
    decide when a room is bookable, and an operator chasing a wrong-time booking
    should not have to know which half owns which field. The grid records which
    cmdlet each row came from so the write goes back to the right one.
#>
$script:RegionalSource = 'Regional'
$script:CalendarSource = 'Calendar'

$script:RrExcluded = @(
    'PSComputerName', 'PSShowComputerName', 'RunspaceId',
    'PropertyNames', 'AddedProperties', 'RemovedProperties',
    'ModifiedProperties', 'PropertyCount', 'ObjectState',
    'IsValid', 'ExchangeVersion', 'OriginatingServer', 'Identity',
    'MailboxOwnerId'
)

# The fields worth showing. Get-MailboxCalendarConfiguration returns forty-odd
# properties, most of them about the OWNER's Outlook experience -- reminders,
# weather, birthday calendars -- which mean nothing on a room. Listing them all
# would bury the four that matter.
# ORDER IS DELIBERATE, not alphabetical. Start before End, because the pair is
# read as a range and seeing the end of it first is a small constant tax on
# everyone who opens this window.
$script:CalendarShow = @(
    'WorkingHoursStartTime', 'WorkingHoursEndTime', 'WorkingHoursTimeZone',
    'WorkDays', 'WeekStartDay'
)

$script:RrWritable = @{}
$script:RrCurrent  = $null

#endregion

#region Data -------------------------------------------------------------------

function Get-RoomRegional {
    param([Parameter(Mandatory)][string]$Smtp)

    $out = [PSCustomObject]@{
        Mailbox = $null; Regional = $null; Calendar = $null
        IsResource = $false; Error = ''
    }
    try {
        $out.Mailbox = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop'
        }
        $out.IsResource = ("$($out.Mailbox.RecipientTypeDetails)" -match '(?i)room|equipment|workspace')
    } catch {
        $out.Error = "Could not read the mailbox for '$Smtp': $($_.Exception.Message)"
        return $out
    }
    try {
        $out.Regional = Invoke-EXOCommand -Name 'Get-MailboxRegionalConfiguration' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop'
        }
    } catch {
        $out.Error = "Could not read regional configuration: $($_.Exception.Message)"
    }
    try {
        $out.Calendar = Invoke-EXOCommand -Name 'Get-MailboxCalendarConfiguration' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop'
        }
    } catch {
        if (-not $out.Error) {
            $out.Error = "Could not read calendar configuration: $($_.Exception.Message)"
        }
    }
    return $out
}

function Format-RrValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [System.Array]) { return (@($Value | ForEach-Object { "$_" }) -join ', ') }
    return "$Value"
}

#endregion

#region Write ------------------------------------------------------------------

function Set-RoomRegionalProperty {
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Source,
        [string]$Value
    )

    $cmdlet = $(if ($Source -eq $script:CalendarSource) {
        'Set-MailboxCalendarConfiguration'
    } else {
        'Set-MailboxRegionalConfiguration'
    })

    $key = "$Source`:$Name"
    $typed = ConvertTo-CmdletParameterValue -Value $Value -Type $script:RrWritable[$key]

    $params = @{ Identity = $Smtp; ErrorAction = 'Stop' }
    $params[$Name] = $typed
    Invoke-EXOCommand -Name $cmdlet -Parameters $params | Out-Null
}

#endregion

#region UI ---------------------------------------------------------------------

function Show-RrEditDialog {
    param([string]$Name, [string]$Current, [type]$Type)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Edit $Name"
    $dlg.Size = New-Object System.Drawing.Size(560, 230)
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false
    $dlg.StartPosition = 'CenterParent'
    $dlg.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(12, 12)
    $lbl.Size = New-Object System.Drawing.Size(520, 34)
    $dlg.Controls.Add($lbl)

    $under = $null
    if ($Type) { $under = [System.Nullable]::GetUnderlyingType($Type) }
    $eff = $(if ($under) { $under } else { $Type })
    $tn  = $(if ($eff) { $eff.Name } else { 'String' })

    $input = $null
    if ($Name -match 'TimeZone') {
        <#
            Timezone IDs come from Windows, not from a list typed in here.

            Exchange accepts Windows timezone IDs -- 'Central Standard Time',
            not 'CST' and not an IANA name like 'America/Chicago'. Those are the
            two things people reach for, and both are rejected with a message
            about an invalid value rather than about the format.

            GetSystemTimeZones is the same source Exchange validates against, so
            a dropdown built from it cannot offer something invalid.
        #>
        $lbl.Text = "$Name -- Windows timezone ID."
        $input = New-Object System.Windows.Forms.ComboBox
        $input.DropDownStyle = 'DropDownList'
        $ids = @([System.TimeZoneInfo]::GetSystemTimeZones() | ForEach-Object { $_.Id } | Sort-Object)
        [void]$input.Items.AddRange($ids)
        if ($input.Items.Contains($Current)) { $input.SelectedItem = $Current }
    } elseif ($tn -eq 'Boolean') {
        $lbl.Text = "$Name -- yes or no."
        $input = New-Object System.Windows.Forms.ComboBox
        $input.DropDownStyle = 'DropDownList'
        [void]$input.Items.AddRange(@('True', 'False'))
        $input.SelectedItem = $(if ($Current -match '(?i)^true$') { 'True' } else { 'False' })
    } elseif ($eff -and $eff.IsEnum) {
        # WorkDays and WeekStartDay are enums, and WorkDays is a FLAGS enum --
        # 'Weekday', 'AllDays', or a comma-separated set. The names come from
        # the type so they cannot be misspelt, and a free-text box is offered
        # alongside for the combined values a single-select list cannot express.
        $lbl.Text = "$Name -- pick one, or type a comma-separated set."
        $input = New-Object System.Windows.Forms.ComboBox
        $input.DropDownStyle = 'DropDown'
        [void]$input.Items.AddRange([enum]::GetNames($eff))
        $input.Text = $Current
    } else {
        $hint = switch ($tn) {
            'TimeSpan' { 'Time of day, for example 08:00:00.' }
            'Int32'    { 'Whole number.' }
            default    { 'Blank clears the value.' }
        }
        $lbl.Text = "$Name -- $hint"
        $input = New-Object System.Windows.Forms.TextBox
        $input.Text = $Current
    }
    $input.Location = New-Object System.Drawing.Point(12, 54)
    $input.Size = New-Object System.Drawing.Size(520, 24)
    $dlg.Controls.Add($input)

    $lblWarn = New-Object System.Windows.Forms.Label
    $lblWarn.Location = New-Object System.Drawing.Point(12, 88)
    $lblWarn.Size = New-Object System.Drawing.Size(520, 50)
    $lblWarn.Text = ('Saving writes to Exchange Online immediately. Existing bookings are not ' +
                     'moved -- a timezone change re-interprets future requests only.')
    $dlg.Controls.Add($lblWarn)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Save'; $ok.Location = New-Object System.Drawing.Point(336, 150)
    $ok.Size = New-Object System.Drawing.Size(90, 28)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($ok); $dlg.AcceptButton = $ok

    $ca = New-Object System.Windows.Forms.Button
    $ca.Text = 'Cancel'; $ca.Location = New-Object System.Drawing.Point(434, 150)
    $ca.Size = New-Object System.Drawing.Size(90, 28)
    $ca.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($ca); $dlg.CancelButton = $ca

    Set-ConsoleTheme -Form $dlg
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    if ($input -is [System.Windows.Forms.ComboBox]) {
        if ($input.DropDownStyle -eq 'DropDownList') { return "$($input.SelectedItem)" }
        return "$($input.Text)"
    }
    return "$($input.Text)"
}

function Show-RoomRegionalForm {
    # NO -Smtp PARAMETER. The address can change while the window is open, so
    # it lives in $script:Target rather than a parameter fixed at build time.

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Room Regional -- $script:Target"
    $form.Size = New-Object System.Drawing.Size(880, 620)
    $form.StartPosition = 'CenterScreen'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $lblHead = New-Object System.Windows.Forms.Label
    $lblHead.Location = New-Object System.Drawing.Point(12, 10)
    $lblHead.Size = New-Object System.Drawing.Size(840, 50)
    $form.Controls.Add($lblHead)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(12, 66)
    $grid.Size = New-Object System.Drawing.Size(840, 330)
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.ReadOnly = $true
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $false
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.RowTemplate.Height = 24
    foreach ($c in @(
        @{ N = 'Setting'; W = 30 }, @{ N = 'Value'; W = 34 },
        @{ N = 'From'; W = 14 },    @{ N = 'State'; W = 22 })) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c.N; $col.HeaderText = $c.N; $col.FillWeight = $c.W
        $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
        [void]$grid.Columns.Add($col)
    }
    $form.Controls.Add($grid)

    $btnEdit = New-Object System.Windows.Forms.Button
    $btnEdit.Text = 'Edit'; $btnEdit.Location = New-Object System.Drawing.Point(12, 406)
    $btnEdit.Size = New-Object System.Drawing.Size(120, 30)
    $form.Controls.Add($btnEdit)

    $btnReload = New-Object System.Windows.Forms.Button
    $btnReload.Text = 'Reload'; $btnReload.Location = New-Object System.Drawing.Point(140, 406)
    $btnReload.Size = New-Object System.Drawing.Size(120, 30)
    $form.Controls.Add($btnReload)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Location = New-Object System.Drawing.Point(732, 406)
    $btnClose.Size = New-Object System.Drawing.Size(120, 30)
    $form.Controls.Add($btnClose)
    $btnClose.Add_Click({ $form.Close() })

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Location = New-Object System.Drawing.Point(12, 446)
    $txtOut.Size = New-Object System.Drawing.Size(840, 120)
    $txtOut.Multiline = $true
    $txtOut.ScrollBars = 'Vertical'
    $txtOut.ReadOnly = $true
    $txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtOut)

    $addRows = {
        param($Obj, [string]$Source, $Only)
        if (-not $Obj) { return }

        # WHEN A CURATED LIST IS GIVEN, ITS ORDER IS THE POINT.
        #
        # This sorted everything alphabetically, which put WorkingHoursEndTime
        # above WorkingHoursStartTime -- end before start, in a pair that is
        # read as a range. $script:CalendarShow was already written in the
        # sensible order and the sort threw that away.
        #
        # Alphabetical is a reasonable default for a list nobody curated. It is
        # the wrong answer for one somebody did.
        if ($Only) {
            $byName = @{}
            foreach ($pp in $Obj.PSObject.Properties) { $byName[$pp.Name] = $pp }
            $props = @()
            foreach ($n in $Only) {
                if ($byName.ContainsKey($n) -and ($script:RrExcluded -notcontains $n)) {
                    $props += $byName[$n]
                }
            }
        } else {
            # Plain alphabetical, and deliberately left that way.
            #
            # A "pull TimeZone to the top" reorder was added here and dropped
            # TimeZone from the grid entirely. It was never asked for -- it was
            # tidiness bolted onto a requested fix -- and it cost a row that
            # matters more than any ordering. Reverted rather than debugged: an
            # unrequested improvement that loses data has negative value even
            # once it works.
            #
            # The Calendar block above orders by its curated list, which IS what
            # was asked for and touches only which order five known names come
            # in. Nothing there can go missing, because that branch only ever
            # emits names it was given.
            $props = @($Obj.PSObject.Properties |
                       Where-Object { $script:RrExcluded -notcontains $_.Name } |
                       Sort-Object Name)
        }

        foreach ($p in $props) {
            $key   = "$Source`:$($p.Name)"
            $state = $(if ($script:RrWritable.ContainsKey($key)) { 'Editable' } else { 'Read-only' })
            $i = $grid.Rows.Add($p.Name, (Format-RrValue $p.Value), $Source, $state)
            $grid.Rows[$i].Tag = @{ Name = $p.Name; Source = $Source; Key = $key }
            if ($state -ne 'Editable') {
                $grid.Rows[$i].Cells['State'].Style.ForeColor = [System.Drawing.Color]::DimGray
            }
        }
    }

    $load = {
        # OPENED WITH NO ADDRESS IS A NORMAL STATE NOW -- the window appears
        # before anything has been typed, so the refresh must not run empty.
        if (-not "$script:Target".Trim()) {
            $form.Text = 'Room Regional'
            return
        }
        $form.Text = 'Room Regional: ' + $script:Target
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $grid.Rows.Clear()
            $rec = Get-RoomRegional -Smtp $script:Target
            $script:RrCurrent = $rec
            if ($rec.Error -and -not $rec.Regional -and -not $rec.Calendar) {
                $txtOut.Text = $rec.Error; return
            }

            & $addRows $rec.Regional $script:RegionalSource $null
            & $addRows $rec.Calendar $script:CalendarSource $script:CalendarShow

            $tzReg = "$($rec.Regional.TimeZone)"
            $tzWrk = "$($rec.Calendar.WorkingHoursTimeZone)"
            $head  = "$($rec.Mailbox.DisplayName)  --  $($rec.Mailbox.RecipientTypeDetails)"
            $head += "`r`nTimezone: $(if ($tzReg) { $tzReg } else { '(not set)' })"
            if (-not $rec.IsResource) {
                $head += "`r`nThis is not a room or equipment mailbox."
            }
            $lblHead.Text = $head

            $sb = New-Object System.Text.StringBuilder
            if (-not $tzReg) {
                [void]$sb.AppendLine('  WARNING: no regional timezone is set on this mailbox.')
                [void]$sb.AppendLine('  It will fall back to a default, which is how a room accepts a')
                [void]$sb.AppendLine('  booking that shows at the wrong hour to everyone in it.')
                [void]$sb.AppendLine('')
            }
            if ($tzReg -and $tzWrk -and ($tzReg -ne $tzWrk)) {
                [void]$sb.AppendLine('  WARNING: the two timezones disagree.')
                [void]$sb.AppendLine("    Regional      : $tzReg")
                [void]$sb.AppendLine("    Working hours : $tzWrk")
                [void]$sb.AppendLine('  Working hours are interpreted in the SECOND one, so')
                [void]$sb.AppendLine('  ScheduleOnlyDuringWorkHours will not mean what it looks like.')
                [void]$sb.AppendLine('')
            }
            [void]$sb.AppendLine('  Regional rows come from Get-MailboxRegionalConfiguration.')
            [void]$sb.AppendLine('  Calendar rows come from Get-MailboxCalendarConfiguration and are')
            [void]$sb.AppendLine('  filtered to the ones that affect a room. Double-click a row to edit.')
            $txtOut.Text = $sb.ToString()
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Reading regional configuration'
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $doEdit = {
        if ($grid.SelectedRows.Count -eq 0) { return }
        $row = $grid.SelectedRows[0]
        $tag = $row.Tag
        if (-not $tag) { return }
        if ("$($row.Cells['State'].Value)" -ne 'Editable') {
            $null = Show-ToolDialog(
                "'$($tag.Name)' is returned by the Get- cmdlet but the matching Set- cmdlet has no parameter for it.",
                'Read-only', 'OK', 'Information')
            return
        }

        $current = "$($row.Cells['Value'].Value)"
        $new = Show-RrEditDialog -Name $tag.Name -Current $current -Type $script:RrWritable[$tag.Key]
        if ($null -eq $new) { return }
        if ($new -eq $current) { $txtOut.Text = "$($tag.Name) unchanged."; return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Set-RoomRegionalProperty -Smtp $script:Target -Name $tag.Name -Source $tag.Source -Value $new
            Write-CompletedLog -Action 'Configure' -Target $script:Target -Result 'Success' `
                -Detail "$($tag.Source) $($tag.Name) : '$current' -> '$new'"
            $txtOut.Text = "$($tag.Name) updated."
            & $load
        } catch {
            Write-FailedLog -Action 'Configure' -Target $script:Target -ErrorRecord $_ `
                -Detail "$($tag.Source) $($tag.Name) : '$current' -> '$new'"
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context "Setting $($tag.Name)"
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $btnEdit.Add_Click($doEdit)
    $grid.Add_CellDoubleClick({ if ($_.RowIndex -ge 0) { & $doEdit } })
    $btnReload.Add_Click($load)
    # BEFORE Add-ToolHelpButton: the bar shifts every existing control down
    # and grows the form, and the help button places itself from ClientSize at
    # run time.
    $bar = Add-ToolTargetBar -Form $form -Label 'Room address:' -InitialValue $script:Target
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        & $load
    })


    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Room Regional' | Out-Null
    Set-ConsoleTheme -Form $form
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    & $load
    [void]$form.ShowDialog()
}

#endregion

# Writable sets from BOTH cmdlets, keyed by source so a name appearing in each
# cannot be written back to the wrong one. TimeZone exists on the regional
# cmdlet and WorkingHoursTimeZone on the calendar one -- close enough in a flat
# hashtable to be worth keeping apart deliberately.
foreach ($pair in @(
    @{ S = $script:RegionalSource; C = 'Set-MailboxRegionalConfiguration' },
    @{ S = $script:CalendarSource; C = 'Set-MailboxCalendarConfiguration' })) {
    $w = Get-EXOWritableParameter -Name $pair.C
    foreach ($k in $w.Keys) { $script:RrWritable["$($pair.S)`:$k"] = $w[$k] }
}

if ($script:RrWritable.Count -eq 0) {
    $null = Show-ToolDialog(
        ("The regional configuration cmdlets could not be found in this session." + "`r`n`r`n" +
         "Connect to Exchange Online first."),
        'Not connected', 'OK', 'Warning')
    return
}

Show-RoomRegionalForm
