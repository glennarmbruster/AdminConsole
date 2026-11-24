#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# ---------------------------------------------------------------------------
# CLOUD BUILD. Runs everything through Invoke-EXOCommand.
#
# Get-CalendarProcessing and Set-CalendarProcessing exist in BOTH Exchanges, so
# the bare names are ambiguous in a process holding both sessions -- exactly the
# collision Invoke-EXOCommand exists to remove. Never call them bare here.
# ---------------------------------------------------------------------------
<#
    CalendarProcessing.ps1
    Resource booking policy viewer/editor -- Layer 2 tool script.

    Shows everything Get-CalendarProcessing returns for a room or equipment
    mailbox and allows editing of the properties Set-CalendarProcessing accepts.

    Author   : Glenn Armbruster
    Revised  : 2026-08-09
#>
# TOOLKIT-HELP-BEGIN
# How a room decides what to do with a meeting request: accept it, decline it,
# hold it for someone to approve, and what the reply says.
#
# This is where "the room accepts everything", "the room declines recurring
# meetings", and "nobody ever approves the bookings" are actually decided.
#
# AutomateProcessing is the master switch and worth understanding before you
# change anything else:
#   AutoAccept    -- the room books itself by the rules below
#   AutoUpdate    -- requests are tentatively marked, a delegate decides
#   None          -- nothing is automatic; the mailbox behaves like a person
# Most of the other settings only take effect under AutoAccept.
#
# BookInPolicy / RequestInPolicy / RequestOutOfPolicy hold lists of people, and
# they interact with the AllBookInPolicy / AllRequestInPolicy switches. Setting
# a list while the matching All- switch is still true does nothing visible.
#
# Editing writes with Set-CalendarProcessing immediately. There is no undo.
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

<#
    THIS TOOL NEEDS A COMMON.PS1 AT LEAST AS NEW AS ITSELF.

    An older Common.ps1 dot-sources perfectly well. It parses, defines what it
    knows, and says nothing -- then this script dies four hundred lines later on
    a function that file has never contained, reporting "the term is not
    recognized" against a line number inside a form.

    That is exactly what happened on first launch: the tool scripts were copied
    into a console still holding the previous Common.ps1. Checking here turns a
    cryptic runtime error into an instruction.
#>
$script:RequiredCommon = @('Get-EXOWritableParameter', 'ConvertTo-CmdletParameterValue')
$missing = @($script:RequiredCommon | Where-Object {
    -not (Get-Command -Name $_ -ErrorAction SilentlyContinue)
})
if ($missing.Count -gt 0) {
    $null = Show-ToolDialog(
        ("This tool needs a newer Common.ps1 than the one in this folder." + "`r`n`r`n" +
         "Missing: $($missing -join ', ')" + "`r`n`r`n" +
         "Copy Common.ps1 from the same release as CalendarProcessing.ps1 into" + "`r`n" +
         "$PSScriptRoot" + "`r`n`r`n" +
         "Config\ does not need to be touched."),
        'Common.ps1 is out of date', 'OK', 'Warning')
    return
}

# NO PROMPT BEFORE THE WINDOW. The address box is ON the form -- see
# Add-ToolTargetBar in Common.ps1. An empty address is fine; the window
# opens and asks for one.
$script:Target = Get-TargetSmtp -Text $SmtpAddress

#region Configuration ----------------------------------------------------------

$script:CalExcluded = @(
    'PSComputerName', 'PSShowComputerName', 'RunspaceId',
    'PropertyNames', 'AddedProperties', 'RemovedProperties',
    'ModifiedProperties', 'PropertyCount', 'ObjectState',
    'IsValid', 'ExchangeVersion', 'OriginatingServer', 'Identity',
    'MailboxOwnerId'
)

<#
    Settings that do nothing unless AutomateProcessing is AutoAccept.

    Not a guess and not cosmetic. An operator who sets AllowConflicts while
    processing is None gets a successful write, no error, and no change in
    behaviour -- then reports that the tool does not work. Marking the row is
    the cheapest possible fix for a whole class of support question.

    The list is conservative: only settings whose entire purpose is deciding
    how to auto-book. Anything ambiguous is left unmarked, because a warning
    that is sometimes wrong is worse than no warning.
#>
$script:CalAutoAcceptOnly = @(
    'AllowConflicts', 'AllowRecurringMeetings', 'BookingWindowInDays',
    'MaximumDurationInMinutes', 'MinimumDurationInMinutes',
    'ConflictPercentageAllowed', 'MaximumConflictInstances',
    'EnforceCapacity', 'EnforceSchedulingHorizon',
    'ScheduleOnlyDuringWorkHours', 'ProcessExternalMeetingMessages'
)

$script:CalWritable = @{}
$script:CalCurrent  = $null

#endregion

#region Data -------------------------------------------------------------------

function Get-CalendarRecord {
    <#
        The booking policy, plus the mailbox type it belongs to.

        The type matters: Get-CalendarProcessing answers for a USER mailbox too,
        returning a policy that is almost entirely inert. Without saying so, the
        grid looks like a room whose settings are simply all off.
    #>
    param([Parameter(Mandatory)][string]$Smtp)

    $out = [PSCustomObject]@{
        Cal      = $null
        Mailbox  = $null
        IsResource = $false
        Error    = ''
    }

    try {
        $out.Mailbox = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop'
        }
        $rtd = "$($out.Mailbox.RecipientTypeDetails)"
        $out.IsResource = ($rtd -match '(?i)room|equipment|workspace')
    } catch {
        $out.Error = "Could not read the mailbox for '$Smtp': $($_.Exception.Message)"
        return $out
    }

    try {
        $out.Cal = Invoke-EXOCommand -Name 'Get-CalendarProcessing' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop'
        }
    } catch {
        $out.Error = "Could not read calendar processing for '$Smtp': $($_.Exception.Message)"
    }
    return $out
}

function Format-CalValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [System.Array]) {
        if ($Value.Count -eq 0) { return '' }
        return (@($Value | ForEach-Object { "$_" }) -join ', ')
    }
    return "$Value"
}

function Get-CalRowState {
    param([string]$Name)
    if (-not $script:CalWritable.ContainsKey($Name)) { return 'Read-only' }
    return 'Editable'
}

function Get-CalRowNote {
    <#
        Why a row might not behave as typed. Returns '' when there is nothing
        worth saying -- a note on every row is a note nobody reads.
    #>
    param([string]$Name, $Record)

    if ($script:CalAutoAcceptOnly -contains $Name) {
        $mode = "$($Record.AutomateProcessing)"
        if ($mode -and $mode -ne 'AutoAccept') {
            return "inactive while AutomateProcessing is $mode"
        }
    }
    if ($Name -eq 'BookInPolicy'      -and $Record.AllBookInPolicy)      { return 'ignored while AllBookInPolicy is True' }
    if ($Name -eq 'RequestInPolicy'   -and $Record.AllRequestInPolicy)   { return 'ignored while AllRequestInPolicy is True' }
    if ($Name -eq 'RequestOutOfPolicy'-and $Record.AllRequestOutOfPolicy){ return 'ignored while AllRequestOutOfPolicy is True' }
    return ''
}

#endregion

#region Write ------------------------------------------------------------------

function Set-CalendarProperty {
    <#
        One property, one call -- same reasoning as RoomPlace. A batch that
        half-applies leaves the grid disagreeing with the tenant, and these
        settings interact, so a rejected combination is a real possibility.
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][string]$Name,
        [string]$Value
    )

    $typed = ConvertTo-CmdletParameterValue -Value $Value -Type $script:CalWritable[$Name]

    $params = @{ Identity = $Smtp; ErrorAction = 'Stop' }
    $params[$Name] = $typed
    Invoke-EXOCommand -Name 'Set-CalendarProcessing' -Parameters $params | Out-Null
}

#endregion

#region UI ---------------------------------------------------------------------

function Show-CalEditDialog {
    param([string]$Name, [string]$Current, [type]$Type, [string]$Note)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Edit $Name"
    $dlg.Size = New-Object System.Drawing.Size(560, 250)
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.StartPosition = 'CenterParent'
    $dlg.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(12, 12)
    $lbl.Size = New-Object System.Drawing.Size(520, 34)
    $dlg.Controls.Add($lbl)

    $under = $null
    if ($Type) { $under = [System.Nullable]::GetUnderlyingType($Type) }
    $tn = if ($under) { $under.Name } elseif ($Type) { $Type.Name } else { 'String' }

    $input = $null
    if ($tn -eq 'Boolean') {
        # Dropdown rather than free text. "false" cast to [bool] is $true, so a
        # typo here would enable the thing it was meant to disable, silently.
        $lbl.Text = "$Name -- yes or no."
        $input = New-Object System.Windows.Forms.ComboBox
        $input.DropDownStyle = 'DropDownList'
        [void]$input.Items.AddRange(@('True', 'False'))
        $input.SelectedItem = $(if ($Current -match '(?i)^true$') { 'True' } else { 'False' })
    } elseif ($under -and $under.IsEnum) {
        # The cmdlet knows its own accepted values. Typing them by hand is how
        # you discover at save time that it is 'AutoAccept', not 'Auto Accept'.
        $lbl.Text = "$Name -- pick a value."
        $input = New-Object System.Windows.Forms.ComboBox
        $input.DropDownStyle = 'DropDownList'
        [void]$input.Items.AddRange([enum]::GetNames($under))
        if ($input.Items.Contains($Current)) { $input.SelectedItem = $Current }
        elseif ($input.Items.Count -gt 0)    { $input.SelectedIndex = 0 }
    } elseif ($Type -and $Type.IsEnum) {
        $lbl.Text = "$Name -- pick a value."
        $input = New-Object System.Windows.Forms.ComboBox
        $input.DropDownStyle = 'DropDownList'
        [void]$input.Items.AddRange([enum]::GetNames($Type))
        if ($input.Items.Contains($Current)) { $input.SelectedItem = $Current }
        elseif ($input.Items.Count -gt 0)    { $input.SelectedIndex = 0 }
    } else {
        $hint = switch ($tn) {
            'Int32' { 'Whole number.' }
            default {
                if ($Type -and $Type.IsArray) { 'Comma-separated list. Blank clears it.' }
                else { 'Blank clears the value.' }
            }
        }
        $lbl.Text = "$Name -- $hint"
        $input = New-Object System.Windows.Forms.TextBox
        $input.Text = $Current
    }
    $input.Location = New-Object System.Drawing.Point(12, 54)
    $input.Size = New-Object System.Drawing.Size(520, 24)
    $dlg.Controls.Add($input)

    $lblWarn = New-Object System.Windows.Forms.Label
    $lblWarn.Location = New-Object System.Drawing.Point(12, 86)
    $lblWarn.Size = New-Object System.Drawing.Size(520, 60)
    $txt = 'Saving writes to Exchange Online immediately. There is no undo.'
    if ($Note) {
        $txt = ("This setting is currently $Note." + "`r`n`r`n" +
                "The value will be saved, but it will not change how the room behaves " +
                "until that is addressed." + "`r`n`r`n" + $txt)
    }
    $lblWarn.Text = $txt
    $dlg.Controls.Add($lblWarn)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Save'; $ok.Location = New-Object System.Drawing.Point(336, 168)
    $ok.Size = New-Object System.Drawing.Size(90, 28)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($ok); $dlg.AcceptButton = $ok

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'; $cancel.Location = New-Object System.Drawing.Point(434, 168)
    $cancel.Size = New-Object System.Drawing.Size(90, 28)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($cancel); $dlg.CancelButton = $cancel

    Set-ConsoleTheme -Form $dlg
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    if ($input -is [System.Windows.Forms.ComboBox]) { return "$($input.SelectedItem)" }
    return "$($input.Text)"
}

function Show-CalForm {
    # NO -Smtp PARAMETER. The address can change while the window is open, so
    # it lives in $script:Target rather than a parameter fixed at build time.

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Calendar Processing -- $script:Target"
    $form.ClientSize = New-Object System.Drawing.Size(884, 632)
    $form.StartPosition = 'CenterScreen'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $lblHead = New-Object System.Windows.Forms.Label
    $lblHead.Location = New-Object System.Drawing.Point(12, 10)
    $lblHead.Size = New-Object System.Drawing.Size(860, 34)
    $form.Controls.Add($lblHead)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(12, 50)
    $grid.Size = New-Object System.Drawing.Size(860, 490)
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.ReadOnly = $true
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $false
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.RowTemplate.Height = 24
    foreach ($c in @(
        @{ N = 'Setting'; W = 26 }, @{ N = 'Value'; W = 34 },
        @{ N = 'State'; W = 12 },   @{ N = 'Note'; W = 28 })) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c.N; $col.HeaderText = $c.N; $col.FillWeight = $c.W
        $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
        [void]$grid.Columns.Add($col)
    }
    $form.Controls.Add($grid)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(12, 548)
    $lblStatus.Size = New-Object System.Drawing.Size(860, 34)
    $form.Controls.Add($lblStatus)

    $btnEdit = New-Object System.Windows.Forms.Button
    $btnEdit.Text = 'Edit'; $btnEdit.Location = New-Object System.Drawing.Point(12, 590)
    $btnEdit.Size = New-Object System.Drawing.Size(110, 30)
    $form.Controls.Add($btnEdit)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = 'Refresh'; $btnRefresh.Location = New-Object System.Drawing.Point(130, 590)
    $btnRefresh.Size = New-Object System.Drawing.Size(110, 30)
    $form.Controls.Add($btnRefresh)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Location = New-Object System.Drawing.Point(762, 590)
    $btnClose.Size = New-Object System.Drawing.Size(110, 30)
    $form.Controls.Add($btnClose)
    $btnClose.Add_Click({ $form.Close() })

    $load = {
        # OPENED WITH NO ADDRESS IS A NORMAL STATE NOW -- the window appears
        # before anything has been typed, so the refresh must not run empty.
        if (-not "$script:Target".Trim()) {
            $form.Text = 'Calendar Processing'
            return
        }
        $form.Text = 'Calendar Processing: ' + $script:Target
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $grid.Rows.Clear()
            $rec = Get-CalendarRecord -Smtp $script:Target
            $script:CalCurrent = $rec.Cal

            if ($rec.Error) { $lblStatus.Text = $rec.Error; return }
            if (-not $rec.Cal) { $lblStatus.Text = "No calendar processing returned for '$script:Target'."; return }

            $head = "$($rec.Mailbox.DisplayName)  --  $($rec.Mailbox.RecipientTypeDetails)"
            $head += "  --  AutomateProcessing: $($rec.Cal.AutomateProcessing)"
            if (-not $rec.IsResource) {
                $head += ("`r`nThis is not a room or equipment mailbox. These settings exist on it " +
                          "but do almost nothing on a user mailbox.")
            }
            $lblHead.Text = $head

            $props = @($rec.Cal.PSObject.Properties |
                       Where-Object { $script:CalExcluded -notcontains $_.Name } |
                       Sort-Object Name)
            $noted = 0
            foreach ($p in $props) {
                $state = Get-CalRowState -Name $p.Name
                $note  = Get-CalRowNote -Name $p.Name -Record $rec.Cal
                if ($note) { $noted++ }
                $i = $grid.Rows.Add($p.Name, (Format-CalValue $p.Value), $state, $note)
                if ($state -ne 'Editable') {
                    $grid.Rows[$i].Cells['State'].Style.ForeColor = [System.Drawing.Color]::DimGray
                }
                if ($note) {
                    $grid.Rows[$i].Cells['Note'].Style.ForeColor = [System.Drawing.Color]::DarkOrange
                }
            }

            $ed = @($props | Where-Object { (Get-CalRowState -Name $_.Name) -eq 'Editable' }).Count
            $msg = "$($props.Count) settings. $ed editable."
            if ($noted -gt 0) {
                $msg += "  $noted currently have no effect -- see the Note column."
            }
            $lblStatus.Text = $msg
        } catch {
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Reading calendar processing'
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $doEdit = {
        if ($grid.SelectedRows.Count -eq 0) { return }
        $row   = $grid.SelectedRows[0]
        $name  = "$($row.Cells['Setting'].Value)"
        $state = "$($row.Cells['State'].Value)"
        $note  = "$($row.Cells['Note'].Value)"

        if ($state -ne 'Editable') {
            $null = Show-ToolDialog(
                "'$name' is returned by Get-CalendarProcessing but Set-CalendarProcessing has no parameter for it, so it cannot be changed from here.",
                'Read-only', 'OK', 'Information')
            return
        }

        $current = "$($row.Cells['Value'].Value)"
        $new = Show-CalEditDialog -Name $name -Current $current `
                                  -Type $script:CalWritable[$name] -Note $note
        if ($null -eq $new) { return }
        if ($new -eq $current) { $lblStatus.Text = "$name unchanged."; return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Set-CalendarProperty -Smtp $script:Target -Name $name -Value $new
            Write-CompletedLog -Action 'Configure' -Target $script:Target -Result 'Success' `
                -Detail "Set-CalendarProcessing $name : '$current' -> '$new'"
            $lblStatus.Text = "$name updated."
            & $load
        } catch {
            Write-FailedLog -Action 'Configure' -Target $script:Target -ErrorRecord $_ `
                -Detail "Set-CalendarProcessing $name : '$current' -> '$new'"
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context "Setting $name"
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $btnEdit.Add_Click($doEdit)
    $grid.Add_CellDoubleClick({ if ($_.RowIndex -ge 0) { & $doEdit } })
    $btnRefresh.Add_Click($load)
    # BEFORE Add-ToolHelpButton: the bar shifts every existing control down
    # and grows the form, and the help button places itself from ClientSize at
    # run time.
    $bar = Add-ToolTargetBar -Form $form -Label 'Room address:' -InitialValue $script:Target
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        & $load
    })


    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Calendar Processing' | Out-Null
    Set-ConsoleTheme -Form $form
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    & $load
    [void]$form.ShowDialog()
}

#endregion

$script:CalWritable = Get-EXOWritableParameter -Name 'Set-CalendarProcessing'
if ($script:CalWritable.Count -eq 0) {
    $null = Show-ToolDialog(
        ("Set-CalendarProcessing could not be found in this session." + "`r`n`r`n" +
         "Connect to Exchange Online first."),
        'Not connected', 'OK', 'Warning')
    return
}

Show-CalForm
