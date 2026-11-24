#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# ---------------------------------------------------------------------------
# CLOUD ONLY, AND NOT BY CHOICE. Get-Place and Set-Place exist only in Exchange
# Online -- there is no on-premises equivalent and no AD- twin of this script.
# ---------------------------------------------------------------------------
<#
    RoomPlace.ps1
    Room place metadata viewer/editor -- Layer 2 tool script.

    Shows everything Get-Place returns for a room mailbox and allows editing of
    the properties Set-Place actually accepts.

    Author   : Glenn Armbruster
    Revised  : 2026-08-09
#>
# TOOLKIT-HELP-BEGIN
# The room metadata behind Outlook's room finder -- capacity, floor, building,
# devices, accessibility, tags.
#
# This is NOT the same as the mailbox. A room has a mailbox (name, address,
# permissions) and a "place" record (what the room is like), and only the place
# record drives room suggestions. A room that never appears in a search usually
# has an empty place record rather than a broken mailbox.
#
# Editing writes with Set-Place immediately. There is no staging and no undo.
#
# IN A HYBRID ORGANISATION some fields cannot be set here at all. If the room
# mailbox originated on-premises and syncs up, Microsoft blocks Set-Place from
# changing City, CountryOrRegion, GeoCoordinates, Phone, PostalCode, State and
# Street -- those come from the on-premises directory and must be changed there
# with Set-User or Set-Mailbox. This tool marks those rows rather than letting
# you discover it from a failure.
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

    Dot-sourcing an older Common.ps1 succeeds. It parses, it defines everything
    it knows about, and nothing complains -- the tool then runs for four hundred
    lines and dies on the first call to a function that file has never heard of,
    reporting "the term is not recognized" against a line number in the middle
    of a form.

    That happened on the first launch of this script: the two new tool files
    were copied into a console whose Common.ps1 was still the previous one. The
    message named the symptom and not one useful thing about the cause.

    A copied toolkit is the normal way this ships, so a tool script arriving
    ahead of its library is a NORMAL failure, not an exotic one. It deserves a
    sentence that says what to do.
#>
$script:RequiredCommon = @('Get-EXOWritableParameter', 'ConvertTo-CmdletParameterValue')
$missing = @($script:RequiredCommon | Where-Object {
    -not (Get-Command -Name $_ -ErrorAction SilentlyContinue)
})
if ($missing.Count -gt 0) {
    $null = Show-ToolDialog(
        ("This tool needs a newer Common.ps1 than the one in this folder." + "`r`n`r`n" +
         "Missing: $($missing -join ', ')" + "`r`n`r`n" +
         "Copy Common.ps1 from the same release as RoomPlace.ps1 into" + "`r`n" +
         "$PSScriptRoot" + "`r`n`r`n" +
         "Config\ does not need to be touched."),
        'Common.ps1 is out of date', 'OK', 'Warning')
    return
}

# NO PROMPT BEFORE THE WINDOW. The address box is ON the form -- see
# Add-ToolTargetBar in Common.ps1. An empty address is fine; the window opens
# and asks for one.
$script:Target = Get-TargetSmtp -Text $SmtpAddress

#region Configuration ----------------------------------------------------------

<#
    Blocked on synchronised rooms.

    Straight from the Set-Place reference: "In hybrid environments, this cmdlet
    doesn't work on the following properties on synchronized room mailboxes:
    City, CountryOrRegion, GeoCoordinates, Phone, PostalCode, State, or Street."

    This tenant is hybrid and most rooms originate on-premises, so for this tenant
    that is the majority of the address block. The cmdlet does not always error
    on these -- it can accept the call and simply not persist the value, which
    is the worst possible outcome: the operator sees success, reopens the tool
    later, and finds the old value back with no explanation.

    So they are shown, marked, and refused BEFORE the call rather than after.
#>
$script:PlaceHybridBlocked = @(
    'City', 'CountryOrRegion', 'GeoCoordinates', 'Phone',
    'PostalCode', 'State', 'Street'
)

# Noise from the remoting layer, not properties of the room.
$script:PlaceExcluded = @(
    'PSComputerName', 'PSShowComputerName', 'RunspaceId',
    'PropertyNames', 'AddedProperties', 'RemovedProperties',
    'ModifiedProperties', 'PropertyCount', 'ObjectState',
    'IsValid', 'ExchangeVersion', 'OriginatingServer', 'Identity'
)

$script:PlaceRoom      = $null
$script:PlaceWritable  = @{}
$script:PlaceIsSynced  = $false

#endregion

#region Data -------------------------------------------------------------------

function Get-PlaceRecord {
    <#
        Get-Place for the metadata, Get-Mailbox for two facts the place record
        does not carry: whether this is actually a room, and whether it came
        from on-premises.

        RecipientTypeDetails is checked because Get-Place accepts any mailbox
        and returns a near-empty record for a user -- which reads as "a room
        with nothing filled in" rather than "that is not a room".
    #>
    param([Parameter(Mandatory)][string]$Smtp)

    $out = [PSCustomObject]@{
        Place    = $null
        Mailbox  = $null
        IsRoom   = $false
        IsSynced = $false
        Error    = ''
    }

    try {
        $out.Mailbox = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop'
        }
    } catch {
        $out.Error = "Could not read the mailbox for '$Smtp': $($_.Exception.Message)"
        return $out
    }

    $rtd = "$($out.Mailbox.RecipientTypeDetails)"
    $out.IsRoom = ($rtd -match '(?i)room|workspace')

    # IsDirSynced is the on-premises origin flag. A room created in the cloud
    # is editable in full; a synced one is not, and the difference decides
    # which rows this tool will let you touch.
    if ($out.Mailbox.PSObject.Properties['IsDirSynced']) {
        $out.IsSynced = [bool]$out.Mailbox.IsDirSynced
    }

    try {
        $out.Place = Invoke-EXOCommand -Name 'Get-Place' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop'
        }
    } catch {
        $out.Error = "Could not read the place record for '$Smtp': $($_.Exception.Message)"
    }
    return $out
}

function Format-PlaceValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [System.Array]) {
        return (@($Value | ForEach-Object { "$_" }) -join ', ')
    }
    return "$Value"
}

function Get-PlaceRowState {
    <#
        Three states, and the operator is told which and why:

          Editable   -- Set-Place takes it and this room is not synced for it
          On-prem    -- Set-Place takes it but Microsoft ignores it here
          Read-only  -- Set-Place has no parameter for it at all

        Read-only is the honest answer for anything Get-Place returns that
        Set-Place cannot accept. Showing those rows as editable and failing on
        save teaches people not to trust the tool.
    #>
    param([string]$Name)

    if (-not $script:PlaceWritable.ContainsKey($Name)) { return 'Read-only' }
    if ($script:PlaceIsSynced -and ($script:PlaceHybridBlocked -contains $Name)) { return 'On-prem' }
    return 'Editable'
}

#endregion

#region Write ------------------------------------------------------------------

function Set-PlaceProperty {
    <#
        One property, one call.

        Deliberately not batched. Set-Place is documented as failing on certain
        COMBINATIONS of properties in older module versions, and a batch that
        half-applies leaves the grid disagreeing with the tenant. One call per
        edit is slower and always truthful about what landed.
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [Parameter(Mandatory)][string]$Name,
        [string]$Value
    )

    $typed = ConvertTo-CmdletParameterValue -Value $Value -Type $script:PlaceWritable[$Name]

    $params = @{ Identity = $Smtp; ErrorAction = 'Stop' }
    $params[$Name] = $typed
    Invoke-EXOCommand -Name 'Set-Place' -Parameters $params | Out-Null
}

#endregion

#region UI ---------------------------------------------------------------------

function Show-PlaceEditDialog {
    param([string]$Name, [string]$Current, [type]$Type)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Edit $Name"
    $dlg.Size = New-Object System.Drawing.Size(520, 210)
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.StartPosition = 'CenterParent'
    $dlg.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(12, 12)
    $lbl.Size = New-Object System.Drawing.Size(480, 34)
    $dlg.Controls.Add($lbl)

    $under = $null
    if ($Type) { $under = [System.Nullable]::GetUnderlyingType($Type) }
    $tn = if ($under) { $under.Name } elseif ($Type) { $Type.Name } else { 'String' }

    $input = $null
    if ($tn -eq 'Boolean') {
        # A dropdown, not a textbox. Typing "false" into a textbox and having it
        # become $true is the coercion trap ConvertTo-CmdletParameterValue
        # exists to block -- removing the chance to type it at all is better.
        $lbl.Text = "$Name -- yes or no."
        $input = New-Object System.Windows.Forms.ComboBox
        $input.DropDownStyle = 'DropDownList'
        [void]$input.Items.AddRange(@('True', 'False'))
        $input.SelectedItem = $(if ($Current -match '(?i)^true$') { 'True' } else { 'False' })
    } else {
        $hint = switch ($tn) {
            'Int32'    { 'Whole number. Blank clears it.' }
            'String[]' { 'Comma-separated. Blank clears all.' }
            default    { 'Blank clears the value.' }
        }
        $lbl.Text = "$Name -- $hint"
        $input = New-Object System.Windows.Forms.TextBox
        $input.Text = $Current
    }
    $input.Location = New-Object System.Drawing.Point(12, 54)
    $input.Size = New-Object System.Drawing.Size(480, 24)
    $dlg.Controls.Add($input)

    $lblWarn = New-Object System.Windows.Forms.Label
    $lblWarn.Location = New-Object System.Drawing.Point(12, 84)
    $lblWarn.Size = New-Object System.Drawing.Size(480, 34)
    $lblWarn.Text = 'Saving writes to Exchange Online immediately. There is no undo.'
    $dlg.Controls.Add($lblWarn)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Save'; $ok.Location = New-Object System.Drawing.Point(296, 128)
    $ok.Size = New-Object System.Drawing.Size(90, 28)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($ok); $dlg.AcceptButton = $ok

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'; $cancel.Location = New-Object System.Drawing.Point(394, 128)
    $cancel.Size = New-Object System.Drawing.Size(90, 28)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($cancel); $dlg.CancelButton = $cancel

    Set-ConsoleTheme -Form $dlg
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return "$($input.Text)"
}

function Show-PlaceForm {
    # NO -Smtp PARAMETER. The room can change while the window is open, so it
    # lives in $script:Target rather than a parameter fixed at build time.

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Room Place'
    $form.ClientSize = New-Object System.Drawing.Size(844, 618)
    $form.StartPosition = 'CenterScreen'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $lblHead = New-Object System.Windows.Forms.Label
    $lblHead.Location = New-Object System.Drawing.Point(12, 10)
    $lblHead.Size = New-Object System.Drawing.Size(820, 34)
    $form.Controls.Add($lblHead)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(12, 50)
    $grid.Size = New-Object System.Drawing.Size(820, 480)
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.ReadOnly = $true
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $false
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.RowTemplate.Height = 24
    foreach ($c in @(
        @{ N = 'Property'; W = 30 }, @{ N = 'Value'; W = 52 }, @{ N = 'State'; W = 18 })) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c.N; $col.HeaderText = $c.N; $col.FillWeight = $c.W
        $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
        [void]$grid.Columns.Add($col)
    }
    $form.Controls.Add($grid)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(12, 538)
    $lblStatus.Size = New-Object System.Drawing.Size(820, 34)
    $form.Controls.Add($lblStatus)

    $btnEdit = New-Object System.Windows.Forms.Button
    $btnEdit.Text = 'Edit'; $btnEdit.Location = New-Object System.Drawing.Point(12, 576)
    $btnEdit.Size = New-Object System.Drawing.Size(110, 30)
    $form.Controls.Add($btnEdit)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = 'Refresh'; $btnRefresh.Location = New-Object System.Drawing.Point(130, 576)
    $btnRefresh.Size = New-Object System.Drawing.Size(110, 30)
    $form.Controls.Add($btnRefresh)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Location = New-Object System.Drawing.Point(722, 576)
    $btnClose.Size = New-Object System.Drawing.Size(110, 30)
    $form.Controls.Add($btnClose)
    $btnClose.Add_Click({ $form.Close() })

    $load = {
        if (-not "$script:Target".Trim()) {
            $grid.Rows.Clear()
            $form.Text      = 'Room Place'
            $lblStatus.Text = 'Enter a room address above and press Look Up.'
            return
        }
        $form.Text   = "Room Place -- $script:Target"
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $grid.Rows.Clear()
            $rec = Get-PlaceRecord -Smtp $script:Target
            $script:PlaceRoom     = $rec.Place
            $script:PlaceIsSynced = $rec.IsSynced

            if ($rec.Error) { $lblStatus.Text = $rec.Error; return }
            if (-not $rec.Place) { $lblStatus.Text = "No place record returned for '$script:Target'."; return }

            $warn = ''
            if (-not $rec.IsRoom) {
                $warn = ("This is a $($rec.Mailbox.RecipientTypeDetails), not a room. " +
                         "Get-Place answers for it but almost everything will be blank.")
            }
            $lblHead.Text = ("$($rec.Mailbox.DisplayName)  --  $($rec.Mailbox.RecipientTypeDetails)" +
                             $(if ($rec.IsSynced) { '  --  synced from on-premises' } else { '  --  cloud-only' }) +
                             $(if ($warn) { "`r`n$warn" } else { '' }))

            $props = @($rec.Place.PSObject.Properties |
                       Where-Object { $script:PlaceExcluded -notcontains $_.Name } |
                       Sort-Object Name)
            foreach ($p in $props) {
                $state = Get-PlaceRowState -Name $p.Name
                $i = $grid.Rows.Add($p.Name, (Format-PlaceValue $p.Value), $state)
                if ($state -ne 'Editable') {
                    $grid.Rows[$i].Cells['State'].Style.ForeColor = [System.Drawing.Color]::DimGray
                }
            }

            $ed = @($props | Where-Object { (Get-PlaceRowState -Name $_.Name) -eq 'Editable' }).Count
            $blocked = @($props | Where-Object { (Get-PlaceRowState -Name $_.Name) -eq 'On-prem' }).Count
            $msg = "$($props.Count) properties. $ed editable here."
            if ($blocked -gt 0) {
                $msg += ("  $blocked can only be changed on-premises -- this room syncs up, and " +
                         "Set-Place does not persist those. Use Set-User or Set-Mailbox on-premises.")
            }
            $lblStatus.Text = $msg
        } catch {
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Reading the place record'
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $doEdit = {
        if ($grid.SelectedRows.Count -eq 0) { return }
        $row   = $grid.SelectedRows[0]
        $name  = "$($row.Cells['Property'].Value)"
        $state = "$($row.Cells['State'].Value)"

        if ($state -eq 'Read-only') {
            $null = Show-ToolDialog(
                "'$name' is returned by Get-Place but Set-Place has no parameter for it, so it cannot be changed from here.",
                'Read-only', 'OK', 'Information')
            return
        }
        if ($state -eq 'On-prem') {
            $null = Show-ToolDialog(
                ("'$name' cannot be set on this room." + "`r`n`r`n" +
                 "The room mailbox syncs from on-premises, and Microsoft blocks Set-Place from " +
                 "changing this property on synchronised rooms. The call can appear to succeed " +
                 "and then not persist." + "`r`n`r`n" +
                 "Change it on-premises with Set-User or Set-Mailbox and let it sync up."),
                'Set on-premises instead', 'OK', 'Warning')
            return
        }

        $current = "$($row.Cells['Value'].Value)"
        $new = Show-PlaceEditDialog -Name $name -Current $current -Type $script:PlaceWritable[$name]
        if ($null -eq $new) { return }
        if ($new -eq $current) { $lblStatus.Text = "$name unchanged."; return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Set-PlaceProperty -Smtp $script:Target -Name $name -Value $new
            Write-CompletedLog -Action 'Configure' -Target $script:Target -Result 'Success' `
                -Detail "Set-Place $name : '$current' -> '$new'"
            $lblStatus.Text = "$name updated."
            & $load
        } catch {
            # Logged, not just shown. A failed write against production is the
            # row you go looking for later, and a message box does not survive
            # the click that dismisses it.
            Write-FailedLog -Action 'Configure' -Target $script:Target -ErrorRecord $_ `
                -Detail "Set-Place $name : '$current' -> '$new'"
            $lblStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context "Setting $name"
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $btnEdit.Add_Click($doEdit)
    $grid.Add_CellDoubleClick({ if ($_.RowIndex -ge 0) { & $doEdit } })
    $btnRefresh.Add_Click($load)

    # BEFORE Add-ToolHelpButton: the bar shifts every existing control down and
    # grows the form, and the help button places itself from ClientSize at run
    # time. Reversed, the '?' lands 40px above where the window now ends.
    $bar = Add-ToolTargetBar -Form $form -Label 'Room address:' -InitialValue $script:Target
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        & $load
    })

    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Room Place' | Out-Null
    Set-ConsoleTheme -Form $form
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    & $load
    [void]$form.ShowDialog()
}

#endregion

$script:PlaceWritable = Get-EXOWritableParameter -Name 'Set-Place'
if ($script:PlaceWritable.Count -eq 0) {
    # Empty means Set-Place could not be resolved, which means no EXO session --
    # not "this cmdlet has no parameters". Saying so beats opening a window
    # where every row reads Read-only for no stated reason.
    $null = Show-ToolDialog(
        ("Set-Place could not be found in this session." + "`r`n`r`n" +
         "Connect to Exchange Online first -- Get-Place and Set-Place exist only in the cloud."),
        'Not connected', 'OK', 'Warning')
    return
}

Show-PlaceForm
