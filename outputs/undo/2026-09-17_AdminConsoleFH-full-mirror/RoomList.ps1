#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# ---------------------------------------------------------------------------
# CLOUD BUILD. Everything goes through Invoke-EXOCommand.
#
# Get-/Add-/Remove-DistributionGroupMember and Get-DistributionGroup exist in
# BOTH Exchanges. The bare names are ambiguous in a process holding both
# sessions, and a membership change landing in the wrong directory is silent.
# ---------------------------------------------------------------------------
<#
    RoomList.ps1
    Room list membership -- Layer 2 tool script.

    Which room lists a room belongs to, and adding or removing it.

    Author   : Glenn Armbruster
    Revised  : 2026-08-09
#>
# TOOLKIT-HELP-BEGIN
# Why a room does not appear in Outlook's Room Finder.
#
# The Room Finder does not list rooms. It lists ROOM LISTS, and then the rooms
# inside them. A room mailbox that belongs to no room list is invisible there no
# matter how correct the mailbox and its place record are -- and the symptom
# people report is "the room does not exist", which points nowhere near this.
#
# A room list is an ordinary distribution group flagged as a room list. It
# cannot be flagged after the fact through this tool: a normal group converted
# to a room list keeps behaving oddly in the Room Finder, so New is the safe
# route and that is all this offers.
#
# Membership changes take effect immediately, but Outlook and Teams cache the
# Room Finder aggressively -- allow up to 24 hours before concluding it did not
# work. Verify here rather than in Outlook.
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

#region Data -------------------------------------------------------------------

function Get-RoomListInventory {
    <#
        Every room list, and whether this room is in it.

        MEMBERSHIP IS ASKED PER LIST, which is one call each. There is no
        reverse lookup worth trusting -- a room mailbox does not carry a usable
        "lists I belong to" property, and reading MemberOfGroup on the recipient
        returns security group membership, not distribution group membership.

        Tenants typically have a handful of room lists, so a handful of calls is
        the honest cost. If that ever becomes slow the fix is caching, not
        guessing from the recipient object.
    #>
    param([Parameter(Mandatory)][string]$Smtp)

    $out = [PSCustomObject]@{ Lists = @(); Error = '' }
    try {
        $lists = @(Invoke-EXOCommand -Name 'Get-DistributionGroup' -Parameters @{
            RecipientTypeDetails = 'RoomList'
            ResultSize           = 'Unlimited'
            ErrorAction          = 'Stop'
            WarningAction        = 'SilentlyContinue'
        })
    } catch {
        $out.Error = "Could not read room lists: $($_.Exception.Message)"
        return $out
    }

    $rows = @()
    foreach ($l in $lists) {
        $isMember = $false
        $count    = 0
        $err      = ''
        try {
            $members = @(Invoke-EXOCommand -Name 'Get-DistributionGroupMember' -Parameters @{
                Identity      = "$($l.Identity)"
                ResultSize    = 'Unlimited'
                ErrorAction   = 'Stop'
                WarningAction = 'SilentlyContinue'
            })
            $count = $members.Count
            foreach ($m in $members) {
                if ("$($m.PrimarySmtpAddress)" -ieq $Smtp) { $isMember = $true; break }
            }
        } catch {
            # A list that cannot be read is NOT reported as "not a member" --
            # that would invite adding the room a second time. It is reported as
            # unknown, and the row refuses to act.
            $err = $_.Exception.Message
        }
        $rows += [PSCustomObject]@{
            Name     = "$($l.DisplayName)"
            Smtp     = "$($l.PrimarySmtpAddress)"
            Identity = "$($l.Identity)"
            Member   = $isMember
            Count    = $count
            Error    = $err
        }
    }
    $out.Lists = @($rows | Sort-Object Name)
    return $out
}

#endregion

#region UI ---------------------------------------------------------------------

function Show-RoomListForm {
    # NO -Smtp PARAMETER. The address can change while the window is open, so
    # it lives in $script:Target rather than a parameter fixed at build time.

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Room Lists -- $script:Target"
    $form.ClientSize = New-Object System.Drawing.Size(844, 598)
    $form.StartPosition = 'CenterScreen'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $lblHead = New-Object System.Windows.Forms.Label
    $lblHead.Location = New-Object System.Drawing.Point(12, 10)
    $lblHead.Size = New-Object System.Drawing.Size(820, 50)
    $form.Controls.Add($lblHead)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(12, 66)
    $grid.Size = New-Object System.Drawing.Size(820, 340)
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.ReadOnly = $true
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $false
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.RowTemplate.Height = 24
    foreach ($c in @(
        @{ N = 'Member'; W = 12 }, @{ N = 'Room list'; W = 34 },
        @{ N = 'Address'; W = 38 }, @{ N = 'Rooms'; W = 16 })) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c.N; $col.HeaderText = $c.N; $col.FillWeight = $c.W
        $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
        [void]$grid.Columns.Add($col)
    }
    $form.Controls.Add($grid)

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = 'Add to list'; $btnAdd.Location = New-Object System.Drawing.Point(12, 416)
    $btnAdd.Size = New-Object System.Drawing.Size(120, 30)
    $form.Controls.Add($btnAdd)

    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text = 'Remove'; $btnRemove.Location = New-Object System.Drawing.Point(140, 416)
    $btnRemove.Size = New-Object System.Drawing.Size(120, 30)
    $form.Controls.Add($btnRemove)

    $btnNew = New-Object System.Windows.Forms.Button
    $btnNew.Text = 'New room list'; $btnNew.Location = New-Object System.Drawing.Point(268, 416)
    $btnNew.Size = New-Object System.Drawing.Size(120, 30)
    $form.Controls.Add($btnNew)

    $btnReload = New-Object System.Windows.Forms.Button
    $btnReload.Text = 'Reload'; $btnReload.Location = New-Object System.Drawing.Point(396, 416)
    $btnReload.Size = New-Object System.Drawing.Size(120, 30)
    $form.Controls.Add($btnReload)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Location = New-Object System.Drawing.Point(712, 416)
    $btnClose.Size = New-Object System.Drawing.Size(120, 30)
    $form.Controls.Add($btnClose)
    $btnClose.Add_Click({ $form.Close() })

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Location = New-Object System.Drawing.Point(12, 456)
    $txtOut.Size = New-Object System.Drawing.Size(820, 130)
    $txtOut.Multiline = $true
    $txtOut.ScrollBars = 'Vertical'
    $txtOut.ReadOnly = $true
    $txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtOut)

    $load = {
        # OPENED WITH NO ADDRESS IS A NORMAL STATE NOW -- the window appears
        # before anything has been typed, so the refresh must not run empty.
        if (-not "$script:Target".Trim()) {
            $form.Text = 'Room Lists'
            return
        }
        $form.Text = 'Room Lists: ' + $script:Target
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $grid.Rows.Clear()
            $inv = Get-RoomListInventory -Smtp $script:Target
            if ($inv.Error) { $txtOut.Text = $inv.Error; return }

            $memberOf = 0
            foreach ($l in $inv.Lists) {
                $mark = $(if ($l.Error) { '?' } elseif ($l.Member) { 'YES' } else { '' })
                if ($l.Member) { $memberOf++ }
                $i = $grid.Rows.Add($mark, $l.Name, $l.Smtp, $(if ($l.Error) { '' } else { $l.Count }))
                $grid.Rows[$i].Tag = $l
                if ($l.Member) {
                    $grid.Rows[$i].Cells['Member'].Style.ForeColor = [System.Drawing.Color]::SeaGreen
                }
                if ($l.Error) {
                    $grid.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkOrange
                }
            }

            $head = "$script:Target"
            $head += "`r`n$($inv.Lists.Count) room list(s) in the tenant. This room is in $memberOf of them."
            if ($inv.Lists.Count -eq 0) {
                $head += "`r`nThere are no room lists at all, so the Room Finder has nothing to show."
            } elseif ($memberOf -eq 0) {
                $head += "`r`nThis room is in NO room list, so it will not appear in the Room Finder."
            }
            $lblHead.Text = $head

            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine('Select a room list, then Add or Remove.')
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine('Outlook and Teams cache the Room Finder -- a change here can take up to')
            [void]$sb.AppendLine('24 hours to show up there. The Member column above is the truth.')
            $bad = @($inv.Lists | Where-Object { $_.Error })
            if ($bad.Count -gt 0) {
                [void]$sb.AppendLine('')
                [void]$sb.AppendLine("$($bad.Count) list(s) could not be read -- shown with '?'. Membership is")
                [void]$sb.AppendLine('UNKNOWN for those, so Add and Remove are refused on them.')
            }
            $txtOut.Text = $sb.ToString()
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Reading room lists'
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $selected = {
        if ($grid.SelectedRows.Count -eq 0) { return $null }
        return $grid.SelectedRows[0].Tag
    }

    $doAdd = {
        $l = & $selected
        if (-not $l) { $txtOut.Text = 'Select a room list first.'; return }
        if ($l.Error) {
            $txtOut.Text = "'$($l.Name)' could not be read, so membership is unknown. Not adding."
            return
        }
        if ($l.Member) { $txtOut.Text = "Already a member of '$($l.Name)'."; return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Invoke-EXOCommand -Name 'Add-DistributionGroupMember' -Parameters @{
                Identity = $l.Identity; Member = $script:Target; ErrorAction = 'Stop'
            } | Out-Null
            Write-CompletedLog -Action 'Add' -Target $script:Target -Result 'Success' `
                -Detail "Added to room list '$($l.Name)' <$($l.Smtp)>"
            $txtOut.Text = "Added to '$($l.Name)'. The Room Finder may take up to 24 hours to catch up."
            & $load
        } catch {
            Write-FailedLog -Action 'Add' -Target $script:Target -ErrorRecord $_ `
                -Detail "Add to room list '$($l.Name)'"
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context "Adding to '$($l.Name)'"
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $doRemove = {
        $l = & $selected
        if (-not $l) { $txtOut.Text = 'Select a room list first.'; return }
        if ($l.Error) {
            $txtOut.Text = "'$($l.Name)' could not be read, so membership is unknown. Not removing."
            return
        }
        if (-not $l.Member) { $txtOut.Text = "Not a member of '$($l.Name)'."; return }

        $a = Show-ToolDialog(
            ("Remove this room from '$($l.Name)'?" + "`r`n`r`n" +
             "If this is its only room list the room stops appearing in the Room Finder."),
            'Confirm', 'OKCancel', 'Question')
        if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            Invoke-EXOCommand -Name 'Remove-DistributionGroupMember' -Parameters @{
                Identity = $l.Identity; Member = $script:Target; Confirm = $false; ErrorAction = 'Stop'
            } | Out-Null
            Write-CompletedLog -Action 'Remove' -Target $script:Target -Result 'Success' `
                -Detail "Removed from room list '$($l.Name)' <$($l.Smtp)>"
            $txtOut.Text = "Removed from '$($l.Name)'."
            & $load
        } catch {
            Write-FailedLog -Action 'Remove' -Target $script:Target -ErrorRecord $_ `
                -Detail "Remove from room list '$($l.Name)'"
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context "Removing from '$($l.Name)'"
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $doNew = {
        <#
            NEW only, never convert.

            Set-DistributionGroup has no supported switch that turns an existing
            group into a room list, and groups converted by other means are a
            known source of Room Finder oddities. Offering only New keeps this
            tool on the documented path.
        #>
        $dlg = New-Object System.Windows.Forms.Form
        $dlg.Text = 'New room list'
        $dlg.Size = New-Object System.Drawing.Size(520, 210)
        $dlg.FormBorderStyle = 'FixedDialog'
        $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false
        $dlg.StartPosition = 'CenterParent'
        $dlg.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

        $l1 = New-Object System.Windows.Forms.Label
        $l1.Text = 'Name -- what people see in the Room Finder, for example a building.'
        $l1.Location = New-Object System.Drawing.Point(12, 12)
        $l1.Size = New-Object System.Drawing.Size(480, 32)
        $dlg.Controls.Add($l1)

        $t1 = New-Object System.Windows.Forms.TextBox
        $t1.Location = New-Object System.Drawing.Point(12, 48)
        $t1.Size = New-Object System.Drawing.Size(480, 24)
        $dlg.Controls.Add($t1)

        $c1 = New-Object System.Windows.Forms.CheckBox
        $c1.Text = 'Add this room to it straight away'
        $c1.Location = New-Object System.Drawing.Point(12, 82)
        $c1.Size = New-Object System.Drawing.Size(480, 20)
        $c1.Checked = $true
        $dlg.Controls.Add($c1)

        $ok = New-Object System.Windows.Forms.Button
        $ok.Text = 'Create'; $ok.Location = New-Object System.Drawing.Point(296, 128)
        $ok.Size = New-Object System.Drawing.Size(90, 28)
        $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dlg.Controls.Add($ok); $dlg.AcceptButton = $ok

        $ca = New-Object System.Windows.Forms.Button
        $ca.Text = 'Cancel'; $ca.Location = New-Object System.Drawing.Point(394, 128)
        $ca.Size = New-Object System.Drawing.Size(90, 28)
        $ca.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $dlg.Controls.Add($ca); $dlg.CancelButton = $ca

        Set-ConsoleTheme -Form $dlg
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $name = "$($t1.Text)".Trim()
        if (-not $name) { $txtOut.Text = 'A name is required.'; return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $p = @{ Name = $name; RoomList = $true; ErrorAction = 'Stop' }
            if ($c1.Checked) { $p['Members'] = @($script:Target) }

            # THE RETURN VALUE IS CAPTURED RATHER THAN DISCARDED, and only for
            # the audit log. New-DistributionGroup hands back the object it just
            # made, including the address Exchange assigned -- which is the one
            # identifier that did not exist a moment ago and cannot be derived
            # from anything on this form.
            #
            # Piping it to Out-Null threw away the only chance to record what
            # was actually created. The name falls back for the case where the
            # create succeeded but returned nothing.
            $created = Invoke-EXOCommand -Name 'New-DistributionGroup' -Parameters $p

            Write-CompletedLog -Action 'Create' `
                -Target (Get-LogTarget $created.PrimarySmtpAddress $name) -Result 'Success' `
                -Detail ("Created room list '$name'" +
                         $(if ($c1.Checked) { " with $script:Target as first member" } else { '' }))
            $txtOut.Text = "Created room list '$name'."
            & $load
        } catch {
            # NO ADDRESS EXISTS HERE and that is not a gap -- the create failed,
            # so Exchange never assigned one. Get-LogTarget rather than a bare
            # $name so the intent is explicit and the rule stays visible.
            Write-FailedLog -Action 'Create' -Target (Get-LogTarget $name) -ErrorRecord $_ `
                -Detail "Create room list '$name'"
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context "Creating '$name'"
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $btnAdd.Add_Click($doAdd)
    $btnRemove.Add_Click($doRemove)
    $btnNew.Add_Click($doNew)
    $btnReload.Add_Click($load)
    # BEFORE Add-ToolHelpButton: the bar shifts every existing control down
    # and grows the form, and the help button places itself from ClientSize at
    # run time.
    $bar = Add-ToolTargetBar -Form $form -Label 'Room address:' -InitialValue $script:Target
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        & $load
    })


    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Room Lists' | Out-Null
    Set-ConsoleTheme -Form $form
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    & $load
    [void]$form.ShowDialog()
}

#endregion

Show-RoomListForm
