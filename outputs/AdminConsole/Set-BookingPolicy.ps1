#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# ---------------------------------------------------------------------------
# CLOUD BUILD. Everything goes through Invoke-EXOCommand.
#
# Get-/Set-CalendarProcessing exist in BOTH Exchanges, so the bare names are
# ambiguous in a process holding both sessions. Never call them bare here.
# ---------------------------------------------------------------------------
<#
    Set-BookingPolicy.ps1
    Restrict a room to a named list of bookers -- Layer 2 tool script.

    Refactored from a working one-off that set BookInPolicy and nothing else.
    See the block below for why that one-off did not do what it appeared to.

    Author   : Glenn Armbruster
    Revised  : 2026-08-09
#>
# TOOLKIT-HELP-BEGIN
# Controls WHO is allowed to book a room.
#
# Pick a room, add or remove people, apply. Everyone not on the list has their
# request declined automatically.
#
# Type an address or alias in the box and press Enter. Bare aliases work -- if
# 'jsmith' does not resolve on its own the tenant's own domain is tried before
# giving up. Select a row and press Delete, or use Remove, to take someone off.
#
# THE LIST ALONE DOES NOTHING. A room also carries a master switch,
# AllBookInPolicy, and while it is True the room accepts bookings from everyone
# and the list is ignored entirely. Restricting a room means setting BOTH. This
# tool does that together, and says so before it writes.
#
# An EMPTY list with restriction on means nobody can book the room at all. That
# is a valid thing to want and a terrible thing to do by accident, so it is
# confirmed separately.
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
    WHAT THE ORIGINAL SCRIPT DID, AND WHY IT LOOKED CORRECT.

        Set-CalendarProcessing -Identity $Room -BookInPolicy @()
        Set-CalendarProcessing -Identity $Room -BookInPolicy "a","b","c","d"
        Get-CalendarProcessing $Room | Select -ExpandProperty BookInPolicy

    1. AllBookInPolicy WAS NEVER SET. BookInPolicy is only consulted when
       AllBookInPolicy is $false, and room mailboxes ship with it $true. So the
       room carried on accepting bookings from everyone while holding a list of
       four names that nothing read.

       The verification line is what made this convincing. It reads the list
       straight back and prints four names, so the script appears to have
       worked. It confirmed the value was STORED; it could not confirm the
       value was USED.

    2. NOTHING RESOLVED THE NAMES. A typo in a booking list is not rejected
       loudly -- it just means one person cannot book, discovered when they try.

    3. The clear-then-set was redundant: -BookInPolicy replaces the whole list.

    Also: a backtick line continuation with one trailing space turns the
    following lines into separate statements and the command runs with no list.
#>

$script:RequiredCommon = @('Get-MailDomain')
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

#region Data -------------------------------------------------------------------

function Get-BookingState {
    param([Parameter(Mandatory)][string]$Smtp)
    $out = [PSCustomObject]@{ Mailbox = $null; Cal = $null; IsResource = $false; Error = '' }
    try {
        $out.Mailbox = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop'
        }
        $out.IsResource = ("$($out.Mailbox.RecipientTypeDetails)" -match '(?i)room|equipment|workspace')
        $out.Cal = Invoke-EXOCommand -Name 'Get-CalendarProcessing' -Parameters @{
            Identity = $Smtp; ErrorAction = 'Stop'
        }
    } catch {
        $out.Error = $_.Exception.Message
    }
    return $out
}

function Resolve-Booker {
    <#
        One entry -> a recipient, or a stated reason why not.

        THREE THINGS ARE HANDLED HERE, and each one exists because of something
        real:

        Legacy DNs. Get-CalendarProcessing does not return addresses -- it
        returns X500 distinguished names like
        '/o=ExchangeLabs/ou=.../cn=Recipients/cn=<guid>-Russell, Me'. A list of
        those in a textbox is unreadable and uneditable, which is what the first
        version of this form asked people to work with. They ARE valid
        identities, so they resolve back to a person and this is how the grid
        gets real names.

        Bare aliases. 'jsmith' on its own may not resolve; 'jsmith@<tenant>'
        usually does. The tenant's own primary domain is tried as a second
        attempt rather than making the operator remember which form works.

        Ambiguity. Two matches is NOT a resolution. Picking the first would be a
        coin toss written into a permissions list.

        Get-Recipient rather than Get-Mailbox, because a booking list may
        legitimately hold a distribution or mail-enabled security group.
    #>
    param([Parameter(Mandatory)][string]$Entry)

    $e = "$Entry".Trim()
    $out = [PSCustomObject]@{
        Entry = $e; Resolved = $false; Smtp = ''; Display = ''; Type = ''; Reason = ''
    }
    if (-not $e) { $out.Reason = 'blank'; return $out }

    $tries = @($e)
    if ($e -notmatch '@' -and $e -notmatch '^/o=') {
        $dom = Get-MailDomain
        if ($dom) { $tries += "$e@$dom" }
    }

    $lastReason = 'no match'
    foreach ($t in $tries) {
        try {
            $r = @(Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{
                Identity = $t; ErrorAction = 'Stop'
            })
            if ($r.Count -eq 0) { $lastReason = 'no match'; continue }
            if ($r.Count -gt 1) {
                $out.Reason = "matches $($r.Count) recipients -- use the address"
                return $out
            }
            $out.Resolved = $true
            $out.Smtp     = "$($r[0].PrimarySmtpAddress)"
            $out.Display  = "$($r[0].DisplayName)"
            $out.Type     = "$($r[0].RecipientTypeDetails)"
            return $out
        } catch {
            $lastReason = $_.Exception.Message
        }
    }
    $out.Reason = $lastReason
    return $out
}

#endregion

#region UI ---------------------------------------------------------------------

function Show-BookingPolicyForm {
    # NO -Smtp PARAMETER. The address can change while the window is open, so
    # it lives in $script:Target rather than a parameter fixed at build time.

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Booking Policy -- $script:Target"
    $form.ClientSize = New-Object System.Drawing.Size(844, 680)
    $form.StartPosition = 'CenterScreen'
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $lblHead = New-Object System.Windows.Forms.Label
    $lblHead.Location = New-Object System.Drawing.Point(12, 10)
    $lblHead.Size = New-Object System.Drawing.Size(820, 50)
    $form.Controls.Add($lblHead)

    # --- add box -----------------------------------------------------------
    $lblAdd = New-Object System.Windows.Forms.Label
    $lblAdd.Location = New-Object System.Drawing.Point(12, 66)
    $lblAdd.Size = New-Object System.Drawing.Size(300, 18)
    $lblAdd.Text = 'Add someone -- address or alias, then Enter:'
    $form.Controls.Add($lblAdd)

    $txtAdd = New-Object System.Windows.Forms.TextBox
    $txtAdd.Location = New-Object System.Drawing.Point(12, 86)
    $txtAdd.Size = New-Object System.Drawing.Size(520, 24)
    $form.Controls.Add($txtAdd)

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = 'Add'; $btnAdd.Location = New-Object System.Drawing.Point(540, 84)
    $btnAdd.Size = New-Object System.Drawing.Size(100, 28)
    $form.Controls.Add($btnAdd)

    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text = 'Remove'; $btnRemove.Location = New-Object System.Drawing.Point(648, 84)
    $btnRemove.Size = New-Object System.Drawing.Size(100, 28)
    $form.Controls.Add($btnRemove)

    # --- the list ----------------------------------------------------------
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(12, 120)
    $grid.Size = New-Object System.Drawing.Size(820, 300)
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.ReadOnly = $true
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $true
    $grid.AutoSizeColumnsMode = 'Fill'
    $grid.RowTemplate.Height = 24
    foreach ($c in @(
        @{ N = 'Name'; W = 32 }, @{ N = 'Address'; W = 40 },
        @{ N = 'Type'; W = 20 },  @{ N = 'State'; W = 12 })) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c.N; $col.HeaderText = $c.N; $col.FillWeight = $c.W
        $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
        [void]$grid.Columns.Add($col)
    }
    $form.Controls.Add($grid)

    $chkRestrict = New-Object System.Windows.Forms.CheckBox
    $chkRestrict.Location = New-Object System.Drawing.Point(12, 428)
    $chkRestrict.Size = New-Object System.Drawing.Size(820, 20)
    $chkRestrict.Text = 'Restrict booking to this list  (sets AllBookInPolicy to False)'
    $chkRestrict.Checked = $true
    $form.Controls.Add($chkRestrict)

    $lblWhy = New-Object System.Windows.Forms.Label
    $lblWhy.Location = New-Object System.Drawing.Point(30, 450)
    $lblWhy.Size = New-Object System.Drawing.Size(802, 32)
    $lblWhy.Text = ('Leave this ticked. While AllBookInPolicy is True the room accepts bookings ' +
                    'from everyone and the list above is ignored entirely.')
    $form.Controls.Add($lblWhy)

    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text = 'Apply'; $btnApply.Location = New-Object System.Drawing.Point(12, 488)
    $btnApply.Size = New-Object System.Drawing.Size(120, 30)
    $form.Controls.Add($btnApply)

    $btnReload = New-Object System.Windows.Forms.Button
    $btnReload.Text = 'Reload'; $btnReload.Location = New-Object System.Drawing.Point(140, 488)
    $btnReload.Size = New-Object System.Drawing.Size(120, 30)
    $form.Controls.Add($btnReload)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Close'; $btnClose.Location = New-Object System.Drawing.Point(712, 488)
    $btnClose.Size = New-Object System.Drawing.Size(120, 30)
    $form.Controls.Add($btnClose)
    $btnClose.Add_Click({ $form.Close() })

    $txtOut = New-Object System.Windows.Forms.TextBox
    $txtOut.Location = New-Object System.Drawing.Point(12, 528)
    $txtOut.Size = New-Object System.Drawing.Size(820, 140)
    $txtOut.Multiline = $true
    $txtOut.ScrollBars = 'Vertical'
    $txtOut.ReadOnly = $true
    $txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
    $form.Controls.Add($txtOut)

    # --- helpers -----------------------------------------------------------

    $addRow = {
        <#
            One row. Tag carries what will actually be WRITTEN, which is not
            always what is displayed.

            For a resolved recipient that is the primary SMTP address. For an
            entry that could NOT be resolved but is already on the room, it is
            the original value read from the tenant -- see the load routine for
            why that distinction protects someone's access.
        #>
        param($Res, [string]$Original, [string]$State)

        $write = $(if ($Res.Resolved) { $Res.Smtp } else { $Original })
        $name  = $(if ($Res.Resolved) { $Res.Display } else { '(could not resolve)' })
        $addr  = $(if ($Res.Resolved) { $Res.Smtp }    else { $Original })
        $type  = $(if ($Res.Resolved) { $Res.Type }    else { '' })

        $i = $grid.Rows.Add($name, $addr, $type, $State)
        $grid.Rows[$i].Tag = @{ Write = $write; Resolved = [bool]$Res.Resolved }
        if (-not $Res.Resolved) {
            $grid.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkOrange
        }
        return $i
    }

    $alreadyListed = {
        param([string]$script:Target2)
        foreach ($r in $grid.Rows) {
            if ("$($r.Cells['Address'].Value)" -ieq $script:Target2) { return $true }
        }
        return $false
    }

    $load = {
        # OPENED WITH NO ADDRESS IS A NORMAL STATE NOW -- the window appears
        # before anything has been typed, so the refresh must not run empty.
        if (-not "$script:Target".Trim()) {
            $form.Text = 'Booking Policy'
            return
        }
        $form.Text = 'Booking Policy: ' + $script:Target
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $grid.Rows.Clear()
            $st = Get-BookingState -Smtp $script:Target
            if ($st.Error) { $txtOut.Text = $st.Error; return }

            $mode    = "$($st.Cal.AutomateProcessing)"
            $allBook = [bool]$st.Cal.AllBookInPolicy
            $current = @($st.Cal.BookInPolicy | ForEach-Object { "$_" } | Where-Object { $_ })

            $head = "$($st.Mailbox.DisplayName)  --  $($st.Mailbox.RecipientTypeDetails)"
            $head += "`r`nAutomateProcessing: $mode      AllBookInPolicy: $allBook      list holds $($current.Count)"
            if (-not $st.IsResource) {
                $head += "`r`nThis is not a room or equipment mailbox -- booking policy does almost nothing on it."
            }
            $lblHead.Text = $head

            $unresolved = 0
            foreach ($c in $current) {
                $r = Resolve-Booker -Entry $c
                if (-not $r.Resolved) { $unresolved++ }
                [void](& $addRow $r $c 'on room')
            }

            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine('CURRENT STATE')
            [void]$sb.AppendLine('')
            if ($mode -ne 'AutoAccept') {
                [void]$sb.AppendLine("  WARNING: AutomateProcessing is '$mode', not AutoAccept.")
                [void]$sb.AppendLine('  Nothing is decided automatically, so a booking list changes nothing.')
                [void]$sb.AppendLine('')
            }
            if ($allBook) {
                [void]$sb.AppendLine('  AllBookInPolicy is True -- EVERYONE can book this room right now.')
                if ($current.Count -gt 0) {
                    [void]$sb.AppendLine("  The $($current.Count) name(s) listed above are stored but ignored.")
                }
            } else {
                [void]$sb.AppendLine('  AllBookInPolicy is False -- only the list above may book.')
                if ($current.Count -eq 0) {
                    [void]$sb.AppendLine('  The list is EMPTY, so nobody can book this room at all.')
                }
            }
            if ($unresolved -gt 0) {
                [void]$sb.AppendLine('')
                [void]$sb.AppendLine("  $unresolved entry/entries could not be resolved to a recipient.")
                [void]$sb.AppendLine('  They are kept exactly as stored and will be written back unchanged,')
                [void]$sb.AppendLine('  so applying will not quietly remove anyone. Remove them deliberately')
                [void]$sb.AppendLine('  if they are stale.')
            }
            $txtOut.Text = $sb.ToString()
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Reading booking policy'
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $doAdd = {
        $entry = "$($txtAdd.Text)".Trim()
        if (-not $entry) { return }
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $r = Resolve-Booker -Entry $entry
            if (-not $r.Resolved) {
                # Not added. A row that does not resolve is a row that will not
                # work, and adding it anyway just moves the discovery later.
                $txtOut.Text = "Could not add '$entry' -- $($r.Reason)." + [Environment]::NewLine +
                               'Nothing has changed. Try the full address.'
                return
            }
            if (& $alreadyListed $r.Smtp) {
                $txtOut.Text = "$($r.Display) is already on the list."
                $txtAdd.Text = ''
                return
            }
            $i = & $addRow $r $r.Smtp 'to add'
            $grid.Rows[$i].Cells['State'].Style.ForeColor = [System.Drawing.Color]::SeaGreen
            $txtAdd.Text = ''
            $txtOut.Text = "$($r.Display) <$($r.Smtp)> queued. Nothing is written until you press Apply."
        } catch {
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context "Adding '$entry'"
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $doRemove = {
        if ($grid.SelectedRows.Count -eq 0) {
            $txtOut.Text = 'Select a row first.'
            return
        }
        $names = @()
        foreach ($r in @($grid.SelectedRows)) {
            $names += "$($r.Cells['Name'].Value)"
            $grid.Rows.Remove($r)
        }
        $txtOut.Text = ("Removed from the list: " + ($names -join ', ') +
                        [Environment]::NewLine +
                        'Nothing is written until you press Apply.')
    }

    $doApply = {
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $writes = @()
            $shown  = @()
            foreach ($r in $grid.Rows) {
                $tag = $r.Tag
                if (-not $tag) { continue }
                $writes += "$($tag.Write)"
                $shown  += "$($r.Cells['Name'].Value)"
            }
            $restrict = $chkRestrict.Checked

            if ($writes.Count -eq 0 -and $restrict) {
                $a = Show-ToolDialog(
                    ('The list is empty and restriction is on.' + "`r`n`r`n" +
                     'This will stop EVERYONE from booking this room, including the people who ' +
                     'can book it today.' + "`r`n`r`nIs that what you want?"),
                    'Nobody will be able to book', 'YesNo', 'Warning')
                if ($a -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            }
            if (-not $restrict) {
                $a = Show-ToolDialog(
                    ('Restriction is unticked, so AllBookInPolicy will be left as it is.' + "`r`n`r`n" +
                     'If it is True the room will keep accepting bookings from everyone and this ' +
                     'list will have no effect.' + "`r`n`r`nApply anyway?"),
                    'List may have no effect', 'YesNo', 'Warning')
                if ($a -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            }

            $confirm = ("Room:  $script:Target" + "`r`n`r`n" +
                        "Allow these $($writes.Count) to book:" + "`r`n  " + ($shown -join "`r`n  ") + "`r`n`r`n" +
                        $(if ($restrict) { 'AllBookInPolicy will be set to False, so everyone else is declined.' }
                          else { 'AllBookInPolicy will be left unchanged.' }) +
                        "`r`n`r`nApply?")
            $a = Show-ToolDialog($confirm, 'Confirm', 'OKCancel', 'Question')
            if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

            # ONE call. -BookInPolicy replaces the whole collection, so the
            # original's clear-then-set only opened a window with no list.
            $params = @{ Identity = $script:Target; BookInPolicy = $writes; ErrorAction = 'Stop' }
            if ($restrict) { $params['AllBookInPolicy'] = $false }
            Invoke-EXOCommand -Name 'Set-CalendarProcessing' -Parameters $params | Out-Null

            Write-CompletedLog -Action 'Configure' -Target $script:Target -Result 'Success' `
                -Detail ("BookInPolicy set to $($writes.Count) recipient(s)" +
                         $(if ($restrict) { '; AllBookInPolicy=False' } else { '' }) +
                         ": $($writes -join '; ')")

            # Read back and prove BOTH halves -- the original only proved the
            # list was stored, which was the half that was already working.
            $after = Invoke-EXOCommand -Name 'Get-CalendarProcessing' -Parameters @{
                Identity = $script:Target; ErrorAction = 'Stop'
            }
            $sb = New-Object System.Text.StringBuilder
            [void]$sb.AppendLine('APPLIED -- read back from the tenant')
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine("  AutomateProcessing : $($after.AutomateProcessing)")
            [void]$sb.AppendLine("  AllBookInPolicy    : $($after.AllBookInPolicy)")
            [void]$sb.AppendLine("  BookInPolicy       : $(@($after.BookInPolicy).Count) entries")
            if ([bool]$after.AllBookInPolicy) {
                [void]$sb.AppendLine('')
                [void]$sb.AppendLine('  NOTE: AllBookInPolicy is still True, so this list is not being used.')
            }
            if ("$($after.AutomateProcessing)" -ne 'AutoAccept') {
                [void]$sb.AppendLine("  NOTE: AutomateProcessing is '$($after.AutomateProcessing)' -- bookings are not decided automatically.")
            }
            $txtOut.Text = $sb.ToString()
            & $load
        } catch {
            Write-FailedLog -Action 'Configure' -Target $script:Target -ErrorRecord $_ `
                -Detail 'Set booking policy'
            $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Applying the booking policy'
        } finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }

    $btnAdd.Add_Click($doAdd)
    $btnRemove.Add_Click($doRemove)
    $btnApply.Add_Click($doApply)
    $btnReload.Add_Click($load)

    # Enter adds, rather than triggering whatever the form's default button
    # happens to be. SuppressKeyPress as well as Handled, or the textbox still
    # gets the keystroke and Windows plays the error beep.
    $txtAdd.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $_.Handled = $true
            $_.SuppressKeyPress = $true
            & $doAdd
        }
    })
    $grid.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Delete) {
            $_.Handled = $true
            & $doRemove
        }
    })
    # BEFORE Add-ToolHelpButton: the bar shifts every existing control down
    # and grows the form, and the help button places itself from ClientSize at
    # run time.
    $bar = Add-ToolTargetBar -Form $form -Label 'Room address:' -InitialValue $script:Target
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        & $load
    })


    Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Booking Policy' | Out-Null
    Set-ConsoleTheme -Form $form
    Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
    & $load
    $form.Add_Shown({ $txtAdd.Focus() })
    [void]$form.ShowDialog()
}

#endregion

Show-BookingPolicyForm
