#Requires -Version 5.1
<#
    ExOSupport.Editors.ps1  --  the Shared Mailbox Editor and the Group Editor.

    Each does one job: add people, remove people.

      Shared mailbox  Full Access, or Full Access + Send As. Automapping is
                      always on. Remove takes away whatever the selected
                      people have (Full Access and/or Send As).
      Group           add members, remove members.

    EVERY WRITE IS PRECEDED BY TWO THINGS:
      1. the object is re-read and must still pass the console's gate
         (cloud-managed, right type, the operator authorized for its attribute) -- Get-ExsLiveRecipient
      2. every line the operator typed must resolve to exactly one recipient;
         if any line does not, nothing is written

    Every change goes to the local log, success or failure.
#>

# BUILD STAMP. Every ExOSupport file carries one; ExOSupport.ps1 refuses to start
# unless all of them match its own, so a half-copied folder fails with a
# sentence instead of a missing-function error. Bump all four together.
if (-not $script:ExsParts) { $script:ExsParts = @{} }
$script:ExsParts['Editors'] = '2026.10.21.7'

#region Shared Mailbox Editor ---------------------------------------------------

function Show-ExsMailboxEditor {
    param([Parameter(Mandatory)][string]$Identity, $Owner)
    $r = Get-ExsLiveRecipient -Identity $Identity
    $script:Ed = @{ Smtp = "$($r.PrimarySmtpAddress)"; Display = "$($r.DisplayName)" }
    $ed = $script:Ed

    $f = New-Object System.Windows.Forms.Form
    $f.Text = "Shared Mailbox Editor  --  $($ed.Display)"
    $f.ClientSize = New-Object System.Drawing.Size(900, 590)
    $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'; $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $ed.Form = $f

    [void](New-ExsLabel $f "$($ed.Display)   <$($ed.Smtp)>" 12 12 876 20 -Bold)
    [void](New-ExsLabel $f "Shared mailbox   |   $($script:Tenant.Attribute) = $("$($r.($script:Tenant.Attribute))".Trim())" 12 34 876 20 -Muted)

    [void](New-ExsLabel $f 'Who has access now (select one or more rows to remove):' 12 62 520 20 -Bold)
    $ed.Grid = New-ExsGrid $f @(@('Person', 300), @('Full Access', 100), @('Send As', 100)) 12 84 520 410
    $ed.BtnRemove  = New-ExsButton $f 'Remove selected' 12 502 170 32
    $ed.BtnRefresh = New-ExsButton $f 'Refresh' 190 502 100 32

    [void](New-ExsLabel $f 'Add people -- one per line (name, alias or address):' 548 62 340 20 -Bold)
    $ed.List = New-ExsListBox $f 548 84 340 260
    [void](New-ExsLabel $f ("A bare alias gets @$($script:Tenant.PrimaryDomain). Automapping is always on: the mailbox " +
        'appears in their Outlook by itself, usually within an hour.') 548 350 340 52 -Muted)
    $ed.BtnFull = New-ExsButton $f 'Add Full Access' 548 410 340 34
    $ed.BtnBoth = New-ExsButton $f 'Add Full Access + Send As' 548 452 340 34

    $ed.Status = New-ExsLabel $f '' 12 552 770 20 -Muted
    $close = New-ExsButton $f 'Close' 798 546 90 32
    $close.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $f.CancelButton = $close

    $ed.BtnRefresh.Add_Click({ Update-ExsMailboxGrid })
    $ed.BtnFull.Add_Click({ Invoke-ExsMailboxAdd -SendAs:$false })
    $ed.BtnBoth.Add_Click({ Invoke-ExsMailboxAdd -SendAs:$true })
    $ed.BtnRemove.Add_Click({ Invoke-ExsMailboxRemove })
    $f.Add_Shown({ Update-ExsMailboxGrid })

    [void]$f.ShowDialog($Owner)
}

function Set-ExsEdStatus {
    param($Ui, [string]$Text, [switch]$Bad)
    $Ui.Status.ForeColor = $(if ($Bad) { [System.Drawing.Color]::Firebrick } else { [System.Drawing.Color]::DimGray })
    $Ui.Status.Text = $Text
    $Ui.Status.Refresh()
}

function Update-ExsMailboxGrid {
    $ed = $script:Ed
    $ed.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    Set-ExsEdStatus $ed 'Reading permissions...'
    try {
        $rows = @(Get-ExsMailboxAccess -Smtp $ed.Smtp)
        $ed.Grid.Rows.Clear()
        foreach ($x in $rows) {
            $i = $ed.Grid.Rows.Add($x.User, $(if ($x.FullAccess) { 'Yes' } else { '' }), $(if ($x.SendAs) { 'Yes' } else { '' }))
            $ed.Grid.Rows[$i].Tag = $x
        }
        $ed.Grid.ClearSelection()
        Set-ExsEdStatus $ed "$($rows.Count) person(s) with access."
    } catch {
        Set-ExsEdStatus $ed "Could not read permissions: $($_.Exception.Message)" -Bad
    } finally { $ed.Form.Cursor = [System.Windows.Forms.Cursors]::Default }
}

function Resolve-ExsPeopleList {
    <#
        Resolve every line first. Returns the resolved list, or $null after
        showing which lines failed -- in which case nothing is written.
    #>
    param([string]$Text, [string]$Title)
    $lines = @(Get-ExsLines $Text)
    if ($lines.Count -eq 0) { [void](Show-ExsMessage -Text 'List at least one person, one per line.' -Title $Title -Icon 'Warning'); return $null }
    $good = @(); $bad = @()
    foreach ($l in $lines) {
        $p = Resolve-ExsPerson -Text $l
        if ($p.Ok) { $good += $p } else { $bad += "$l  --  $($p.Reason)" }
    }
    if ($bad.Count) {
        Show-ExsResultList -Title "$Title -- these did not resolve, nothing was changed" -Lines $bad
        return $null
    }
    return @($good | Sort-Object Smtp -Unique)
}

function Invoke-ExsMailboxAdd {
    param([switch]$SendAs)
    $ed = $script:Ed
    $what = $(if ($SendAs) { 'Full Access + Send As' } else { 'Full Access' })
    $people = Resolve-ExsPeopleList -Text $ed.List.Text -Title "Add $what"
    if (-not $people) { return }
    $people = @($people | Where-Object { $_.Smtp -ine $ed.Smtp })
    if ($people.Count -eq 0) { Set-ExsEdStatus $ed 'The only entry was the mailbox itself.' -Bad; return }

    $a = Show-ExsMessage -Title "Add $what" -Icon 'Question' -Buttons 'OKCancel' -Text (
        "Give $what on`r`n  $($ed.Display) <$($ed.Smtp)>`r`nto:`r`n`r`n  " +
        (@($people | ForEach-Object { "$($_.Display)  <$($_.Smtp)>" }) -join "`r`n  ") +
        "`r`n`r`nAutomapping: on")
    if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $ed.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $out = @(); $ok = 0; $bad = 0
    try {
        [void](Get-ExsLiveRecipient -Identity $ed.Smtp)
        foreach ($p in $people) {
            Set-ExsEdStatus $ed "Adding $($p.Smtp)..."
            try {
                Add-ExsMailboxAccess -Smtp $ed.Smtp -User $p.Smtp -SendAs:$SendAs
                $out += "OK      $($p.Smtp)  ($what)"; $ok++
                Write-ExsLog -Action 'MailboxAccessAdd' -Target $ed.Smtp -Detail "$what for $($p.Smtp); automapping on"
            } catch {
                $m = "$($_.Exception.Message)"
                if ($m -match '(?i)already (has|exists|present)|existing permission') {
                    $out += "SKIP    $($p.Smtp)  already has it"; $ok++
                    Write-ExsLog -Action 'MailboxAccessAdd' -Target $ed.Smtp -Detail "$what for $($p.Smtp): already present" -Result 'Skipped'
                } else {
                    $out += "FAILED  $($p.Smtp)  $m"; $bad++
                    Write-ExsLog -Action 'MailboxAccessAdd' -Target $ed.Smtp -Detail "$what for $($p.Smtp): $m" -Result 'Failed'
                }
            }
        }
    } catch {
        $out += "STOPPED  $($_.Exception.Message)"; $bad++
    } finally { $ed.Form.Cursor = [System.Windows.Forms.Cursors]::Default }
    Show-ExsResultList -Title "Add $what -- $ok done, $bad failed" -Lines $out
    if ($bad -eq 0) { $ed.List.Text = '' }
    Update-ExsMailboxGrid
}

function Invoke-ExsMailboxRemove {
    $ed = $script:Ed
    $sel = @($ed.Grid.SelectedRows | ForEach-Object { $_.Tag } | Where-Object { $_ })
    if ($sel.Count -eq 0) { Set-ExsEdStatus $ed 'Select one or more people in the list first.' -Bad; return }
    $a = Show-ExsMessage -Title 'Remove access' -Icon 'Warning' -Buttons 'OKCancel' -Text (
        "Remove access to`r`n  $($ed.Display) <$($ed.Smtp)>`r`nfrom:`r`n`r`n  " +
        (@($sel | ForEach-Object {
                    $k = @(); if ($_.FullAccess) { $k += 'Full Access' }; if ($_.SendAs) { $k += 'Send As' }
                    "$($_.User)  ($($k -join ' + '))" }) -join "`r`n  "))
    if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $ed.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $out = @(); $ok = 0; $bad = 0
    try {
        [void](Get-ExsLiveRecipient -Identity $ed.Smtp)
        foreach ($x in $sel) {
            Set-ExsEdStatus $ed "Removing $($x.User)..."
            $k = @(); if ($x.FullAccess) { $k += 'Full Access' }; if ($x.SendAs) { $k += 'Send As' }
            try {
                Remove-ExsMailboxAccess -Smtp $ed.Smtp -User $x.User -FullAccess $x.FullAccess -SendAs $x.SendAs
                $out += "OK      $($x.User)  ($($k -join ' + ') removed)"; $ok++
                Write-ExsLog -Action 'MailboxAccessRemove' -Target $ed.Smtp -Detail "$($k -join ' + ') removed from $($x.User)"
            } catch {
                $out += "FAILED  $($x.User)  $($_.Exception.Message)"; $bad++
                Write-ExsLog -Action 'MailboxAccessRemove' -Target $ed.Smtp -Detail "$($x.User): $($_.Exception.Message)" -Result 'Failed'
            }
        }
    } catch {
        $out += "STOPPED  $($_.Exception.Message)"; $bad++
    } finally { $ed.Form.Cursor = [System.Windows.Forms.Cursors]::Default }
    Show-ExsResultList -Title "Remove access -- $ok done, $bad failed" -Lines $out
    Update-ExsMailboxGrid
}
#endregion

#region Group Editor ------------------------------------------------------------

function Show-ExsGroupEditor {
    param([Parameter(Mandatory)][string]$Identity, $Owner)
    $r = Get-ExsLiveRecipient -Identity $Identity
    $script:Ge = @{ Smtp = "$($r.PrimarySmtpAddress)"; Display = "$($r.DisplayName)" }
    $ge = $script:Ge

    $owners = ''
    try {
        $dg = Get-DistributionGroup -Identity $ge.Smtp -ErrorAction Stop
        $owners = (@($dg.ManagedBy | ForEach-Object { "$_" }) -join ', ')
    } catch { $owners = '(could not be read)' }

    $f = New-Object System.Windows.Forms.Form
    $f.Text = "Group Editor  --  $($ge.Display)"
    $f.ClientSize = New-Object System.Drawing.Size(900, 590)
    $f.StartPosition = 'CenterParent'
    $f.FormBorderStyle = 'FixedDialog'; $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $ge.Form = $f

    [void](New-ExsLabel $f "$($ge.Display)   <$($ge.Smtp)>" 12 12 876 20 -Bold)
    [void](New-ExsLabel $f ("$($script:ExsTypeNames["$($r.RecipientTypeDetails)"])   |   $($script:Tenant.Attribute) = $("$($r.($script:Tenant.Attribute))".Trim())   |   Owners: " +
        $(if ($owners) { $owners } else { '(none)' })) 12 34 876 20 -Muted)

    [void](New-ExsLabel $f 'Members (select one or more rows to remove):' 12 62 520 20 -Bold)
    $ge.Grid = New-ExsGrid $f @(@('Name', 200), @('Address', 210), @('Type', 100)) 12 84 520 410
    $ge.BtnRemove  = New-ExsButton $f 'Remove selected' 12 502 170 32
    $ge.BtnRefresh = New-ExsButton $f 'Refresh' 190 502 100 32

    [void](New-ExsLabel $f 'Add members -- one per line (name, alias or address):' 548 62 340 20 -Bold)
    $ge.List = New-ExsListBox $f 548 84 340 260
    [void](New-ExsLabel $f ("A bare alias gets @$($script:Tenant.PrimaryDomain). Every line is checked before anything " +
        'is changed; one bad line stops the whole add.') 548 350 340 52 -Muted)
    $ge.BtnAdd = New-ExsButton $f 'Add members' 548 410 340 34

    $ge.Status = New-ExsLabel $f '' 12 552 770 20 -Muted
    $close = New-ExsButton $f 'Close' 798 546 90 32
    $close.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $f.CancelButton = $close

    $ge.BtnRefresh.Add_Click({ Update-ExsGroupGrid })
    $ge.BtnAdd.Add_Click({ Invoke-ExsGroupAdd })
    $ge.BtnRemove.Add_Click({ Invoke-ExsGroupRemove })
    $f.Add_Shown({ Update-ExsGroupGrid })

    [void]$f.ShowDialog($Owner)
}

function Update-ExsGroupGrid {
    $ge = $script:Ge
    $ge.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    Set-ExsEdStatus $ge 'Reading members...'
    try {
        $rows = @(Get-ExsGroupMembers -Group $ge.Smtp)
        $ge.Grid.Rows.Clear()
        foreach ($x in $rows) {
            $i = $ge.Grid.Rows.Add($x.Name, $x.Address, $x.Type)
            $ge.Grid.Rows[$i].Tag = $x
        }
        $ge.Grid.ClearSelection()
        Set-ExsEdStatus $ge "$($rows.Count) member(s)."
    } catch {
        Set-ExsEdStatus $ge "Could not read members: $($_.Exception.Message)" -Bad
    } finally { $ge.Form.Cursor = [System.Windows.Forms.Cursors]::Default }
}

function Invoke-ExsGroupAdd {
    $ge = $script:Ge
    $people = Resolve-ExsPeopleList -Text $ge.List.Text -Title 'Add members'
    if (-not $people) { return }
    $people = @($people | Where-Object { $_.Smtp -ine $ge.Smtp })
    if ($people.Count -eq 0) { Set-ExsEdStatus $ge 'A group cannot be a member of itself.' -Bad; return }

    $a = Show-ExsMessage -Title 'Add members' -Icon 'Question' -Buttons 'OKCancel' -Text (
        "Add to`r`n  $($ge.Display) <$($ge.Smtp)>`r`n`r`n  " +
        (@($people | ForEach-Object { "$($_.Display)  <$($_.Smtp)>" }) -join "`r`n  "))
    if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $ge.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $out = @(); $ok = 0; $bad = 0
    try {
        [void](Get-ExsLiveRecipient -Identity $ge.Smtp)
        foreach ($p in $people) {
            Set-ExsEdStatus $ge "Adding $($p.Smtp)..."
            try {
                Set-ExsGroupMember -Verb Add -Group $ge.Smtp -Member $p.Smtp
                $out += "OK      $($p.Smtp)"; $ok++
                Write-ExsLog -Action 'GroupMemberAdd' -Target $ge.Smtp -Detail "added $($p.Smtp)"
            } catch {
                $m = "$($_.Exception.Message)"
                if ($m -match '(?i)already a member') {
                    $out += "SKIP    $($p.Smtp)  already a member"; $ok++
                    Write-ExsLog -Action 'GroupMemberAdd' -Target $ge.Smtp -Detail "$($p.Smtp): already a member" -Result 'Skipped'
                } else {
                    $out += "FAILED  $($p.Smtp)  $m"; $bad++
                    Write-ExsLog -Action 'GroupMemberAdd' -Target $ge.Smtp -Detail "$($p.Smtp): $m" -Result 'Failed'
                }
            }
        }
    } catch {
        $out += "STOPPED  $($_.Exception.Message)"; $bad++
    } finally { $ge.Form.Cursor = [System.Windows.Forms.Cursors]::Default }
    Show-ExsResultList -Title "Add members -- $ok done, $bad failed" -Lines $out
    if ($bad -eq 0) { $ge.List.Text = '' }
    Update-ExsGroupGrid
}

function Invoke-ExsGroupRemove {
    $ge = $script:Ge
    $sel = @($ge.Grid.SelectedRows | ForEach-Object { $_.Tag } | Where-Object { $_ })
    if ($sel.Count -eq 0) { Set-ExsEdStatus $ge 'Select one or more members in the list first.' -Bad; return }
    $a = Show-ExsMessage -Title 'Remove members' -Icon 'Warning' -Buttons 'OKCancel' -Text (
        "Remove from`r`n  $($ge.Display) <$($ge.Smtp)>`r`n`r`n  " +
        (@($sel | ForEach-Object { "$($_.Name)  <$($_.Address)>" }) -join "`r`n  "))
    if ($a -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $ge.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $out = @(); $ok = 0; $bad = 0
    try {
        [void](Get-ExsLiveRecipient -Identity $ge.Smtp)
        foreach ($x in $sel) {
            Set-ExsEdStatus $ge "Removing $($x.Name)..."
            try {
                Set-ExsGroupMember -Verb Remove -Group $ge.Smtp -Member $x.Id
                $out += "OK      $($x.Name)  <$($x.Address)>"; $ok++
                Write-ExsLog -Action 'GroupMemberRemove' -Target $ge.Smtp -Detail "removed $($x.Name) <$($x.Address)>"
            } catch {
                $out += "FAILED  $($x.Name)  $($_.Exception.Message)"; $bad++
                Write-ExsLog -Action 'GroupMemberRemove' -Target $ge.Smtp -Detail "$($x.Name): $($_.Exception.Message)" -Result 'Failed'
            }
        }
    } catch {
        $out += "STOPPED  $($_.Exception.Message)"; $bad++
    } finally { $ge.Form.Cursor = [System.Windows.Forms.Cursors]::Default }
    Show-ExsResultList -Title "Remove members -- $ok done, $bad failed" -Lines $out
    Update-ExsGroupGrid
}
#endregion
