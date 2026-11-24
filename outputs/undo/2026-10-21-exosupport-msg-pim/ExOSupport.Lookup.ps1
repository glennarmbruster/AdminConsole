#Requires -Version 5.1
<#
    ExOSupport.Lookup.ps1  --  Cloud Lookup, built INTO the main window.

    The lookup is the console's main job, so it is not a separate window: the
    search box has the focus when the console opens, Enter searches, and the
    two editor buttons sit under the results.

    Type a name, an address or part of either; the grid lists the cloud-managed
    shared mailboxes and distribution groups the operator is responsible for.
    Nothing else is ever returned: no user or resource mailboxes, no security,
    Microsoft 365 or dynamic groups, and nothing still synced from on-premises.

    Controls live in $script:Lk so the event handlers reach them however
    PowerShell scopes the handler.
#>

# BUILD STAMP. Every ExOSupport file carries one; ExOSupport.ps1 refuses to start
# unless all of them match its own, so a half-copied folder fails with a
# sentence instead of a missing-function error. Bump all four together.
if (-not $script:ExsParts) { $script:ExsParts = @{} }
$script:ExsParts['Lookup'] = '2026.10.20.4'

function Add-ExsLookupPanel {
    <#
        Builds the lookup controls on $Form between $Top and $Bottom (pixels
        from the bottom edge kept clear for the status strip). Everything
        anchors so the grid grows with the window.
    #>
    param([Parameter(Mandatory)]$Form, [int]$Top, [int]$Bottom = 30)
    $script:Lk = @{ Form = $Form }
    $lk = $script:Lk
    $w = $Form.ClientSize.Width
    $h = $Form.ClientSize.Height

    $cap = New-ExsLabel $Form 'Cloud Lookup' 16 $Top 300 26 -Bold
    $cap.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)

    $lk.Txt = New-Object System.Windows.Forms.TextBox
    $lk.Txt.Location = New-Object System.Drawing.Point(16, ($Top + 32))
    $lk.Txt.Size = New-Object System.Drawing.Size(($w - 160), 28)
    $lk.Txt.Font = New-Object System.Drawing.Font('Segoe UI', 12)
    $lk.Txt.Anchor = 'Top,Left,Right'
    $Form.Controls.Add($lk.Txt)
    $lk.BtnSearch = New-ExsButton $Form 'Search' ($w - 134) ($Top + 31) 118 32
    $lk.BtnSearch.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $lk.BtnSearch.Anchor = 'Top,Right'
    $Form.AcceptButton = $lk.BtnSearch

    $hint = New-ExsLabel $Form ("Name, address or part of either.  * on its own = everything.  'sales*' = starts with.  'name@' adds @$($script:Tenant.PrimaryDomain)." +
        '  Shared mailboxes and distribution groups managed in the cloud.') 16 ($Top + 66) ($w - 32) 18 -Muted
    $hint.Anchor = 'Top,Left,Right'
    $lk.Vis = New-ExsLabel $Form (Get-ExsVisibilityText) 16 ($Top + 86) ($w - 32) 18
    $lk.Vis.Anchor = 'Top,Left,Right'

    $gridTop = $Top + 110
    $gridH = $h - $gridTop - $Bottom - 84
    $lk.Grid = New-ExsGrid $Form @(@('Display name', 300), @('Address', 320), @('Type', 140), @("Team ($($script:Tenant.Attribute))", 120)) 16 $gridTop ($w - 32) $gridH
    $lk.Grid.MultiSelect = $false
    $lk.Grid.Anchor = 'Top,Bottom,Left,Right'

    $lk.Count = New-ExsLabel $Form '' 16 ($gridTop + $gridH + 6) ($w - 32) 18 -Muted
    $lk.Count.Anchor = 'Bottom,Left,Right'

    $btnY = $gridTop + $gridH + 30
    $lk.BtnMbx = New-ExsButton $Form 'Shared Mailbox Editor' 16 $btnY 200 36
    $lk.BtnGrp = New-ExsButton $Form 'Group Editor' 224 $btnY 160 36
    foreach ($b in @($lk.BtnMbx, $lk.BtnGrp)) { $b.Anchor = 'Bottom,Left'; $b.Enabled = $false }

    $lk.BtnSearch.Add_Click({ Invoke-ExsLookupSearch })
    $lk.Grid.Add_SelectionChanged({ Update-ExsLookupButtons })
    $lk.Grid.Add_CellDoubleClick({
            param($s, $e)
            if ($e.RowIndex -lt 0) { return }
            $row = Get-ExsLookupSelected
            if ($row) { Open-ExsEditorFor -Row $row }
        })
    $lk.BtnMbx.Add_Click({ $row = Get-ExsLookupSelected; if ($row -and $row.Kind -eq 'Mailbox') { Open-ExsEditorFor -Row $row } })
    $lk.BtnGrp.Add_Click({ $row = Get-ExsLookupSelected; if ($row -and $row.Kind -eq 'Group') { Open-ExsEditorFor -Row $row } })
    return $btnY
}

function Get-ExsLookupSelected {
    $g = $script:Lk.Grid
    if ($g.SelectedRows.Count -lt 1) { return $null }
    return $g.SelectedRows[0].Tag
}

function Update-ExsLookupButtons {
    $row = Get-ExsLookupSelected
    $script:Lk.BtnMbx.Enabled = [bool]($row -and $row.Kind -eq 'Mailbox')
    $script:Lk.BtnGrp.Enabled = [bool]($row -and $row.Kind -eq 'Group')
}

function Invoke-ExsLookupSearch {
    $lk = $script:Lk
    $term = "$($lk.Txt.Text)".Trim()
    if ($term -notmatch '\*' -and $term.Length -lt 2) {
        $lk.Count.ForeColor = [System.Drawing.Color]::Firebrick
        $lk.Count.Text = 'Type at least two characters (or use * on purpose).'
        return
    }
    if (-not (Get-ExsConnection).Connected) {
        $lk.Count.ForeColor = [System.Drawing.Color]::Firebrick
        $lk.Count.Text = 'Not connected to Exchange Online. Press Connect ExO at the top first.'
        return
    }
    $script:LastSearch = $term
    $lk.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $lk.Count.ForeColor = [System.Drawing.Color]::DimGray
    $lk.Count.Text = $(if ($term -match '^\*+$') { 'Listing everything you can see -- this can take a little while...' } else { 'Searching...' })
    $lk.Form.Refresh()
    try {
        $res = Find-ExsRecipient -Text $term
        $lk.Grid.Rows.Clear()
        foreach ($r in $res.Rows) {
            $i = $lk.Grid.Rows.Add($r.DisplayName, $r.Address, $r.Type, $r.Team)
            $lk.Grid.Rows[$i].Tag = $r
        }
        $lk.Grid.ClearSelection()
        $lk.Count.Text = "$(@($res.Rows).Count) found$(if ($res.Note) { "   ($($res.Note))" })"
        $lk.Vis.Text = Get-ExsVisibilityText
    } catch {
        $lk.Count.ForeColor = [System.Drawing.Color]::Firebrick
        $lk.Count.Text = "Search failed: $($_.Exception.Message)"
    } finally {
        $lk.Form.Cursor = [System.Windows.Forms.Cursors]::Default
        Update-ExsLookupButtons
    }
}

function Open-ExsEditorFor {
    param([Parameter(Mandatory)]$Row)
    try {
        if ($Row.Kind -eq 'Mailbox') { Show-ExsMailboxEditor -Identity $Row.Address -Owner $script:Lk.Form }
        else { Show-ExsGroupEditor -Identity $Row.Address -Owner $script:Lk.Form }
    } catch {
        [void](Show-ExsMessage -Text "Could not open the editor: $($_.Exception.Message)" -Icon 'Error')
    }
}
