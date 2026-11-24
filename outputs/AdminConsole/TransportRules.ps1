#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# ---------------------------------------------------------------------------
<#
    TransportRules.ps1
    Read the organisation's mail flow rules, and take a dated snapshot.

    READ ONLY, deliberately. Authoring stays in the admin centre -- see the
    help block for why.

    Author   : Glenn Armbruster
    Revised  : 2026-08-12
#>
# TOOLKIT-HELP-BEGIN
# Every mail flow (transport) rule in the organisation, in one searchable list.
#
# READ ONLY. Nothing here changes a rule, enables one, or reorders anything.
#
# WHY NOT EDIT HERE. A transport rule applies to every message the organisation
# sends or receives, and Set-TransportRule has well over a hundred conditions
# and actions. A form covering all of them would be the admin centre rebuilt
# worse; a form covering some of them can only edit the rules it understands,
# and would quietly drop the parts it does not. The admin centre's condition
# builder is the one thing it genuinely does well. Author there.
#
# WHAT THIS DOES BETTER:
#
#   Search       one box, matched against the WHOLE rule -- name, description,
#                every condition and every action. "Which rule mentions this
#                domain" is one keystroke here and a long afternoon in EAC.
#   Priority     the running order in a single list. Which rule wins is decided
#                by this, and it is the thing EAC shows least clearly.
#   Detail       every property the rule actually SETS, with the couple of
#                hundred empty ones left out.
#   Snapshot     the whole collection to a dated XML file. Mail flow has no
#                change history, so a snapshot before anyone edits is the only
#                way to answer "what did this look like last week".
#
# TAKE A SNAPSHOT BEFORE CHANGING RULES, not after. It is the only version of
# yesterday you will ever have.
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
#
# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost: the window flashes up, work may have started, and
# the operator has seen a tool they were not meant to be in.
if (-not (Test-ConsoleLaunch)) { return }

#region Data --------------------------------------------------------------------

# Properties that describe the rule itself rather than what it DOES. Held apart
# so the detail pane can lead with them and the search blob can skip them --
# matching 'Enabled' against every enabled rule is not a search result.
$script:MetaProps = @(
    'Name', 'Priority', 'State', 'Mode', 'Description', 'Comments',
    'WhenChanged', 'CreatedBy', 'LastModifiedBy', 'Guid', 'Identity',
    'DistinguishedName', 'ObjectClass', 'IsValid', 'ObjectState',
    'RuleVersion', 'ImmutableId', 'ExchangeVersion', 'OrganizationId',
    'RunspaceId', 'PSComputerName', 'PSShowComputerName'
)

function Get-RuleSetting {
    <#
        The properties a rule actually sets -- its conditions, exceptions and
        actions -- with the empty ones dropped.

        A transport rule object carries a couple of hundred properties and a
        typical rule populates six. Showing all of them is why reading a rule in
        raw PowerShell output is unpleasant; showing only what is set is the
        entire point of this window.
    #>
    param($Rule)

    $out = New-Object System.Collections.ArrayList
    foreach ($p in @($Rule.PSObject.Properties | Sort-Object Name)) {
        if ($script:MetaProps -contains $p.Name) { continue }

        $v = $p.Value
        if ($null -eq $v) { continue }

        # An empty collection is not a set value, and neither is an empty
        # string. Both are the overwhelming majority of a rule object.
        if ($v -is [System.Collections.IEnumerable] -and $v -isnot [string]) {
            $items = @($v)
            if ($items.Count -eq 0) { continue }
            $text = ($items -join ', ')
        } else {
            $text = "$v".Trim()
            if (-not $text) { continue }
            if ($text -eq 'False') { continue }
        }

        [void]$out.Add([PSCustomObject]@{ Name = $p.Name; Value = $text })
    }
    return $out
}

function Get-RuleRow {
    param($Rule)

    $settings = Get-RuleSetting -Rule $Rule

    # The search blob is built ONCE per rule, not per keystroke. Filtering runs
    # on every character typed, and re-flattening two hundred properties each
    # time is how a search box becomes laggy on a tenant with 300 rules.
    $blob = (@("$($Rule.Name)", "$($Rule.Description)", "$($Rule.Comments)") +
             @($settings | ForEach-Object { "$($_.Name) $($_.Value)" })) -join ' '

    return [PSCustomObject]@{
        Priority = $(if ($null -ne $Rule.Priority) { [int]$Rule.Priority } else { 9999 })
        Name     = "$($Rule.Name)"
        State    = "$($Rule.State)"
        Mode     = "$($Rule.Mode)"
        Summary  = (("$($Rule.Description)" -replace '\s+', ' ')).Trim()
        Settings = $settings
        Blob     = $blob
        Raw      = $Rule
    }
}

function Get-TransportRuleRow {
    # ResultSize defaults to 1000, and Exchange Online caps rules at 300, so
    # this changes nothing today. It is here because "the default happens to be
    # larger than the limit" is a fact about someone else's service that could
    # stop being true, and a truncated list of rules is a list that looks
    # complete. A read this cheap should not be quietly capped.
    return @(Invoke-EXOCommand -Name 'Get-TransportRule' -Parameters @{
                 ResultSize  = 'Unlimited'
                 ErrorAction = 'Stop'
             } |
             ForEach-Object { Get-RuleRow -Rule $_ } |
             Sort-Object Priority)
}

#endregion

#region Form --------------------------------------------------------------------

$ReportDir = Get-ReportPath

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Mail Flow Rules -- read only'
# ClientSize, not Size. Size includes the border and title bar, so setting it
# leaves you guessing how much room the controls actually have; ClientSize IS
# the drawable area and is what any later assignment would override anyway.
$form.ClientSize = New-Object System.Drawing.Size(1024, 738)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblTop = New-Object System.Windows.Forms.Label
$lblTop.Text = 'Every mail flow rule, in priority order. Read only -- author rules in the admin centre. Search matches the whole rule, not just its name.'
$lblTop.Location = New-Object System.Drawing.Point(12, 12)
$lblTop.Size = New-Object System.Drawing.Size(1000, 18)
$lblTop.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblTop)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = 'Refresh'
$btnRefresh.Location = New-Object System.Drawing.Point(12, 40)
$btnRefresh.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnRefresh)

$lblFind = New-Object System.Windows.Forms.Label
$lblFind.Text = 'Search:'
$lblFind.Location = New-Object System.Drawing.Point(124, 46)
$lblFind.Size = New-Object System.Drawing.Size(52, 18)
$form.Controls.Add($lblFind)

$txtFind = New-Object System.Windows.Forms.TextBox
$txtFind.Location = New-Object System.Drawing.Point(180, 43)
$txtFind.Size = New-Object System.Drawing.Size(300, 22)
$form.Controls.Add($txtFind)

$chkDisabled = New-Object System.Windows.Forms.CheckBox
$chkDisabled.Text = 'Disabled only'
$chkDisabled.Location = New-Object System.Drawing.Point(492, 44)
$chkDisabled.Size = New-Object System.Drawing.Size(110, 22)
$form.Controls.Add($chkDisabled)

$btnSnapshot = New-Object System.Windows.Forms.Button
$btnSnapshot.Text = 'Snapshot to XML'
$btnSnapshot.Location = New-Object System.Drawing.Point(612, 40)
$btnSnapshot.Size = New-Object System.Drawing.Size(130, 28)
$form.Controls.Add($btnSnapshot)

$btnReports = New-Object System.Windows.Forms.Button
$btnReports.Text = 'Reports folder'
$btnReports.Location = New-Object System.Drawing.Point(750, 40)
$btnReports.Size = New-Object System.Drawing.Size(110, 28)
$form.Controls.Add($btnReports)

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Text = 'Loading...'
$lblCount.Location = New-Object System.Drawing.Point(872, 46)
$lblCount.Size = New-Object System.Drawing.Size(140, 18)
$form.Controls.Add($lblCount)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = ''
$lblStatus.Location = New-Object System.Drawing.Point(12, 704)
$lblStatus.Size = New-Object System.Drawing.Size(880, 18)
$lblStatus.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblStatus)

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'TransportRules' `
             -X 12 -Y 78 -Width 1000 -StatusLabel $lblStatus -Extension 'xml'

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(12, 114)
$grid.Size = New-Object System.Drawing.Size(1000, 198)
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.SelectionMode = 'FullRowSelect'
$grid.MultiSelect = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.RowHeadersVisible = $false
[void]$grid.Columns.Add('Priority', 'Pri')
[void]$grid.Columns.Add('Name', 'Rule')
[void]$grid.Columns.Add('State', 'State')
[void]$grid.Columns.Add('Mode', 'Mode')
[void]$grid.Columns.Add('Summary', 'What it does')
$grid.Columns['Priority'].FillWeight = 6
$grid.Columns['Name'].FillWeight     = 26
$grid.Columns['State'].FillWeight    = 9
$grid.Columns['Mode'].FillWeight     = 10
$grid.Columns['Summary'].FillWeight  = 49
$form.Controls.Add($grid)

$lblDetail = New-Object System.Windows.Forms.Label
$lblDetail.Text = 'What the selected rule sets -- empty properties are left out:'
$lblDetail.Location = New-Object System.Drawing.Point(12, 316)
$lblDetail.Size = New-Object System.Drawing.Size(500, 18)
$form.Controls.Add($lblDetail)

$txtDetail = New-Object System.Windows.Forms.TextBox
$txtDetail.Location = New-Object System.Drawing.Point(12, 338)
$txtDetail.Size = New-Object System.Drawing.Size(1000, 300)
$txtDetail.Multiline = $true
$txtDetail.ReadOnly = $true
$txtDetail.ScrollBars = 'Both'
$txtDetail.WordWrap = $false
$txtDetail.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtDetail)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Location = New-Object System.Drawing.Point(12, 648)
$txtOut.Size = New-Object System.Drawing.Size(1000, 40)
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(912, 694)
$btnClose.Size = New-Object System.Drawing.Size(100, 28)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

#endregion

#region Behaviour ---------------------------------------------------------------

$script:Rules = @()

$renderGrid = {
    $needle = "$($txtFind.Text)".Trim()
    $show = @($script:Rules | Where-Object {
        (-not $chkDisabled.Checked -or "$($_.State)" -notmatch '(?i)^enabled') -and
        (-not $needle -or "$($_.Blob)" -like "*$needle*")
    })

    $grid.Rows.Clear()
    foreach ($r in $show) {
        $i = $grid.Rows.Add()
        $grid.Rows[$i].Cells['Priority'].Value = $r.Priority
        $grid.Rows[$i].Cells['Name'].Value     = $r.Name
        $grid.Rows[$i].Cells['State'].Value    = $r.State
        $grid.Rows[$i].Cells['Mode'].Value     = $r.Mode
        $grid.Rows[$i].Cells['Summary'].Value  = $r.Summary
        $grid.Rows[$i].Tag = $r

        # A disabled rule is greyed, and a rule in test mode is amber. Both look
        # active in EAC's list, and both mean "this is not doing what the name
        # says" -- which is exactly the thing you are hunting when a rule
        # mysteriously is not firing.
        if ("$($r.State)" -notmatch '(?i)^enabled') {
            $grid.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray
        } elseif ("$($r.Mode)" -notmatch '(?i)^enforce') {
            $grid.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(180, 110, 0)
        }
    }

    $lblCount.Text = $(if ($needle -or $chkDisabled.Checked) {
        "$($show.Count) of $($script:Rules.Count)"
    } else {
        "$($script:Rules.Count) rule(s)"
    })
}

$loadRules = {
    $form.Cursor = 'WaitCursor'
    $btnRefresh.Enabled = $false
    try {
        $script:Rules = @(Get-TransportRuleRow)
        & $renderGrid
        $txtDetail.Text = ''

        $off = @($script:Rules | Where-Object { "$($_.State)" -notmatch '(?i)^enabled' }).Count
        $test = @($script:Rules | Where-Object {
            "$($_.State)" -match '(?i)^enabled' -and "$($_.Mode)" -notmatch '(?i)^enforce' }).Count
        $notes = @()
        if ($off -gt 0)  { $notes += "$off disabled" }
        if ($test -gt 0) { $notes += "$test enabled but not enforcing" }
        $txtOut.Text = $(if ($notes.Count) {
            "$($script:Rules.Count) rule(s): " + ($notes -join ', ') + '. Those are shown greyed and amber.'
        } else {
            "$($script:Rules.Count) rule(s), all enabled and enforcing."
        })
    } catch {
        $lblCount.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Reading transport rules'
    } finally {
        $btnRefresh.Enabled = $true
        $form.Cursor = 'Default'
    }
}

$btnRefresh.Add_Click($loadRules)
$txtFind.Add_TextChanged($renderGrid)
$chkDisabled.Add_CheckedChanged($renderGrid)
$btnReports.Add_Click({ Open-ReportFolder -Path $ReportDir })

$grid.Add_SelectionChanged({
    if ($grid.SelectedRows.Count -eq 0) { $txtDetail.Text = ''; return }
    $r = $grid.SelectedRows[0].Tag
    if (-not $r) { return }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("$($r.Name)")
    [void]$sb.AppendLine(('-' * 78))
    [void]$sb.AppendLine(("{0,-34} {1}" -f 'Priority', $r.Priority))
    [void]$sb.AppendLine(("{0,-34} {1}" -f 'State', $r.State))
    [void]$sb.AppendLine(("{0,-34} {1}" -f 'Mode', $r.Mode))
    foreach ($n in @('WhenChanged', 'LastModifiedBy', 'CreatedBy', 'Comments')) {
        if ($r.Raw.PSObject.Properties[$n] -and "$($r.Raw.$n)".Trim()) {
            [void]$sb.AppendLine(("{0,-34} {1}" -f $n, "$($r.Raw.$n)".Trim()))
        }
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('CONDITIONS, EXCEPTIONS AND ACTIONS SET:')
    [void]$sb.AppendLine('')
    if (@($r.Settings).Count -eq 0) {
        [void]$sb.AppendLine('  (none -- this rule sets nothing, which usually means it is a stub)')
    } else {
        foreach ($s in $r.Settings) {
            [void]$sb.AppendLine(("  {0,-32} {1}" -f $s.Name, $s.Value))
        }
    }
    if ("$($r.Summary)".Trim()) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('AS THE SERVICE DESCRIBES IT:')
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("  $($r.Summary)")
    }
    $txtDetail.Text = $sb.ToString()
})

$btnSnapshot.Add_Click({
    $btnSnapshot.Enabled = $false
    $form.Cursor = 'WaitCursor'
    try {
        # Two steps, and that is the documented shape: the cmdlet returns an
        # object carrying the bytes, it does not write a file. Piping it to
        # Out-File produces a text rendering of the object rather than the XML.
        $file = Invoke-EXOCommand -Name 'Export-TransportRuleCollection' -Parameters @{
            ErrorAction = 'Stop'
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $path  = Join-Path $ReportDir "TransportRules_$stamp.xml"
        [System.IO.File]::WriteAllBytes($path, $file.FileData)

        $files.SelectPath($path)
        $txtOut.Text = ("Snapshot written." + [Environment]::NewLine +
                        "  $path" + [Environment]::NewLine +
                        "This is the whole collection and can be re-imported with " +
                        "Import-TransportRuleCollection -- which REPLACES every rule, so treat it " +
                        "as a restore of last resort rather than a merge.")
        Write-CompletedLog -Action 'Snapshot' -Target 'TransportRules' -Result 'Success' `
            -Detail "$($script:Rules.Count) rule(s) to $path"
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Exporting the rule collection'
        Write-FailedLog -Action 'Snapshot' -Target 'TransportRules' -ErrorRecord $_ `
            -Detail 'Export-TransportRuleCollection failed'
    } finally {
        $btnSnapshot.Enabled = $true
        $form.Cursor = 'Default'
    }
})

$form.Add_Shown({
    & $loadRules
    if ("$SmtpAddress".Trim()) { $txtFind.Text = "$SmtpAddress".Trim() }
})

Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Mail Flow Rules' | Out-Null
Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'

[void]$form.ShowDialog()
$form.Dispose()
#endregion
