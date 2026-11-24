#Requires -Version 5.1
<#
    DynamicDLPreview.ps1
    Resolves a dynamic distribution group to its members right now, and answers "is this person in it".

    Author  : Glenn Armbruster
    Revised : 2026-09-12
#>
# TOOLKIT-HELP-RUNS: Exchange Online only. Reads dynamic distribution groups and recipients. Changes nothing.
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# Shows who a dynamic distribution group resolves to RIGHT NOW, and whether a
# given person is in that list.
#
# A dynamic group has no member list. Exchange evaluates its recipient filter
# every time a message is sent, so "why did / didn't I get that email" cannot
# be answered from the group object alone. This tool runs the same evaluation
# Exchange does -- Get-Recipient with the group's RecipientFilter and
# RecipientContainer -- and puts the result in a grid.
#
# Load groups, pick one, Preview members. Type an address in the "Is this
# person in it?" box and Check: the answer is YES or NO, matched on the
# person's primary SMTP address against the preview.
#
# Display is capped at 5000 rows; the CSV export and the person check use the
# full result. Read-only. Never changes a group, a filter or a recipient.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# FIRST THING AFTER Common LOADS. See Test-ConsoleLaunch in Common.ps1.
if (-not (Test-ConsoleLaunch)) { return }

#region Configuration --------------------------------------------------------
$ReportDir    = Get-ReportPath
$ReportPrefix = 'DynamicDLPreview'
$GridRowCap   = 5000

# File-scope state. Kept in $script: at file scope only -- never read from
# anything Common.ps1 owns.
$script:AllGroups    = @()      # every dynamic DL, sorted by DisplayName
$script:GroupByItem  = @{}      # combo item text -> group object
$script:CurrentGroup = $null    # the group the facts panel is showing
$script:PreviewRows  = @()      # full preview result for CurrentGroup
$script:PreviewSmtp  = $null    # HashSet of PrimarySmtpAddress in the preview
$script:PreviewFor   = ''       # identity the preview was run for

# Theme colours for the YES / NO answer. Falls back to plain green / red when
# the theme does not define them, so the answer is always readable.
$script:ColorYes = $null
$script:ColorNo  = $null
try {
    $theme = Get-ConsoleTheme
    if ($theme) {
        $script:ColorYes = ConvertTo-ThemeColor "$($theme.SelectedText)"
        $script:ColorNo  = ConvertTo-ThemeColor "$($theme.DangerColor)"
    }
} catch { }
if (-not $script:ColorYes) { $script:ColorYes = [System.Drawing.Color]::FromArgb(22, 160, 114) }
if (-not $script:ColorNo)  { $script:ColorNo  = [System.Drawing.Color]::FromArgb(214, 69, 69) }
#endregion

#region Data -----------------------------------------------------------------
function Assert-ExoCmdlet {
    param([string]$Name)
    if (-not (Get-Command -Name $Name -ErrorAction SilentlyContinue)) {
        throw "$Name is not available. Connect to Exchange Online first."
    }
}

function Get-DynamicGroupList {
    param([scriptblock]$Progress)

    Assert-ExoCmdlet -Name 'Get-DynamicDistributionGroup'
    & $Progress 'Reading dynamic distribution groups...' 10

    # ResultSize defaults to 1000; omitting it silently truncates the tenant.
    $params = @{ ErrorAction = 'Stop' }
    $declared = Get-CmdletParameterNames -Name 'Get-DynamicDistributionGroup'
    if ($declared -contains 'ResultSize') { $params['ResultSize'] = 'Unlimited' }
    # Cloud-only switch that fills ManagedByWithDisplayNames; without it
    # ManagedBy is a list of GUIDs on recent module builds.
    if ($declared -contains 'IncludeManagedByWithDisplayNames') { $params['IncludeManagedByWithDisplayNames'] = $true }

    $groups = @(Invoke-EXOCommand -Name 'Get-DynamicDistributionGroup' -Parameters $params)
    return @($groups | Sort-Object { "$($_.DisplayName)" })
}

function Get-DynamicGroupPreview {
    param($Group, [scriptblock]$Progress)

    Assert-ExoCmdlet -Name 'Get-Recipient'

    $filter = "$($Group.RecipientFilter)".Trim()
    if (-not $filter) {
        throw "'$($Group.DisplayName)' has an empty RecipientFilter. Nothing to evaluate."
    }

    & $Progress "Evaluating the recipient filter for '$($Group.DisplayName)'... (one call; large tenants take a while)" 20

    # THIS IS THE SAME EVALUATION EXCHANGE DOES when a message is sent to the
    # group: the filter, scoped to the group's RecipientContainer. Without the
    # container the preview can include recipients the group never delivers to.
    $params = @{
        RecipientPreviewFilter = $filter
        ErrorAction            = 'Stop'
    }
    $declared = Get-CmdletParameterNames -Name 'Get-Recipient'
    if ($declared -contains 'ResultSize') { $params['ResultSize'] = 'Unlimited' }
    $container = "$($Group.RecipientContainer)".Trim()
    if ($container -and ($declared -contains 'OrganizationalUnit')) { $params['OrganizationalUnit'] = $container }

    $recipients = @(Invoke-EXOCommand -Name 'Get-Recipient' -Parameters $params)

    & $Progress "Shaping $($recipients.Count) recipient(s)..." 80

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $recipients) {
        $rows.Add([PSCustomObject][ordered]@{
            DisplayName          = "$($r.DisplayName)"
            PrimarySmtpAddress   = "$($r.PrimarySmtpAddress)"
            RecipientTypeDetails = "$($r.RecipientTypeDetails)"
            Department           = "$($r.Department)"
            Title                = "$($r.Title)"
            Office               = "$($r.Office)"
        })
    }

    return @($rows | Sort-Object DisplayName)
}

function Resolve-PersonSmtp {
    param([string]$Text)

    Assert-ExoCmdlet -Name 'Get-Recipient'
    $r = @(Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{ Identity = $Text; ErrorAction = 'Stop' })
    if ($r.Count -eq 0) { throw "No recipient matched '$Text'." }
    if ($r.Count -gt 1) {
        throw "'$Text' matched $($r.Count) recipients. Use the full email address."
    }
    return $r[0]
}

function Format-ManagedBy {
    param($Group)
    # Prefer the display-name form the cloud switch fills; fall back to
    # whatever ManagedBy holds (names on-prem, GUIDs on newer cloud builds).
    $names = @()
    if ($Group.PSObject.Properties['ManagedByWithDisplayNames']) {
        $names = @($Group.ManagedByWithDisplayNames | ForEach-Object { "$_" } | Where-Object { $_ })
    }
    if ($names.Count -eq 0) {
        $names = @($Group.ManagedBy | ForEach-Object { "$_" } | Where-Object { $_ })
    }
    if ($names.Count -eq 0) { return '(none)' }
    return ($names -join '; ')
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Dynamic DL Preview -- who does this group resolve to right now?'
$form.Size            = New-Object System.Drawing.Size(900, 730)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Evaluates a dynamic distribution group the way Exchange does when mail is sent to it. Read-only.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(860, 18)
$form.Controls.Add($lblInfo)

# --- Group picker row ---
$lblFilter          = New-Object System.Windows.Forms.Label
$lblFilter.Text     = 'Filter:'
$lblFilter.Location = New-Object System.Drawing.Point(12, 42)
$lblFilter.Size     = New-Object System.Drawing.Size(44, 18)
$form.Controls.Add($lblFilter)

$txtGroupFilter          = New-Object System.Windows.Forms.TextBox
$txtGroupFilter.Location = New-Object System.Drawing.Point(58, 39)
$txtGroupFilter.Size     = New-Object System.Drawing.Size(180, 24)
$form.Controls.Add($txtGroupFilter)

$cboGroups                   = New-Object System.Windows.Forms.ComboBox
$cboGroups.DropDownStyle     = 'DropDownList'
$cboGroups.Location          = New-Object System.Drawing.Point(250, 39)
$cboGroups.Size              = New-Object System.Drawing.Size(500, 24)
$cboGroups.Enabled           = $false
$form.Controls.Add($cboGroups)

$btnLoad          = New-Object System.Windows.Forms.Button
$btnLoad.Text     = 'Load groups'
$btnLoad.Location = New-Object System.Drawing.Point(762, 36)
$btnLoad.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnLoad)

# --- Facts panel ---
function New-FactLabel {
    param([string]$Text, [int]$Y)
    $l = New-Object System.Windows.Forms.Label
    $l.Text      = $Text
    $l.Location  = New-Object System.Drawing.Point(12, $Y)
    $l.Size      = New-Object System.Drawing.Size(96, 18)
    $l.ForeColor = [System.Drawing.Color]::DimGray
    return $l
}
function New-FactValue {
    param([int]$Y)
    $l = New-Object System.Windows.Forms.Label
    $l.Text     = ''
    $l.Location = New-Object System.Drawing.Point(110, $Y)
    $l.Size     = New-Object System.Drawing.Size(762, 18)
    $l.AutoEllipsis = $true
    return $l
}

$form.Controls.Add((New-FactLabel -Text 'Address:'    -Y 78))
$form.Controls.Add((New-FactLabel -Text 'Container:'  -Y 98))
$form.Controls.Add((New-FactLabel -Text 'Managed by:' -Y 118))
$form.Controls.Add((New-FactLabel -Text 'Changed:'    -Y 138))
$form.Controls.Add((New-FactLabel -Text 'Filter:'     -Y 160))

$lblAddress   = New-FactValue -Y 78;  $form.Controls.Add($lblAddress)
$lblContainer = New-FactValue -Y 98;  $form.Controls.Add($lblContainer)
$lblManagedBy = New-FactValue -Y 118; $form.Controls.Add($lblManagedBy)
$lblChanged   = New-FactValue -Y 138; $form.Controls.Add($lblChanged)

$txtRecipientFilter            = New-Object System.Windows.Forms.TextBox
$txtRecipientFilter.Multiline  = $true
$txtRecipientFilter.ReadOnly   = $true
$txtRecipientFilter.ScrollBars = 'Vertical'
$txtRecipientFilter.WordWrap   = $true
$txtRecipientFilter.Location   = New-Object System.Drawing.Point(110, 160)
$txtRecipientFilter.Size       = New-Object System.Drawing.Size(762, 64)
$txtRecipientFilter.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtRecipientFilter)

# --- Action row ---
$btnPreview          = New-Object System.Windows.Forms.Button
$btnPreview.Text     = 'Preview members'
$btnPreview.Location = New-Object System.Drawing.Point(12, 234)
$btnPreview.Size     = New-Object System.Drawing.Size(130, 30)
$btnPreview.Enabled  = $false
$form.Controls.Add($btnPreview)

$btnExport          = New-Object System.Windows.Forms.Button
$btnExport.Text     = 'Export CSV'
$btnExport.Location = New-Object System.Drawing.Point(150, 234)
$btnExport.Size     = New-Object System.Drawing.Size(110, 30)
$btnExport.Enabled  = $false
$form.Controls.Add($btnExport)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Reports Folder'
$btnOpen.Location = New-Object System.Drawing.Point(268, 234)
$btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnOpen)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(762, 234)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle. Click Load groups to begin.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 272)
$lblStatus.Size     = New-Object System.Drawing.Size(860, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 292 -Width 860 -Height 12

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix $ReportPrefix `
                -X 12 -Y 310 -Width 860 -StatusLabel $lblStatus

# --- Result grid ---
$lblCount          = New-Object System.Windows.Forms.Label
$lblCount.Text     = 'No preview yet.'
$lblCount.Location = New-Object System.Drawing.Point(12, 348)
$lblCount.Size     = New-Object System.Drawing.Size(860, 18)
$form.Controls.Add($lblCount)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location              = New-Object System.Drawing.Point(12, 368)
$grid.Size                  = New-Object System.Drawing.Size(860, 232)
$grid.ReadOnly              = $true
$grid.AllowUserToAddRows    = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.SelectionMode         = 'FullRowSelect'
$grid.MultiSelect           = $false
$grid.AutoSizeColumnsMode   = 'Fill'
$grid.RowHeadersVisible     = $false
[void]$grid.Columns.Add('DisplayName',          'Display name')
[void]$grid.Columns.Add('PrimarySmtpAddress',   'Primary SMTP')
[void]$grid.Columns.Add('RecipientTypeDetails', 'Type')
[void]$grid.Columns.Add('Department',           'Department')
[void]$grid.Columns.Add('Title',                'Title')
[void]$grid.Columns.Add('Office',               'Office')
$grid.Columns['DisplayName'].FillWeight          = 20
$grid.Columns['PrimarySmtpAddress'].FillWeight   = 24
$grid.Columns['RecipientTypeDetails'].FillWeight = 14
$grid.Columns['Department'].FillWeight           = 14
$grid.Columns['Title'].FillWeight                = 14
$grid.Columns['Office'].FillWeight               = 14
$form.Controls.Add($grid)

# --- Person check row ---
$lblPerson          = New-Object System.Windows.Forms.Label
$lblPerson.Text     = 'Is this person in it?'
$lblPerson.Location = New-Object System.Drawing.Point(12, 616)
$lblPerson.Size     = New-Object System.Drawing.Size(130, 18)
$form.Controls.Add($lblPerson)

$txtPerson          = New-Object System.Windows.Forms.TextBox
$txtPerson.Location = New-Object System.Drawing.Point(146, 613)
$txtPerson.Size     = New-Object System.Drawing.Size(380, 24)
$txtPerson.Text     = "$SmtpAddress"
$form.Controls.Add($txtPerson)

$btnCheck          = New-Object System.Windows.Forms.Button
$btnCheck.Text     = 'Check'
$btnCheck.Location = New-Object System.Drawing.Point(534, 610)
$btnCheck.Size     = New-Object System.Drawing.Size(110, 30)
$btnCheck.Enabled  = $false
$form.Controls.Add($btnCheck)

$lblAnswer           = New-Object System.Windows.Forms.Label
$lblAnswer.Text      = ''
$lblAnswer.Location  = New-Object System.Drawing.Point(652, 606)
$lblAnswer.Size      = New-Object System.Drawing.Size(220, 36)
$lblAnswer.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 18, [System.Drawing.FontStyle]::Bold)
$lblAnswer.TextAlign = 'MiddleLeft'
$form.Controls.Add($lblAnswer)

$lblAnswerDetail          = New-Object System.Windows.Forms.Label
$lblAnswerDetail.Text     = ''
$lblAnswerDetail.Location = New-Object System.Drawing.Point(12, 648)
$lblAnswerDetail.Size     = New-Object System.Drawing.Size(860, 36)
$form.Controls.Add($lblAnswerDetail)
#endregion

#region UI helpers -----------------------------------------------------------
function Set-GroupCombo {
    # Rebuilds the combo from $script:AllGroups narrowed by the filter box.
    # Item text is the DisplayName; a duplicate DisplayName gets its address
    # appended so two groups never collapse into one item.
    param([string]$Filter)

    $cboGroups.BeginUpdate()
    try {
        $cboGroups.Items.Clear()
        $script:GroupByItem = @{}
        $seen = @{}
        foreach ($g in $script:AllGroups) {
            $name = "$($g.DisplayName)"
            if (-not $name) { $name = "$($g.Name)" }
            if ($Filter -and ($name -notlike "*$Filter*") -and ("$($g.PrimarySmtpAddress)" -notlike "*$Filter*")) { continue }
            $item = $name
            if ($seen.ContainsKey($item.ToLowerInvariant())) { $item = "$name  [$($g.PrimarySmtpAddress)]" }
            $seen[$item.ToLowerInvariant()] = $true
            $script:GroupByItem[$item] = $g
            [void]$cboGroups.Items.Add($item)
        }
    } finally {
        $cboGroups.EndUpdate()
    }
    $cboGroups.Enabled = ($cboGroups.Items.Count -gt 0)
    if ($cboGroups.Items.Count -eq 1) { $cboGroups.SelectedIndex = 0 }
}

function Clear-Preview {
    $script:PreviewRows = @()
    $script:PreviewSmtp = $null
    $script:PreviewFor  = ''
    $grid.Rows.Clear()
    $lblCount.Text          = 'No preview yet.'
    $lblAnswer.Text         = ''
    $lblAnswerDetail.Text   = ''
    $btnExport.Enabled      = $false
}

function Show-GroupFacts {
    param($Group)
    $script:CurrentGroup = $Group
    Clear-Preview
    if (-not $Group) {
        $lblAddress.Text = ''; $lblContainer.Text = ''; $lblManagedBy.Text = ''; $lblChanged.Text = ''
        $txtRecipientFilter.Text = ''
        $btnPreview.Enabled = $false
        $btnCheck.Enabled   = $false
        return
    }
    $lblAddress.Text   = "$($Group.PrimarySmtpAddress)"
    $lblContainer.Text = "$($Group.RecipientContainer)"
    $lblManagedBy.Text = Format-ManagedBy -Group $Group
    $lblChanged.Text   = if ($Group.WhenChanged) { "$($Group.WhenChanged)" } else { '' }
    $txtRecipientFilter.Text = "$($Group.RecipientFilter)"
    $btnPreview.Enabled = $true
    $btnCheck.Enabled   = $true
    $lblStatus.Text = "Selected '$($Group.DisplayName)'. Preview members to evaluate the filter."
}

function Invoke-Preview {
    # Runs the preview for the selected group, fills the grid, logs once.
    # Shared by the Preview button and by Check when no preview exists yet.
    $g = $script:CurrentGroup
    if (-not $g) { throw 'Select a dynamic distribution group first.' }

    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $btnPreview.Enabled = $false
    $btnCheck.Enabled   = $false
    try {
        Clear-Preview
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity 'Dynamic DL preview'
        $rows = @(Get-DynamicGroupPreview -Group $g -Progress $progress)
        Complete-ProgressReporter -Activity 'Dynamic DL preview' -Bar $pbar

        $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($r in $rows) { if ($r.PrimarySmtpAddress) { [void]$set.Add($r.PrimarySmtpAddress) } }

        $script:PreviewRows = $rows
        $script:PreviewSmtp = $set
        $script:PreviewFor  = "$($g.Identity)"

        $shown = [Math]::Min($rows.Count, $GridRowCap)
        $grid.SuspendLayout()
        try {
            for ($i = 0; $i -lt $shown; $i++) {
                $r = $rows[$i]
                [void]$grid.Rows.Add(@($r.DisplayName, $r.PrimarySmtpAddress, $r.RecipientTypeDetails, $r.Department, $r.Title, $r.Office))
            }
        } finally {
            $grid.ResumeLayout()
        }

        if ($rows.Count -gt $GridRowCap) {
            $lblCount.Text = "$($rows.Count) member(s) resolved. Showing the first $GridRowCap -- the CSV export and the person check use all $($rows.Count)."
        } else {
            $lblCount.Text = "$($rows.Count) member(s) resolved right now."
        }
        $lblStatus.Text    = "Done. '$($g.DisplayName)' resolves to $($rows.Count) recipient(s)."
        $btnExport.Enabled = ($rows.Count -gt 0)

        Write-CompletedLog -Task "Dynamic DL preview -- $($g.DisplayName) -- $($rows.Count) members" -Result 'Success'
    } finally {
        Complete-ProgressReporter -Activity 'Dynamic DL preview'
        $form.Cursor        = $prev
        $btnPreview.Enabled = $true
        $btnCheck.Enabled   = $true
    }
}
#endregion

#region Event handlers -------------------------------------------------------
$btnOpen.Add_Click({
    try { Open-ReportFolder -Path $ReportDir }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Open reports folder' }
})

$btnLoad.Add_Click({
    $btnLoad.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Show-GroupFacts -Group $null
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity 'Dynamic DL list'
        $script:AllGroups = @(Get-DynamicGroupList -Progress $progress)
        Complete-ProgressReporter -Activity 'Dynamic DL list' -Bar $pbar

        Set-GroupCombo -Filter $txtGroupFilter.Text
        if ($script:AllGroups.Count -eq 0) {
            $lblStatus.Text = 'No dynamic distribution groups were returned.'
        } else {
            $lblStatus.Text = "$($script:AllGroups.Count) dynamic distribution group(s) loaded. Pick one."
        }
    } catch {
        $lblStatus.Text = 'Failed to load groups.'
        Show-HandlerError -ErrorRecord $_ -Context 'Load dynamic distribution groups'
    } finally {
        Complete-ProgressReporter -Activity 'Dynamic DL list'
        $form.Cursor     = $prev
        $btnLoad.Enabled = $true
    }
})

$txtGroupFilter.Add_TextChanged({
    try {
        if ($script:AllGroups.Count -eq 0) { return }
        $keep = $script:CurrentGroup
        Set-GroupCombo -Filter $txtGroupFilter.Text
        # Keep the current selection if it survived the filter; otherwise
        # the facts panel would show a group the combo no longer lists.
        $idx = -1
        if ($keep) {
            foreach ($k in $script:GroupByItem.Keys) {
                if ("$($script:GroupByItem[$k].Identity)" -eq "$($keep.Identity)") { $idx = $cboGroups.Items.IndexOf($k); break }
            }
        }
        if ($idx -ge 0) { $cboGroups.SelectedIndex = $idx }
        elseif ($cboGroups.Items.Count -ne 1) { Show-GroupFacts -Group $null }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Filter group list' }
})

$cboGroups.Add_SelectedIndexChanged({
    try {
        $item = "$($cboGroups.SelectedItem)"
        if (-not $item -or -not $script:GroupByItem.ContainsKey($item)) { return }
        $g = $script:GroupByItem[$item]
        if ($script:CurrentGroup -and ("$($script:CurrentGroup.Identity)" -eq "$($g.Identity)")) { return }
        Show-GroupFacts -Group $g
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Select group' }
})

$btnPreview.Add_Click({
    try {
        Invoke-Preview
    } catch {
        $lblStatus.Text = 'Preview failed.'
        Show-HandlerError -ErrorRecord $_ -Context 'Preview dynamic DL members'
    }
})

$btnCheck.Add_Click({
    $btnCheck.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $g = $script:CurrentGroup
        if (-not $g) { throw 'Select a dynamic distribution group first.' }
        $who = "$($txtPerson.Text)".Trim()
        if (-not $who) { throw 'Type an email address (or name) to check.' }

        $lblAnswer.Text       = ''
        $lblAnswerDetail.Text = ''

        # The answer is only as good as the preview it is matched against, so
        # a stale or missing preview is re-run rather than trusted.
        if (-not $script:PreviewSmtp -or $script:PreviewFor -ne "$($g.Identity)") { Invoke-Preview }

        $lblStatus.Text = "Resolving '$who'..."
        $lblStatus.Refresh()
        $person = Resolve-PersonSmtp -Text $who
        $smtp   = "$($person.PrimarySmtpAddress)"
        if (-not $smtp) { throw "'$who' resolved to '$($person.DisplayName)' but it has no primary SMTP address." }

        $inGroup = $script:PreviewSmtp.Contains($smtp)
        if ($inGroup) {
            $lblAnswer.ForeColor  = $script:ColorYes
            $lblAnswer.Text       = 'YES'
            $lblAnswerDetail.Text = "$($person.DisplayName) <$smtp> ($($person.RecipientTypeDetails)) IS in '$($g.DisplayName)' right now -- the filter matches them."
        } else {
            $lblAnswer.ForeColor  = $script:ColorNo
            $lblAnswer.Text       = 'NO'
            $lblAnswerDetail.Text = "$($person.DisplayName) <$smtp> ($($person.RecipientTypeDetails)) is NOT in '$($g.DisplayName)' right now -- the filter does not match them, or they are outside the group's container."
        }
        $lblStatus.Text = "Checked '$smtp' against $($script:PreviewRows.Count) resolved member(s)."
    } catch {
        $lblStatus.Text = 'Check failed.'
        Show-HandlerError -ErrorRecord $_ -Context 'Check person against dynamic DL'
    } finally {
        $form.Cursor      = $prev
        $btnCheck.Enabled = ($null -ne $script:CurrentGroup)
    }
})

$btnExport.Add_Click({
    try {
        $g = $script:CurrentGroup
        $rows = @($script:PreviewRows)
        if (-not $g -or $rows.Count -eq 0) { throw 'Nothing to export. Preview members first.' }

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        # Alias in the filename so a folder of these is readable; scrubbed to
        # characters every filesystem accepts.
        $tag   = ("$($g.Alias)" -replace '[^A-Za-z0-9_.-]', '_')
        if (-not $tag) { $tag = 'group' }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir ("{0}_{1}_{2}.csv" -f $ReportPrefix, $tag, $stamp)
        $rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $files.SelectPath($csv)
        $lblStatus.Text = "Exported $($rows.Count) row(s) to $csv"
    } catch {
        $lblStatus.Text = 'Export failed.'
        Show-HandlerError -ErrorRecord $_ -Context 'Export dynamic DL preview'
    }
})
#endregion

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Dynamic DL Preview' | Out-Null

[void]$form.ShowDialog()
