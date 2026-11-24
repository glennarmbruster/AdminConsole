#Requires -Version 5.1
<#
    FreeBusyViewer.ps1
    Everything that decides whether free/busy crosses an organisation boundary, in one window, plus a test.

    Author   : Glenn Armbruster
    Revised  : 2026-09-12
#>
# TOOLKIT-HELP-RUNS: reads organization relationships, intra-organization connectors and configuration, and sharing policies from Exchange Online. "Test" runs Test-OrganizationRelationship for one user. Changes nothing.
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# "Why can't they see my calendar" lives in four places, and this window
# shows all four:
#
#   * Organization relationships -- federation with another Exchange
#     organisation, and the hybrid relationship with on-premises. Domains,
#     whether free/busy is on, at what level, and the endpoints it targets.
#   * Intra-organization connectors -- the hybrid OAuth path, which takes
#     over from the organization relationship for the on-premises side when
#     it is present and enabled.
#   * Intra-organization configuration -- what Exchange Online believes the
#     on-premises target addresses and endpoints are.
#   * Sharing policies -- per-user calendar sharing to outside domains,
#     which is a different mechanism from the three above and is the one
#     that governs "share my calendar with a partner".
#
# Load reads all four. Select a row to see every property of it in the box
# under the grids. Export writes the visible tab to CSV.
#
# The Test tab runs Test-OrganizationRelationship for one relationship and
# one user mailbox and shows every line the cmdlet produces, verbose and
# warning output included -- that is where the useful part of the answer
# is. It reads and verifies; it changes nothing.
#
# Read-only throughout.
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
$ReportPrefix = 'FreeBusyViewer'

# Grid column order per tab. Blank where a build does not expose a column.
$RelColumns = @(
    'Name', 'Enabled', 'DomainNames', 'FreeBusyAccessEnabled', 'FreeBusyAccessLevel', 'FreeBusyAccessScope',
    'TargetApplicationUri', 'TargetAutodiscoverEpr', 'TargetSharingEpr',
    'MailboxMoveEnabled', 'MailTipsAccessEnabled', 'MailTipsAccessLevel', 'ArchiveAccessEnabled',
    'PhotosEnabled', 'OrganizationContact', 'WhenChanged'
)
$RelWidths = @{
    Name = 180; Enabled = 60; DomainNames = 220; FreeBusyAccessEnabled = 80; FreeBusyAccessLevel = 110; FreeBusyAccessScope = 140
    TargetApplicationUri = 200; TargetAutodiscoverEpr = 260; TargetSharingEpr = 260
    MailboxMoveEnabled = 80; MailTipsAccessEnabled = 80; MailTipsAccessLevel = 100; ArchiveAccessEnabled = 80
    PhotosEnabled = 70; OrganizationContact = 180; WhenChanged = 120
}
$ConColumns = @('Name', 'Enabled', 'TargetAddressDomains', 'DiscoveryEndpoint', 'TargetSharingEpr', 'WhenChanged')
$ConWidths  = @{ Name = 200; Enabled = 60; TargetAddressDomains = 260; DiscoveryEndpoint = 300; TargetSharingEpr = 300; WhenChanged = 120 }
$PolColumns = @('Name', 'Enabled', 'Default', 'Domains', 'WhenChanged')
$PolWidths  = @{ Name = 220; Enabled = 60; Default = 60; Domains = 500; WhenChanged = 120 }

# File-scope state. Kept in $script: at file scope only -- never read from
# anything Common.ps1 owns.
$script:Data = $null      # the shaped result of the last Load
#endregion

#region Data -----------------------------------------------------------------
function Format-PropertyText {
    <#
        Format-List by hand: every Name = Value line of a shaped object.
        Built from the object so the text lands in a TextBox, not the host.
    #>
    param($Object)
    if ($null -eq $Object) { return '' }
    $sb = New-Object System.Text.StringBuilder
    $width = 0
    foreach ($p in $Object.PSObject.Properties) { if ($p.Name.Length -gt $width) { $width = $p.Name.Length } }
    foreach ($p in $Object.PSObject.Properties) {
        [void]$sb.AppendLine(("{0,-$width} = {1}" -f $p.Name, "$($p.Value)"))
    }
    return $sb.ToString()
}

function Get-RowValue {
    param($Row, [string]$Name)
    $p = $Row.PSObject.Properties[$Name]
    if ($null -eq $p) { return '' }
    return "$($p.Value)"
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Free/Busy Viewer -- organization relationships, connectors, sharing policies'
$form.Size            = New-Object System.Drawing.Size(1100, 800)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'The four places that decide whether free/busy crosses an organisation boundary. Read-only.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(1060, 18)
$form.Controls.Add($lblInfo)

$btnLoad          = New-Object System.Windows.Forms.Button
$btnLoad.Text     = 'Load'
$btnLoad.Location = New-Object System.Drawing.Point(12, 38)
$btnLoad.Size     = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnLoad)

$btnExport          = New-Object System.Windows.Forms.Button
$btnExport.Text     = 'Export CSV'
$btnExport.Location = New-Object System.Drawing.Point(140, 38)
$btnExport.Size     = New-Object System.Drawing.Size(110, 30)
$btnExport.Enabled  = $false
$form.Controls.Add($btnExport)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Reports Folder'
$btnOpen.Location = New-Object System.Drawing.Point(258, 38)
$btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnOpen)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(962, 38)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle. Click Load to begin.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 76)
$lblStatus.Size     = New-Object System.Drawing.Size(1060, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 96 -Width 1060 -Height 12

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix $ReportPrefix `
                -X 12 -Y 114 -Width 1060 -StatusLabel $lblStatus

function New-ReadGrid {
    param([string[]]$Columns, [hashtable]$Widths)
    $g                        = New-Object System.Windows.Forms.DataGridView
    $g.Dock                   = 'Fill'
    $g.ReadOnly               = $true
    $g.AllowUserToAddRows     = $false
    $g.AllowUserToDeleteRows  = $false
    $g.AllowUserToResizeRows  = $false
    $g.SelectionMode          = 'FullRowSelect'
    $g.MultiSelect            = $false
    $g.RowHeadersVisible      = $false
    $g.AutoSizeColumnsMode    = 'None'
    $g.ScrollBars             = 'Both'
    foreach ($col in $Columns) {
        $idx = $g.Columns.Add($col, $col)
        $g.Columns[$idx].Width    = $Widths[$col]
        $g.Columns[$idx].SortMode = 'Automatic'
    }
    return $g
}

$tabs          = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(12, 152)
$tabs.Size     = New-Object System.Drawing.Size(1060, 360)
$form.Controls.Add($tabs)

$tabRel = New-Object System.Windows.Forms.TabPage
$tabRel.Text = 'Organization relationships'
$gridRel = New-ReadGrid -Columns $RelColumns -Widths $RelWidths
$tabRel.Controls.Add($gridRel)
$tabs.TabPages.Add($tabRel)

$tabCon = New-Object System.Windows.Forms.TabPage
$tabCon.Text = 'Intra-org connectors'
$gridCon = New-ReadGrid -Columns $ConColumns -Widths $ConWidths
$tabCon.Controls.Add($gridCon)
$tabs.TabPages.Add($tabCon)

$tabCfg = New-Object System.Windows.Forms.TabPage
$tabCfg.Text = 'Intra-org configuration'
$txtCfg            = New-Object System.Windows.Forms.TextBox
$txtCfg.Multiline  = $true
$txtCfg.ReadOnly   = $true
$txtCfg.ScrollBars = 'Both'
$txtCfg.WordWrap   = $false
$txtCfg.Dock       = 'Fill'
$txtCfg.Font       = New-Object System.Drawing.Font('Consolas', 9)
$tabCfg.Controls.Add($txtCfg)
$tabs.TabPages.Add($tabCfg)

$tabPol = New-Object System.Windows.Forms.TabPage
$tabPol.Text = 'Sharing policies'
$gridPol = New-ReadGrid -Columns $PolColumns -Widths $PolWidths
$tabPol.Controls.Add($gridPol)
$tabs.TabPages.Add($tabPol)

# --- Test tab ---
$tabTest = New-Object System.Windows.Forms.TabPage
$tabTest.Text = 'Test'

$lblTestRel          = New-Object System.Windows.Forms.Label
$lblTestRel.Text     = 'Organization relationship:'
$lblTestRel.Location = New-Object System.Drawing.Point(10, 14)
$lblTestRel.Size     = New-Object System.Drawing.Size(160, 18)
$tabTest.Controls.Add($lblTestRel)

$cboTestRel               = New-Object System.Windows.Forms.ComboBox
$cboTestRel.DropDownStyle = 'DropDownList'
$cboTestRel.Location      = New-Object System.Drawing.Point(176, 11)
$cboTestRel.Size          = New-Object System.Drawing.Size(360, 24)
$cboTestRel.Enabled       = $false
$tabTest.Controls.Add($cboTestRel)

$lblTestUser          = New-Object System.Windows.Forms.Label
$lblTestUser.Text     = 'User mailbox:'
$lblTestUser.Location = New-Object System.Drawing.Point(10, 46)
$lblTestUser.Size     = New-Object System.Drawing.Size(160, 18)
$tabTest.Controls.Add($lblTestUser)

$txtTestUser          = New-Object System.Windows.Forms.TextBox
$txtTestUser.Location = New-Object System.Drawing.Point(176, 43)
$txtTestUser.Size     = New-Object System.Drawing.Size(360, 24)
$txtTestUser.Text     = "$SmtpAddress"
$tabTest.Controls.Add($txtTestUser)

$btnTest          = New-Object System.Windows.Forms.Button
$btnTest.Text     = 'Run test'
$btnTest.Location = New-Object System.Drawing.Point(548, 40)
$btnTest.Size     = New-Object System.Drawing.Size(110, 30)
$btnTest.Enabled  = $false
$tabTest.Controls.Add($btnTest)

$lblTestNote           = New-Object System.Windows.Forms.Label
$lblTestNote.Text      = 'Runs Test-OrganizationRelationship -Identity <relationship> -UserIdentity <user>. Read-only; the useful part of the answer is usually in the verbose lines.'
$lblTestNote.Location  = New-Object System.Drawing.Point(10, 76)
$lblTestNote.Size      = New-Object System.Drawing.Size(1030, 18)
$lblTestNote.ForeColor = [System.Drawing.Color]::DimGray
$tabTest.Controls.Add($lblTestNote)

$txtTest            = New-Object System.Windows.Forms.TextBox
$txtTest.Multiline  = $true
$txtTest.ReadOnly   = $true
$txtTest.ScrollBars = 'Both'
$txtTest.WordWrap   = $false
$txtTest.Location   = New-Object System.Drawing.Point(10, 98)
$txtTest.Size       = New-Object System.Drawing.Size(1032, 226)
$txtTest.Font       = New-Object System.Drawing.Font('Consolas', 9)
$tabTest.Controls.Add($txtTest)
$tabs.TabPages.Add($tabTest)

$lblDetail          = New-Object System.Windows.Forms.Label
$lblDetail.Text     = 'Every property of the selected row:'
$lblDetail.Location = New-Object System.Drawing.Point(12, 520)
$lblDetail.Size     = New-Object System.Drawing.Size(1060, 18)
$form.Controls.Add($lblDetail)

$txtDetail            = New-Object System.Windows.Forms.TextBox
$txtDetail.Multiline  = $true
$txtDetail.ReadOnly   = $true
$txtDetail.ScrollBars = 'Both'
$txtDetail.WordWrap   = $false
$txtDetail.Location   = New-Object System.Drawing.Point(12, 540)
$txtDetail.Size       = New-Object System.Drawing.Size(1060, 200)
$txtDetail.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtDetail)
#endregion

#region UI helpers -----------------------------------------------------------
function Set-ReadGrid {
    <#
        Rows into a grid. Each grid row's Tag carries the shaped object so
        the detail box and the export read from the object, not the cells.
    #>
    param($Grid, [string[]]$Columns, $Rows)
    $Grid.SuspendLayout()
    try {
        $Grid.Rows.Clear()
        foreach ($r in @($Rows)) {
            $values = @()
            foreach ($col in $Columns) { $values += (Get-RowValue -Row $r -Name $col) }
            $idx = $Grid.Rows.Add($values)
            $Grid.Rows[$idx].Tag = $r
            if ((Get-RowValue -Row $r -Name 'Enabled') -eq 'False') {
                $Grid.Rows[$idx].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray
            }
        }
    } finally {
        $Grid.ResumeLayout()
    }
    $Grid.ClearSelection()
}

function Show-RowDetail {
    param($Grid)
    if ($Grid.SelectedRows.Count -eq 0) { return }
    $r = $Grid.SelectedRows[0].Tag
    if ($null -eq $r) { return }
    $txtDetail.Text = Format-PropertyText -Object $r
    $txtDetail.SelectionStart  = 0
    $txtDetail.SelectionLength = 0
    $txtDetail.ScrollToCaret()
}

function Get-ActiveExport {
    <#
        The rows and columns of the tab in front, or $null for the tabs
        that have no grid (configuration and test).
    #>
    $t = $tabs.SelectedTab
    if ($t -eq $tabRel) { return [PSCustomObject]@{ Name = 'OrganizationRelationships'; Grid = $gridRel; Columns = $RelColumns } }
    if ($t -eq $tabCon) { return [PSCustomObject]@{ Name = 'IntraOrgConnectors';        Grid = $gridCon; Columns = $ConColumns } }
    if ($t -eq $tabPol) { return [PSCustomObject]@{ Name = 'SharingPolicies';           Grid = $gridPol; Columns = $PolColumns } }
    return $null
}

function Update-ExportButton {
    $x = Get-ActiveExport
    $btnExport.Enabled = ($null -ne $x -and $x.Grid.Rows.Count -gt 0)
}
#endregion

#region Event handlers -------------------------------------------------------
$btnOpen.Add_Click({
    try { Open-ReportFolder -Path $ReportDir }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Open reports folder' }
})

$gridRel.Add_SelectionChanged({
    try { Show-RowDetail -Grid $gridRel }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Relationship detail' }
})
$gridCon.Add_SelectionChanged({
    try { Show-RowDetail -Grid $gridCon }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Connector detail' }
})
$gridPol.Add_SelectionChanged({
    try { Show-RowDetail -Grid $gridPol }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Sharing policy detail' }
})

$tabs.Add_SelectedIndexChanged({
    try { Update-ExportButton }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Switch tab' }
})

$btnLoad.Add_Click({
    try {
        $exo = Test-ExoConnectionState
        if (-not $exo.Connected) {
            Show-ToolDialog(
                "Not connected to Exchange Online. Use the Connect button on the console first.`r`n`r`n$($exo.Detail)",
                'Exchange Online required', 'OK', 'Warning') | Out-Null
            return
        }

        $btnLoad.Enabled   = $false
        $btnExport.Enabled = $false
        $btnTest.Enabled   = $false
        $script:Data = $null
        $gridRel.Rows.Clear(); $gridCon.Rows.Clear(); $gridPol.Rows.Clear()
        $txtCfg.Text = ''; $txtDetail.Text = ''
        $cboTestRel.Items.Clear(); $cboTestRel.Enabled = $false
        $lblStatus.Text = 'Reading free/busy configuration...'

        # OFF THE UI THREAD when the process allows it (Start-ToolJob in
        # Common.ps1 explains the test and the fallback). The block travels
        # to the worker as text: it shapes every object into a flat record
        # of strings itself, because the worker sees Common.ps1 and nothing
        # else in this file. Each read is its own try so one cmdlet that a
        # tenant does not have (intra-org configuration on a non-hybrid
        # tenant is the usual one) costs one message, not the whole load.
        Start-ToolJob -Form $form -Text 'Reading free/busy configuration ...' -Script {
            param()
            $toText = {
                param($Value)
                if ($null -eq $Value) { return '' }
                if ($Value -is [string]) { return $Value }
                if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm') }
                if ($Value -is [System.Collections.IEnumerable]) {
                    return (@($Value | ForEach-Object { "$_" } | Where-Object { $_ }) -join '; ')
                }
                return "$Value"
            }
            $shape = {
                param($Object)
                $o = [ordered]@{}
                foreach ($p in $Object.PSObject.Properties) {
                    if ($p.Name -in @('PSComputerName', 'RunspaceId', 'PSShowComputerName')) { continue }
                    $o[$p.Name] = (& $toText $p.Value)
                }
                return [PSCustomObject]$o
            }
            $read = {
                param([string]$Name)
                $res = [ordered]@{ Rows = @(); Error = '' }
                try {
                    $items = @(Invoke-EXOCommand -Name $Name -Parameters @{ ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' })
                    $res.Rows = @($items | ForEach-Object { & $shape $_ })
                } catch { $res.Error = "$($_.Exception.Message)" }
                return [PSCustomObject]$res
            }
            [PSCustomObject]@{
                Relationships = (& $read 'Get-OrganizationRelationship')
                Connectors    = (& $read 'Get-IntraOrganizationConnector')
                Configuration = (& $read 'Get-IntraOrganizationConfiguration')
                Policies      = (& $read 'Get-SharingPolicy')
            }
        } -OnDone {
            param($r, $st)
            try {
                if ($r.Failed) {
                    $lblStatus.Text = "Load failed: $($r.Error)"
                    return
                }
                $data = @($r.Output) | Select-Object -First 1
                $script:Data = $data

                Set-ReadGrid -Grid $gridRel -Columns $RelColumns -Rows $data.Relationships.Rows
                Set-ReadGrid -Grid $gridCon -Columns $ConColumns -Rows $data.Connectors.Rows
                Set-ReadGrid -Grid $gridPol -Columns $PolColumns -Rows $data.Policies.Rows

                $cfgRows = @($data.Configuration.Rows)
                if ($data.Configuration.Error) {
                    $txtCfg.Text = 'Get-IntraOrganizationConfiguration did not answer:' + [Environment]::NewLine +
                                   $data.Configuration.Error + [Environment]::NewLine + [Environment]::NewLine +
                                   'On a tenant with no hybrid deployment this is expected.'
                } elseif ($cfgRows.Count -eq 0) {
                    $txtCfg.Text = 'Get-IntraOrganizationConfiguration returned nothing.'
                } else {
                    $parts = @()
                    foreach ($c in $cfgRows) { $parts += (Format-PropertyText -Object $c) }
                    $txtCfg.Text = ($parts -join ([Environment]::NewLine + ('-' * 60) + [Environment]::NewLine))
                }

                $cboTestRel.Items.Clear()
                foreach ($rel in @($data.Relationships.Rows)) {
                    $n = Get-RowValue -Row $rel -Name 'Name'
                    if ($n) { [void]$cboTestRel.Items.Add($n) }
                }
                $cboTestRel.Enabled = ($cboTestRel.Items.Count -gt 0)
                if ($cboTestRel.Items.Count -gt 0) { $cboTestRel.SelectedIndex = 0 }
                $btnTest.Enabled = ($cboTestRel.Items.Count -gt 0)

                $errs = @()
                foreach ($k in @('Relationships', 'Connectors', 'Policies')) {
                    if ($data.$k.Error) { $errs += "${k}: $($data.$k.Error)" }
                }
                $lblStatus.Text = "Done. $(@($data.Relationships.Rows).Count) relationship(s), $(@($data.Connectors.Rows).Count) intra-org connector(s), $(@($data.Policies.Rows).Count) sharing polic$(if (@($data.Policies.Rows).Count -eq 1) { 'y' } else { 'ies' })."
                if ($errs.Count -gt 0) {
                    $lblStatus.Text += ' Some reads failed -- see below.'
                    $txtDetail.Text = ($errs -join [Environment]::NewLine)
                }
                Update-ExportButton
            } finally {
                $btnLoad.Enabled = $true
            }
        }
    } catch {
        $lblStatus.Text = 'Load failed.'
        $btnLoad.Enabled = $true
        Show-HandlerError -ErrorRecord $_ -Context 'Load free/busy configuration'
    }
})

$btnTest.Add_Click({
    try {
        $rel  = "$($cboTestRel.SelectedItem)".Trim()
        $user = "$($txtTestUser.Text)".Trim()
        if (-not $rel)  { $lblStatus.Text = 'Pick an organization relationship first.'; return }
        if (-not $user) { $lblStatus.Text = 'Type the user mailbox to test with.'; return }

        $exo = Test-ExoConnectionState
        if (-not $exo.Connected) {
            Show-ToolDialog(
                "Not connected to Exchange Online. Use the Connect button on the console first.`r`n`r`n$($exo.Detail)",
                'Exchange Online required', 'OK', 'Warning') | Out-Null
            return
        }

        $btnTest.Enabled = $false
        $txtTest.Text = "Test-OrganizationRelationship -Identity '$rel' -UserIdentity '$user'" + [Environment]::NewLine +
                        "Started $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" + [Environment]::NewLine + 'Working...'

        # OFF THE UI THREAD when the process allows it. The worker merges
        # the verbose, warning and error streams into the output so every
        # line the cmdlet says comes back as text -- the verdict is usually
        # in the verbose stream, and a test that shows only the objects
        # shows almost nothing.
        Start-ToolJob -Form $form -Text "Testing '$rel' for $user ..." -Script {
            param($Rel, $User)
            $cmd = Get-EXOCommand -Name 'Test-OrganizationRelationship'
            $p = @{ Identity = $Rel; UserIdentity = $User; Confirm = $false; ErrorAction = 'Continue'; Verbose = $true }
            $lines = @()
            $items = @(& $cmd @p 4>&1 3>&1 2>&1)
            foreach ($i in $items) {
                if ($i -is [System.Management.Automation.VerboseRecord])     { $lines += "VERBOSE: $($i.Message)"; continue }
                if ($i -is [System.Management.Automation.WarningRecord])     { $lines += "WARNING: $($i.Message)"; continue }
                if ($i -is [System.Management.Automation.ErrorRecord])       { $lines += "ERROR: $($i.Exception.Message)"; continue }
                if ($i -is [System.Management.Automation.InformationRecord]) { $lines += "$($i.MessageData)"; continue }
                if ($i -is [string]) { $lines += $i; continue }
                $parts = @()
                foreach ($prop in $i.PSObject.Properties) {
                    if ($prop.Name -in @('PSComputerName', 'RunspaceId', 'PSShowComputerName')) { continue }
                    $v = $prop.Value
                    if ($v -is [System.Collections.IEnumerable] -and -not ($v -is [string])) { $v = (@($v | ForEach-Object { "$_" }) -join '; ') }
                    $parts += "$($prop.Name) = $v"
                }
                $lines += ($parts -join [Environment]::NewLine)
                $lines += ''
            }
            if ($lines.Count -eq 0) { $lines += 'The cmdlet returned nothing on any stream.' }
            ,$lines
        } -ArgumentList $rel, $user -State @{ Rel = $rel; User = $user } -OnDone {
            param($r, $st)
            try {
                $head = "Test-OrganizationRelationship -Identity '$($st.Rel)' -UserIdentity '$($st.User)'" + [Environment]::NewLine +
                        "Finished $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" + [Environment]::NewLine + ('-' * 60) + [Environment]::NewLine
                if ($r.Failed) {
                    $txtTest.Text = $head + "FAILED: $($r.Error)"
                    $lblStatus.Text = 'Test failed.'
                    return
                }
                $lines = @()
                foreach ($o in @($r.Output)) { foreach ($l in @($o)) { $lines += "$l" } }
                $txtTest.Text = $head + ($lines -join [Environment]::NewLine)
                $txtTest.SelectionStart  = 0
                $txtTest.SelectionLength = 0
                $txtTest.ScrollToCaret()
                $lblStatus.Text = "Test finished for '$($st.Rel)' -- $($lines.Count) line(s)."
            } finally {
                $btnTest.Enabled = ($cboTestRel.Items.Count -gt 0)
            }
        }
    } catch {
        $lblStatus.Text = 'Test failed.'
        $btnTest.Enabled = ($cboTestRel.Items.Count -gt 0)
        Show-HandlerError -ErrorRecord $_ -Context 'Test organization relationship'
    }
})

$btnExport.Add_Click({
    try {
        $x = Get-ActiveExport
        if ($null -eq $x) { $lblStatus.Text = 'This tab has no grid to export. Pick a grid tab.'; return }
        if ($x.Grid.Rows.Count -eq 0) { $lblStatus.Text = 'Nothing to export. Load first.'; return }

        $export = New-Object System.Collections.Generic.List[object]
        foreach ($row in $x.Grid.Rows) {
            $r = $row.Tag
            if ($null -eq $r) { continue }
            $o = [ordered]@{}
            foreach ($col in $x.Columns) { $o[$col] = (Get-RowValue -Row $r -Name $col) }
            $export.Add([PSCustomObject]$o)
        }

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir ("{0}_{1}_{2}.csv" -f $ReportPrefix, $x.Name, $stamp)
        $export | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $files.SelectPath($csv)
        $lblStatus.Text = "Exported $($export.Count) row(s) to $csv"
        Write-CompletedLog -Action 'Report' -Target $csv -Result 'Success' `
            -Detail "Free/busy viewer: $($x.Name), $($export.Count) row(s)."
    } catch {
        $lblStatus.Text = 'Export failed.'
        Show-HandlerError -ErrorRecord $_ -Context 'Export free/busy grid'
    }
})
#endregion

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Free/Busy Viewer' | Out-Null

[void]$form.ShowDialog()
