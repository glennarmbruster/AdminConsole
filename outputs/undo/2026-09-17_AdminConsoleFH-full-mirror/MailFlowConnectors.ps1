#Requires -Version 5.1
<#
    MailFlowConnectors.ps1
    Read-only health view of Exchange Online inbound and outbound connectors. Author : Glenn Armbruster Revised : 2026-09-12

    Author   : Glenn Armbruster
    Revised  : 2026-09-12
#>
# TOOLKIT-HELP-RUNS: reads every inbound and outbound connector in Exchange Online. "Test outbound" sends ONE test message through the selected outbound connector, after asking.
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# Lists every Exchange Online inbound and outbound connector in one grid, with
# the settings hybrid mail flow depends on: type, enabled state, sender or
# recipient domains, smart hosts, TLS and certificate requirements.
#
# Connectors are the part of hybrid mail flow nobody looks at until mail
# stops. The Warnings box flags the handful of states that are unambiguously
# worth a look -- a disabled connector, an on-premises connector routing by
# MX record, an on-premises connector with no TLS or no certificate domain --
# and says why in one line each.
#
# Details shows every property of the selected connector. Export writes the
# grid to CSV. Read-only, with one exception: "Test outbound" runs
# Validate-OutboundConnector against the selected OUTBOUND connector, which
# sends a real test message to an address you type. It asks first and logs
# separately. Nothing is ever changed on a connector.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost.
if (-not (Test-ConsoleLaunch)) { return }

#region Configuration --------------------------------------------------------
$ReportDir = Get-ReportPath

# Grid column order. Blank where a column does not apply to that direction.
$GridColumns = @(
    'Direction', 'Name', 'Enabled', 'ConnectorType',
    'SenderDomains', 'TlsSenderCertificateName', 'RequireTls',
    'RestrictDomainsToCertificate', 'RestrictDomainsToIPAddresses', 'SenderIPAddresses',
    'RecipientDomains', 'SmartHosts', 'TlsSettings', 'TlsDomain', 'UseMXRecord',
    'IsTransportRuleScoped', 'WhenChanged'
)
$GridWidths = @{
    Direction = 70; Name = 200; Enabled = 60; ConnectorType = 95
    SenderDomains = 160; TlsSenderCertificateName = 150; RequireTls = 75
    RestrictDomainsToCertificate = 90; RestrictDomainsToIPAddresses = 90; SenderIPAddresses = 160
    RecipientDomains = 160; SmartHosts = 160; TlsSettings = 110; TlsDomain = 140; UseMXRecord = 85
    IsTransportRuleScoped = 90; WhenChanged = 120
}
#endregion

#region Data -----------------------------------------------------------------
function ConvertTo-CellText {
    <#
        One cell's worth of text from a connector property. Multi-valued
        properties (domains, smart hosts, IP ranges) join with '; ' so a CSV
        cell and a grid cell read the same; $null becomes ''.
    #>
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [System.Collections.IEnumerable]) {
        return (@($Value | ForEach-Object { "$_" } | Where-Object { $_ }) -join '; ')
    }
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm') }
    return "$Value"
}

function Get-ConnectorProperty {
    <#
        Reads a property by name without throwing when the module build in
        front of us does not expose it. Returns $null for missing.
    #>
    param($Object, [string]$Name)
    $p = $Object.PSObject.Properties[$Name]
    if ($null -eq $p) { return $null }
    return $p.Value
}

function Get-ConnectorReport {
    param([scriptblock]$Progress)

    foreach ($cmd in @('Get-InboundConnector', 'Get-OutboundConnector')) {
        if (-not (Get-Command -Name $cmd -ErrorAction SilentlyContinue)) {
            throw "$cmd is not available. Connect to Exchange Online first."
        }
    }

    & $Progress 'Reading inbound connectors...' 10
    # ResultSize defaults to 1000. No tenant has that many connectors, but the
    # rule is the rule: pass it whenever the cmdlet declares it.
    $pIn = @{ ErrorAction = 'Stop' }
    if ((Get-CmdletParameterNames -Name 'Get-InboundConnector') -contains 'ResultSize') { $pIn['ResultSize'] = 'Unlimited' }
    $inbound = @(Get-InboundConnector @pIn)

    & $Progress 'Reading outbound connectors...' 50
    $pOut = @{ ErrorAction = 'Stop' }
    if ((Get-CmdletParameterNames -Name 'Get-OutboundConnector') -contains 'ResultSize') { $pOut['ResultSize'] = 'Unlimited' }
    $outbound = @(Get-OutboundConnector @pOut)

    & $Progress 'Building rows...' 85

    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($c in $inbound) {
        $rows.Add([PSCustomObject][ordered]@{
            Direction                    = 'Inbound'
            Name                         = ConvertTo-CellText (Get-ConnectorProperty $c 'Name')
            Enabled                      = ConvertTo-CellText (Get-ConnectorProperty $c 'Enabled')
            ConnectorType                = ConvertTo-CellText (Get-ConnectorProperty $c 'ConnectorType')
            SenderDomains                = ConvertTo-CellText (Get-ConnectorProperty $c 'SenderDomains')
            TlsSenderCertificateName     = ConvertTo-CellText (Get-ConnectorProperty $c 'TlsSenderCertificateName')
            RequireTls                   = ConvertTo-CellText (Get-ConnectorProperty $c 'RequireTls')
            RestrictDomainsToCertificate = ConvertTo-CellText (Get-ConnectorProperty $c 'RestrictDomainsToCertificate')
            RestrictDomainsToIPAddresses = ConvertTo-CellText (Get-ConnectorProperty $c 'RestrictDomainsToIPAddresses')
            SenderIPAddresses            = ConvertTo-CellText (Get-ConnectorProperty $c 'SenderIPAddresses')
            RecipientDomains             = ''
            SmartHosts                   = ''
            TlsSettings                  = ''
            TlsDomain                    = ''
            UseMXRecord                  = ''
            IsTransportRuleScoped        = ''
            WhenChanged                  = ConvertTo-CellText (Get-ConnectorProperty $c 'WhenChanged')
            Raw                          = $c
        })
    }

    foreach ($c in $outbound) {
        $rows.Add([PSCustomObject][ordered]@{
            Direction                    = 'Outbound'
            Name                         = ConvertTo-CellText (Get-ConnectorProperty $c 'Name')
            Enabled                      = ConvertTo-CellText (Get-ConnectorProperty $c 'Enabled')
            ConnectorType                = ConvertTo-CellText (Get-ConnectorProperty $c 'ConnectorType')
            SenderDomains                = ''
            TlsSenderCertificateName     = ''
            RequireTls                   = ''
            RestrictDomainsToCertificate = ''
            RestrictDomainsToIPAddresses = ''
            SenderIPAddresses            = ''
            RecipientDomains             = ConvertTo-CellText (Get-ConnectorProperty $c 'RecipientDomains')
            SmartHosts                   = ConvertTo-CellText (Get-ConnectorProperty $c 'SmartHosts')
            TlsSettings                  = ConvertTo-CellText (Get-ConnectorProperty $c 'TlsSettings')
            TlsDomain                    = ConvertTo-CellText (Get-ConnectorProperty $c 'TlsDomain')
            UseMXRecord                  = ConvertTo-CellText (Get-ConnectorProperty $c 'UseMXRecord')
            IsTransportRuleScoped        = ConvertTo-CellText (Get-ConnectorProperty $c 'IsTransportRuleScoped')
            WhenChanged                  = ConvertTo-CellText (Get-ConnectorProperty $c 'WhenChanged')
            Raw                          = $c
        })
    }

    return [PSCustomObject]@{
        Rows          = @($rows)
        InboundCount  = @($inbound).Count
        OutboundCount = @($outbound).Count
    }
}

function Get-ConnectorWarnings {
    <#
        CONSERVATIVE ON PURPOSE. Each line here is a state that is worth a
        human's minute, not a style opinion. A connector that has not changed
        in a year is not a finding -- stable is what a working connector looks
        like.
    #>
    param($Rows)
    $out = New-Object System.Collections.Generic.List[string]

    foreach ($r in @($Rows)) {
        $label  = "$($r.Direction) '$($r.Name)'"
        $onPrem = ($r.ConnectorType -eq 'OnPremises')

        if ($r.Enabled -eq 'False') {
            $out.Add("$label -- DISABLED. Mail that matches this connector is not using it; if it was meant to be retired, remove it so it stops looking live.")
        }

        if ($r.Direction -eq 'Outbound' -and $onPrem -and $r.UseMXRecord -eq 'True') {
            $out.Add("$label -- OnPremises connector routes by MX record, not a smart host. Hybrid normally pins delivery to the on-premises host; MX for the recipient domains may point somewhere else.")
        }

        if ($r.Direction -eq 'Inbound' -and $onPrem -and $r.RequireTls -ne 'True') {
            $out.Add("$label -- OnPremises connector without RequireTls. Mail from on-premises is identified by its TLS certificate; without TLS that identification cannot happen. (Docs describe RequireTls for Partner connectors, so confirm before changing.)")
        }

        if ($onPrem) {
            if ($r.Direction -eq 'Inbound' -and -not "$($r.TlsSenderCertificateName)".Trim()) {
                $out.Add("$label -- OnPremises connector with no TlsSenderCertificateName. Nothing ties this connector to the on-premises certificate, so any sender matching its other conditions gets on-premises treatment.")
            }
            if ($r.Direction -eq 'Outbound' -and -not "$($r.TlsDomain)".Trim()) {
                $out.Add("$label -- OnPremises connector with no TlsDomain. Without a domain the connector cannot verify the on-premises certificate name (DomainValidation).")
            }
        }
    }

    return @($out)
}

function Format-ConnectorDetail {
    <#
        Format-List by hand: every property, Name = Value, one per line,
        multi-valued joined with '; '. Built from the object rather than
        Format-List so the text lands in a TextBox instead of the host.
    #>
    param($Object)
    if ($null -eq $Object) { return '' }
    $sb = New-Object System.Text.StringBuilder
    $names = @($Object.PSObject.Properties | ForEach-Object { $_.Name })
    $width = 0
    foreach ($n in $names) { if ($n.Length -gt $width) { $width = $n.Length } }
    foreach ($p in $Object.PSObject.Properties) {
        $v = ConvertTo-CellText $p.Value
        [void]$sb.AppendLine(("{0,-$width} = {1}" -f $p.Name, $v))
    }
    return $sb.ToString()
}

function Format-ValidationResult {
    param($Result, [string]$ConnectorName, [string]$Recipient)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Validate-OutboundConnector -- '$ConnectorName' -> $Recipient")
    [void]$sb.AppendLine("Run at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$sb.AppendLine('')
    $items = @($Result)
    if ($items.Count -eq 0) {
        [void]$sb.AppendLine('The cmdlet returned nothing. Check the recipient mailbox for the test message.')
        return $sb.ToString()
    }
    $i = 0
    foreach ($item in $items) {
        $i++
        if ($items.Count -gt 1) { [void]$sb.AppendLine("--- Result $i of $($items.Count) ---") }
        if ($item -is [string]) { [void]$sb.AppendLine($item) }
        else { [void]$sb.Append((Format-ConnectorDetail -Object $item)) }
        [void]$sb.AppendLine('')
    }
    [void]$sb.AppendLine('Note: this test does not set IsValidated or LastValidationTimestamp on the connector.')
    return $sb.ToString()
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Report -- Mail Flow Connectors'
$form.Size            = New-Object System.Drawing.Size(1100, 780)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Exchange Online inbound and outbound connectors -- the settings hybrid mail flow depends on, in one grid.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(1060, 18)
$form.Controls.Add($lblInfo)

$lblSrc           = New-Object System.Windows.Forms.Label
$lblSrc.Text      = 'Read-only. "Test outbound" is the one exception: it sends a single test message through the selected outbound connector, after asking.'
$lblSrc.Location  = New-Object System.Drawing.Point(12, 32)
$lblSrc.Size      = New-Object System.Drawing.Size(1060, 18)
$lblSrc.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblSrc)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run Report'
$btnRun.Location = New-Object System.Drawing.Point(12, 58)
$btnRun.Size     = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnRun)

$btnDetails          = New-Object System.Windows.Forms.Button
$btnDetails.Text     = 'Details'
$btnDetails.Location = New-Object System.Drawing.Point(140, 58)
$btnDetails.Size     = New-Object System.Drawing.Size(110, 30)
$btnDetails.Enabled  = $false
$form.Controls.Add($btnDetails)

$btnTest          = New-Object System.Windows.Forms.Button
$btnTest.Text     = 'Test outbound'
$btnTest.Location = New-Object System.Drawing.Point(258, 58)
$btnTest.Size     = New-Object System.Drawing.Size(130, 30)
$btnTest.Enabled  = $false
$form.Controls.Add($btnTest)

$btnExport          = New-Object System.Windows.Forms.Button
$btnExport.Text     = 'Export CSV'
$btnExport.Location = New-Object System.Drawing.Point(396, 58)
$btnExport.Size     = New-Object System.Drawing.Size(110, 30)
$btnExport.Enabled  = $false
$form.Controls.Add($btnExport)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Reports Folder'
$btnOpen.Location = New-Object System.Drawing.Point(514, 58)
$btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnOpen)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(962, 58)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 132)
$lblStatus.Size     = New-Object System.Drawing.Size(1060, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 152 -Width 1060 -Height 12

$grid                        = New-Object System.Windows.Forms.DataGridView
$grid.Location               = New-Object System.Drawing.Point(12, 170)
$grid.Size                   = New-Object System.Drawing.Size(1060, 270)
$grid.ReadOnly               = $true
$grid.AllowUserToAddRows     = $false
$grid.AllowUserToDeleteRows  = $false
$grid.AllowUserToResizeRows  = $false
$grid.SelectionMode          = 'FullRowSelect'
$grid.MultiSelect            = $false
$grid.RowHeadersVisible      = $false
$grid.AutoSizeColumnsMode    = 'None'
$grid.ScrollBars             = 'Both'
foreach ($col in $GridColumns) {
    $idx = $grid.Columns.Add($col, $col)
    $grid.Columns[$idx].Width = $GridWidths[$col]
    $grid.Columns[$idx].SortMode = 'Automatic'
}
$form.Controls.Add($grid)

$lblDetail          = New-Object System.Windows.Forms.Label
$lblDetail.Text     = 'Details of the selected connector (every property), or the result of "Test outbound":'
$lblDetail.Location = New-Object System.Drawing.Point(12, 448)
$lblDetail.Size     = New-Object System.Drawing.Size(600, 18)
$form.Controls.Add($lblDetail)

$txtDetail            = New-Object System.Windows.Forms.TextBox
$txtDetail.Multiline  = $true
$txtDetail.ReadOnly   = $true
$txtDetail.ScrollBars = 'Both'
$txtDetail.WordWrap   = $false
$txtDetail.Location   = New-Object System.Drawing.Point(12, 468)
$txtDetail.Size       = New-Object System.Drawing.Size(600, 240)
$txtDetail.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtDetail)

$lblWarn          = New-Object System.Windows.Forms.Label
$lblWarn.Text     = 'Warnings -- only states that are unambiguously worth a look:'
$lblWarn.Location = New-Object System.Drawing.Point(624, 448)
$lblWarn.Size     = New-Object System.Drawing.Size(448, 18)
$form.Controls.Add($lblWarn)

$txtWarn            = New-Object System.Windows.Forms.TextBox
$txtWarn.Multiline  = $true
$txtWarn.ReadOnly   = $true
$txtWarn.ScrollBars = 'Vertical'
$txtWarn.WordWrap   = $true
$txtWarn.Location   = New-Object System.Drawing.Point(624, 468)
$txtWarn.Size       = New-Object System.Drawing.Size(448, 240)
$txtWarn.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtWarn)

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'MailFlowConnectors' `
                -X 12 -Y 94 -Width 1060 -StatusLabel $lblStatus

$btnOpen.Add_Click({
    try { Open-ReportFolder -Path $ReportDir }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Open reports folder' }
})

# The selected grid row's Tag carries the row object (grid cells plus Raw).
# State lives on the controls, not in a variable a handler has to find.
function Get-SelectedConnectorRow {
    if ($grid.SelectedRows.Count -eq 0) { return $null }
    return $grid.SelectedRows[0].Tag
}

$grid.Add_SelectionChanged({
    try {
        $r = Get-SelectedConnectorRow
        $btnDetails.Enabled = ($null -ne $r)
        $btnTest.Enabled    = ($null -ne $r -and $r.Direction -eq 'Outbound')
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Connector selection' }
})

$btnRun.Add_Click({
    $btnRun.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtDetail.Text = ''
    $txtWarn.Text   = ''
    $grid.Rows.Clear()
    $btnExport.Enabled  = $false
    $btnDetails.Enabled = $false
    $btnTest.Enabled    = $false
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity 'Mail flow connectors'
        $report = Get-ConnectorReport -Progress $progress

        foreach ($r in $report.Rows) {
            $values = @()
            foreach ($col in $GridColumns) { $values += "$($r.$col)" }
            $idx = $grid.Rows.Add($values)
            $grid.Rows[$idx].Tag = $r
            if ($r.Enabled -eq 'False') {
                $grid.Rows[$idx].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray
            }
        }
        $grid.ClearSelection()

        Complete-ProgressReporter -Activity 'Mail flow connectors' -Bar $pbar

        $warnings = Get-ConnectorWarnings -Rows $report.Rows
        if ($warnings.Count -eq 0) {
            $txtWarn.Text = 'Nothing flagged.' + [Environment]::NewLine + [Environment]::NewLine +
                            'Checked: disabled connectors; OnPremises outbound routing by MX record; OnPremises inbound without RequireTls; OnPremises connectors with no certificate or TLS domain.'
        } else {
            $txtWarn.Text = (($warnings | ForEach-Object { "* $_" }) -join ([Environment]::NewLine + [Environment]::NewLine))
        }

        $n = $report.InboundCount
        $m = $report.OutboundCount
        $lblStatus.Text = "Done. $n inbound, $m outbound connector(s). $($warnings.Count) warning(s)."
        $btnExport.Enabled = ($report.Rows.Count -gt 0)

        if ($n -eq 0 -and $m -eq 0) {
            $txtDetail.Text = 'No connectors were returned. A tenant with no inbound and no outbound connector has no hybrid or partner mail flow configured -- if that is unexpected, check the connection is to the right tenant.'
        }

        Write-CompletedLog -Task "Mail flow connectors -- $n inbound, $m outbound" -Result 'Success'
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtDetail.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mail flow connectors'
    } finally {
        Complete-ProgressReporter -Activity 'Mail flow connectors'
        $form.Cursor    = $prev
        $btnRun.Enabled = $true
    }
})

$btnDetails.Add_Click({
    try {
        $r = Get-SelectedConnectorRow
        if ($null -eq $r) { $lblStatus.Text = 'Select a connector first.'; return }
        $txtDetail.Text = "$($r.Direction) connector -- $($r.Name)" + [Environment]::NewLine +
                          ('-' * 60) + [Environment]::NewLine +
                          (Format-ConnectorDetail -Object $r.Raw)
        $txtDetail.SelectionStart  = 0
        $txtDetail.SelectionLength = 0
        $txtDetail.ScrollToCaret()
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Connector details' }
})

$btnTest.Add_Click({
    try {
        $r = Get-SelectedConnectorRow
        if ($null -eq $r) { $lblStatus.Text = 'Select an outbound connector first.'; return }
        if ($r.Direction -ne 'Outbound') {
            $lblStatus.Text = 'Validate-OutboundConnector only tests OUTBOUND connectors. Select one.'
            return
        }
        if (-not (Get-Command -Name 'Validate-OutboundConnector' -ErrorAction SilentlyContinue)) {
            throw 'Validate-OutboundConnector is not available. Connect to Exchange Online first.'
        }

        $prompt = "Test outbound connector '$($r.Name)'." + [Environment]::NewLine +
                  "Recipient domains: $($r.RecipientDomains)" + [Environment]::NewLine + [Environment]::NewLine +
                  'One address to send the test message to. It must be in a domain this connector routes to.'
        $recipient = Show-InputDialog -Title 'Test outbound connector' -Prompt $prompt -Width 520
        $recipient = "$recipient".Trim()
        if (-not $recipient) { $lblStatus.Text = 'Test cancelled.'; return }
        if ($recipient -notmatch '^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$') {
            $lblStatus.Text = "'$recipient' does not look like one SMTP address. Nothing sent."
            return
        }

        $preview = "This SENDS a real test message." + [Environment]::NewLine + [Environment]::NewLine +
                   "  Connector : $($r.Name)" + [Environment]::NewLine +
                   "  Type      : $($r.ConnectorType)" + [Environment]::NewLine +
                   "  Smart host: $($r.SmartHosts)" + [Environment]::NewLine +
                   "  Recipient : $recipient" + [Environment]::NewLine + [Environment]::NewLine +
                   'The connector itself is not changed. Continue?'
        $answer = Show-ToolDialog($preview, 'Test outbound connector', 'YesNo', 'Warning', 'Button2')
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { $lblStatus.Text = 'Test cancelled.'; return }

        $prev = $form.Cursor
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $btnTest.Enabled = $false
        $lblStatus.Text = "Testing '$($r.Name)'..."
        try {
            $result = Validate-OutboundConnector -Identity $r.Name -Recipients $recipient -ErrorAction Stop
            $txtDetail.Text = Format-ValidationResult -Result $result -ConnectorName $r.Name -Recipient $recipient
            $lblStatus.Text = "Test sent through '$($r.Name)' to $recipient. See the result on the left."
            # Get-LogTarget, not the bare name (TGT001). A connector has no
            # address -- its Name IS its identity -- so the first candidate is
            # the identity Exchange uses and the name follows it. The point of
            # the helper is that the audit row says which thing it meant.
            Write-CompletedLog -Action 'ValidateConnector' -Target (Get-LogTarget "$($r.Identity)" "$($r.Name)" 'outbound connector') -Result 'Success' `
                -Detail "Validate-OutboundConnector '$($r.Name)' -> $recipient"
        } catch {
            $txtDetail.Text = Format-ErrorDetail -ErrorRecord $_ -Context "Validate-OutboundConnector '$($r.Name)'"
            $lblStatus.Text = 'Test failed.'
            Write-FailedLog -Action 'ValidateConnector' -Target (Get-LogTarget "$($r.Identity)" "$($r.Name)" 'outbound connector') `
                -Detail "Validate-OutboundConnector '$($r.Name)' -> $recipient" -ErrorRecord $_
        } finally {
            $form.Cursor = $prev
            $btnTest.Enabled = $true
        }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Test outbound connector' }
})

$btnExport.Add_Click({
    try {
        if ($grid.Rows.Count -eq 0) { $lblStatus.Text = 'Nothing to export. Run the report first.'; return }

        $export = New-Object System.Collections.Generic.List[object]
        foreach ($row in $grid.Rows) {
            $r = $row.Tag
            if ($null -eq $r) { continue }
            $o = [ordered]@{}
            foreach ($col in $GridColumns) { $o[$col] = "$($r.$col)" }
            $export.Add([PSCustomObject]$o)
        }

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir "MailFlowConnectors_$stamp.csv"
        $export | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $files.SelectPath($csv)
        $lblStatus.Text = "Exported $($export.Count) row(s) to $csv"
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Export connectors CSV' }
})

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Mail Flow Connectors' | Out-Null

[void]$form.ShowDialog()
#endregion
