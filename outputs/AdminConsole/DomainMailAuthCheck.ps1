#Requires -Version 5.1
<#
    DomainMailAuthCheck.ps1
    MX, SPF, DKIM and DMARC for every accepted domain, checked against DNS and against what Exchange Online expects.

    Author   : Glenn Armbruster
    Revised  : 2026-09-12
#>
# TOOLKIT-HELP-RUNS: reads accepted domains and DKIM signing configs from Exchange Online, then queries public DNS for each domain. Changes nothing.
# TOOLKIT-VARIANT: Cloud
# TOOLKIT-HELP-BEGIN
# One grid answering "is mail authentication set up for this domain" for
# every accepted domain in the tenant, or for one you pick.
#
# Load domains reads Get-AcceptedDomain and Get-DkimSigningConfig. Run
# checks then asks DNS, per domain:
#
#   * MX      -- the records that exist. FAIL when there are none.
#   * SPF     -- the TXT record starting v=spf1. FAIL when missing or when
#                there is more than one (RFC 7208 allows exactly one), FAIL
#                on +all, WARN on ?all or no all mechanism at all.
#   * DKIM config      -- Enabled and Status from Get-DkimSigningConfig.
#   * DKIM selector1/2 -- the selector1._domainkey and selector2._domainkey
#                CNAMEs must resolve to the Selector1CNAME / Selector2CNAME
#                values Exchange Online reports. Anything else is FAIL.
#   * DMARC   -- the _dmarc TXT record. FAIL when missing, WARN on p=none,
#                PASS on p=quarantine or p=reject.
#
# Detail is the record actually found, so a FAIL can be read without
# opening a DNS tool. Select a row to see the whole record in the box under
# the grid.
#
# DNS is asked through Resolve-DnsName (the DnsClient module, Windows 8 /
# Server 2012 and later). System DNS is used unless you choose a public
# resolver -- useful when the internal resolver answers for a domain it
# hosts a split view of. Without the module the DNS rows say so and the
# DKIM config rows still run.
#
# Read-only. Export writes the grid to CSV.
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
$ReportPrefix = 'DomainMailAuthCheck'
$AllDomainsItem = '(all accepted domains)'

# Grid column order, shared by the grid and the CSV.
$GridColumns = @('Domain', 'Check', 'Result', 'Detail')
$GridWidths  = @{ Domain = 200; Check = 130; Result = 70; Detail = 560 }

# Resolver choices. The first entry means "whatever this machine uses";
# the rest are public resolvers by address, and the box is editable so any
# other address can be typed.
$SystemDnsItem = 'System DNS (default)'
$ResolverItems = @($SystemDnsItem, '1.1.1.1', '8.8.8.8', '9.9.9.9')

# File-scope state. Kept in $script: at file scope only -- never read from
# anything Common.ps1 owns.
$script:Domains = @()      # shaped accepted-domain objects from the last load
$script:Dkim    = @{}      # domain (lower) -> shaped DKIM config
$script:Rows    = @()      # every result row from the last run

# Row colours. Falls back to plain red when the theme does not define one;
# WARN is a fixed amber that reads on both light and dark palettes.
$script:ColorFail = $null
try {
    $theme = Get-ConsoleTheme
    if ($theme) { $script:ColorFail = ConvertTo-ThemeColor "$($theme.DangerColor)" }
} catch { }
if (-not $script:ColorFail) { $script:ColorFail = [System.Drawing.Color]::FromArgb(214, 69, 69) }
$script:ColorWarn = [System.Drawing.Color]::FromArgb(206, 132, 22)
#endregion

#region Data -----------------------------------------------------------------
function Test-DnsClientAvailable {
    return [bool](Get-Command -Name 'Resolve-DnsName' -ErrorAction SilentlyContinue)
}

function Get-DnsAnswer {
    <#
        One DNS query, as the toolkit wants to see it:

            Found   $true when at least one record of the asked type came back
            Records the records of that type only (a CNAME chain in front of a
                    TXT answer is dropped, not misread as a TXT)
            Error   '' or the resolver's message when the query itself failed;
                    "name does not exist" is NOT an error, it is Found=$false

        -DnsOnly so the hosts file cannot answer for a mail domain, which it
        has done.
    #>
    param([string]$Name, [string]$Type, [string]$Server)

    $out = [ordered]@{ Found = $false; Records = @(); Error = '' }
    $p = @{ Name = $Name; Type = $Type; DnsOnly = $true; ErrorAction = 'Stop' }
    if ($Server) { $p['Server'] = $Server }
    try {
        $all = @(Resolve-DnsName @p)
    } catch {
        $m = "$($_.Exception.Message)"
        if ($m -match '(?i)does not exist|non-existent|NXDOMAIN|no such host') { return [PSCustomObject]$out }
        $out.Error = $m
        return [PSCustomObject]$out
    }
    $recs = @()
    foreach ($r in $all) {
        # DnsRecord objects carry Type; a record without it (older module
        # builds) is kept rather than dropped.
        $t = $null
        if ($r.PSObject.Properties['Type']) { $t = "$($r.Type)" }
        if ($t -and $t -ne $Type) { continue }
        $recs += $r
    }
    $out.Records = @($recs)
    $out.Found   = ($recs.Count -gt 0)
    return [PSCustomObject]$out
}

function Get-TxtStrings {
    <#
        TXT records as whole strings. A long TXT (an SPF with many includes)
        arrives split into 255-byte chunks in .Strings; the record is the
        chunks joined, not the first chunk.
    #>
    param($Records)
    $out = @()
    foreach ($r in @($Records)) {
        $s = ''
        if ($r.PSObject.Properties['Strings']) { $s = (@($r.Strings) -join '') }
        elseif ($r.PSObject.Properties['Text']) { $s = "$($r.Text)" }
        $s = "$s".Trim()
        if ($s) { $out += $s }
    }
    return @($out)
}

function ConvertTo-DnsHostText {
    # DNS answers sometimes carry a trailing dot; comparisons and display
    # should not.
    param([string]$Name)
    return ("$Name".Trim().TrimEnd('.')).ToLower()
}

function New-CheckRow {
    param([string]$Domain, [string]$Check, [string]$Result, [string]$Detail)
    return [PSCustomObject][ordered]@{
        Domain = $Domain
        Check  = $Check
        Result = $Result
        Detail = $Detail
    }
}

function Get-MxRows {
    param([string]$Domain, [string]$Server)
    $a = Get-DnsAnswer -Name $Domain -Type 'MX' -Server $Server
    if ($a.Error) { return @(New-CheckRow $Domain 'MX' 'WARN' "DNS query failed: $($a.Error)") }
    if (-not $a.Found) { return @(New-CheckRow $Domain 'MX' 'FAIL' 'No MX record. Mail for this domain has nowhere to go.') }
    $parts = @()
    foreach ($r in ($a.Records | Sort-Object { if ($_.PSObject.Properties['Preference']) { [int]$_.Preference } else { 0 } })) {
        $pref = if ($r.PSObject.Properties['Preference']) { "$($r.Preference)" } else { '?' }
        $mxHost = if ($r.PSObject.Properties['NameExchange']) { ConvertTo-DnsHostText "$($r.NameExchange)" } else { "$r" }
        $parts += "$pref $mxHost"
    }
    return @(New-CheckRow $Domain 'MX' 'PASS' ($parts -join '; '))
}

function Get-SpfRows {
    param([string]$Domain, [string]$Server)
    $a = Get-DnsAnswer -Name $Domain -Type 'TXT' -Server $Server
    if ($a.Error) { return @(New-CheckRow $Domain 'SPF' 'WARN' "DNS query failed: $($a.Error)") }
    $spf = @(Get-TxtStrings -Records $a.Records | Where-Object { $_ -match '^(?i)v=spf1(\s|$)' })
    if ($spf.Count -eq 0) { return @(New-CheckRow $Domain 'SPF' 'FAIL' 'No TXT record starting v=spf1.') }
    if ($spf.Count -gt 1) {
        return @(New-CheckRow $Domain 'SPF' 'FAIL' "$($spf.Count) SPF records -- receivers treat more than one as a permanent error. Found: " + ($spf -join '  |  '))
    }
    $rec = $spf[0]
    $result = 'PASS'
    $note   = ''
    if ($rec -match '(?i)(^|\s)\+?all(\s|$)' -and $rec -notmatch '(?i)(^|\s)[-~?]all(\s|$)') {
        # A bare or +all passes every sender on earth, which is worse than no
        # record at all because it looks configured.
        $result = 'FAIL'; $note = '+all / bare all: every sender passes. '
    } elseif ($rec -match '(?i)(^|\s)\?all(\s|$)') {
        $result = 'WARN'; $note = '?all (neutral): unauthorised senders are neither passed nor failed. '
    } elseif ($rec -notmatch '(?i)(^|\s)[-~+?]?all(\s|$)' -and $rec -notmatch '(?i)(^|\s)redirect=') {
        $result = 'WARN'; $note = 'No all mechanism and no redirect: unlisted senders get a neutral result. '
    }
    return @(New-CheckRow $Domain 'SPF' $result ($note + $rec))
}

function Get-DkimConfigRows {
    param([string]$Domain, $Config)
    if ($null -eq $Config) {
        return @(New-CheckRow $Domain 'DKIM config' 'WARN' 'No DKIM signing config for this domain in Exchange Online. Outbound mail from it is signed with the default onmicrosoft selector, not this domain.')
    }
    $enabled = "$($Config.Enabled)"
    $status  = "$($Config.Status)"
    $detail  = "Enabled=$enabled; Status=$status"
    if ($Config.Selector1CNAME) { $detail += "; Selector1CNAME=$($Config.Selector1CNAME)" }
    if ($Config.Selector2CNAME) { $detail += "; Selector2CNAME=$($Config.Selector2CNAME)" }
    $result = 'WARN'
    if ($enabled -eq 'True' -and $status -ieq 'Valid') { $result = 'PASS' }
    elseif ($enabled -eq 'True') { $result = 'WARN' }
    return @(New-CheckRow $Domain 'DKIM config' $result $detail)
}

function Get-DkimSelectorRow {
    param([string]$Domain, [string]$Selector, [string]$Expected, [bool]$ConfigEnabled, [string]$Server)
    $name  = "$Selector._domainkey.$Domain"
    $check = "DKIM $Selector"
    $a = Get-DnsAnswer -Name $name -Type 'CNAME' -Server $Server
    if ($a.Error) { return (New-CheckRow $Domain $check 'WARN' "DNS query failed for ${name}: $($a.Error)") }
    $found = @()
    foreach ($r in $a.Records) {
        if ($r.PSObject.Properties['NameHost']) { $found += (ConvertTo-DnsHostText "$($r.NameHost)") }
    }
    $expectedText = ConvertTo-DnsHostText $Expected
    if ($found.Count -eq 0) {
        if (-not $expectedText) { return (New-CheckRow $Domain $check 'WARN' "No CNAME at $name, and Exchange Online reports no expected value (no DKIM config).") }
        $r = if ($ConfigEnabled) { 'FAIL' } else { 'WARN' }
        return (New-CheckRow $Domain $check $r "No CNAME at $name. Expected: $expectedText")
    }
    if (-not $expectedText) {
        return (New-CheckRow $Domain $check 'WARN' "$name -> $($found -join '; ') (nothing to compare against: no DKIM config in Exchange Online)")
    }
    if ($found -contains $expectedText) {
        return (New-CheckRow $Domain $check 'PASS' "$name -> $expectedText")
    }
    return (New-CheckRow $Domain $check 'FAIL' "$name -> $($found -join '; ') but Exchange Online expects $expectedText")
}

function Get-DmarcRows {
    param([string]$Domain, [string]$Server)
    $name = "_dmarc.$Domain"
    $a = Get-DnsAnswer -Name $name -Type 'TXT' -Server $Server
    if ($a.Error) { return @(New-CheckRow $Domain 'DMARC' 'WARN' "DNS query failed for ${name}: $($a.Error)") }
    $recs = @(Get-TxtStrings -Records $a.Records | Where-Object { $_ -match '^(?i)v=DMARC1\s*;' })
    if ($recs.Count -eq 0) { return @(New-CheckRow $Domain 'DMARC' 'FAIL' "No TXT record starting v=DMARC1 at $name.") }
    if ($recs.Count -gt 1) {
        return @(New-CheckRow $Domain 'DMARC' 'FAIL' "$($recs.Count) DMARC records -- receivers ignore the policy when there is more than one. Found: " + ($recs -join '  |  '))
    }
    $rec = $recs[0]
    $p = ''
    if ($rec -match '(?i)(^|;)\s*p\s*=\s*([a-z]+)') { $p = $Matches[2].ToLower() }
    switch ($p) {
        'reject'     { return @(New-CheckRow $Domain 'DMARC' 'PASS' "p=reject. $rec") }
        'quarantine' { return @(New-CheckRow $Domain 'DMARC' 'PASS' "p=quarantine. $rec") }
        'none'       { return @(New-CheckRow $Domain 'DMARC' 'WARN' "p=none: monitoring only, nothing is enforced. $rec") }
        ''           { return @(New-CheckRow $Domain 'DMARC' 'FAIL' "No p= tag -- the record is not a valid policy. $rec") }
        default      { return @(New-CheckRow $Domain 'DMARC' 'FAIL' "Unknown p=$p. $rec") }
    }
}

function Get-DomainCheckRows {
    <#
        Every check for one domain. The DNS half is skipped, with one row
        saying why, when Resolve-DnsName is not on this machine; the DKIM
        config row does not need DNS and always runs.
    #>
    param([string]$Domain, $DkimConfig, [string]$Server, [bool]$DnsAvailable)

    $rows = New-Object System.Collections.Generic.List[object]
    $enabled = ($null -ne $DkimConfig -and "$($DkimConfig.Enabled)" -eq 'True')

    if (-not $DnsAvailable) {
        $rows.Add((New-CheckRow $Domain 'DNS' 'WARN' 'Resolve-DnsName is not available (DnsClient module). MX, SPF, DKIM CNAME and DMARC were not checked on this machine.'))
        foreach ($r in @(Get-DkimConfigRows -Domain $Domain -Config $DkimConfig)) { $rows.Add($r) }
        return @($rows)
    }

    foreach ($r in @(Get-MxRows  -Domain $Domain -Server $Server)) { $rows.Add($r) }
    foreach ($r in @(Get-SpfRows -Domain $Domain -Server $Server)) { $rows.Add($r) }
    foreach ($r in @(Get-DkimConfigRows -Domain $Domain -Config $DkimConfig)) { $rows.Add($r) }
    $sel1 = if ($DkimConfig) { "$($DkimConfig.Selector1CNAME)" } else { '' }
    $sel2 = if ($DkimConfig) { "$($DkimConfig.Selector2CNAME)" } else { '' }
    $rows.Add((Get-DkimSelectorRow -Domain $Domain -Selector 'selector1' -Expected $sel1 -ConfigEnabled $enabled -Server $Server))
    $rows.Add((Get-DkimSelectorRow -Domain $Domain -Selector 'selector2' -Expected $sel2 -ConfigEnabled $enabled -Server $Server))
    foreach ($r in @(Get-DmarcRows -Domain $Domain -Server $Server)) { $rows.Add($r) }
    return @($rows)
}

function Get-SelectedResolver {
    # '' means system DNS. Anything else is passed to -Server as typed.
    $t = "$($cboResolver.Text)".Trim()
    if (-not $t -or $t -eq $SystemDnsItem) { return '' }
    return $t
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Domain Mail Auth Check -- MX, SPF, DKIM, DMARC'
$form.Size            = New-Object System.Drawing.Size(1020, 790)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Checks each accepted domain against DNS and against what Exchange Online expects for DKIM. Read-only.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(980, 18)
$form.Controls.Add($lblInfo)

$lblDomain          = New-Object System.Windows.Forms.Label
$lblDomain.Text     = 'Domain:'
$lblDomain.Location = New-Object System.Drawing.Point(12, 43)
$lblDomain.Size     = New-Object System.Drawing.Size(54, 18)
$form.Controls.Add($lblDomain)

$cboDomain               = New-Object System.Windows.Forms.ComboBox
$cboDomain.DropDownStyle = 'DropDownList'
$cboDomain.Location      = New-Object System.Drawing.Point(68, 40)
$cboDomain.Size          = New-Object System.Drawing.Size(320, 24)
$cboDomain.Enabled       = $false
$form.Controls.Add($cboDomain)

$btnLoad          = New-Object System.Windows.Forms.Button
$btnLoad.Text     = 'Load domains'
$btnLoad.Location = New-Object System.Drawing.Point(396, 37)
$btnLoad.Size     = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnLoad)

$lblResolver          = New-Object System.Windows.Forms.Label
$lblResolver.Text     = 'Resolver:'
$lblResolver.Location = New-Object System.Drawing.Point(540, 43)
$lblResolver.Size     = New-Object System.Drawing.Size(60, 18)
$form.Controls.Add($lblResolver)

$cboResolver               = New-Object System.Windows.Forms.ComboBox
$cboResolver.DropDownStyle = 'DropDown'
$cboResolver.Location      = New-Object System.Drawing.Point(604, 40)
$cboResolver.Size          = New-Object System.Drawing.Size(200, 24)
foreach ($r in $ResolverItems) { [void]$cboResolver.Items.Add($r) }
$cboResolver.SelectedIndex = 0
$form.Controls.Add($cboResolver)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(882, 37)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run checks'
$btnRun.Location = New-Object System.Drawing.Point(12, 76)
$btnRun.Size     = New-Object System.Drawing.Size(120, 30)
$btnRun.Enabled  = $false
$form.Controls.Add($btnRun)

$btnExport          = New-Object System.Windows.Forms.Button
$btnExport.Text     = 'Export CSV'
$btnExport.Location = New-Object System.Drawing.Point(140, 76)
$btnExport.Size     = New-Object System.Drawing.Size(110, 30)
$btnExport.Enabled  = $false
$form.Controls.Add($btnExport)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Reports Folder'
$btnOpen.Location = New-Object System.Drawing.Point(258, 76)
$btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnOpen)

$lblDnsNote           = New-Object System.Windows.Forms.Label
$lblDnsNote.Text      = ''
$lblDnsNote.Location  = New-Object System.Drawing.Point(420, 82)
$lblDnsNote.Size      = New-Object System.Drawing.Size(572, 18)
$lblDnsNote.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblDnsNote)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle. Click Load domains to begin.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 116)
$lblStatus.Size     = New-Object System.Drawing.Size(980, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 136 -Width 980 -Height 12

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix $ReportPrefix `
                -X 12 -Y 154 -Width 980 -StatusLabel $lblStatus

$lblFilter          = New-Object System.Windows.Forms.Label
$lblFilter.Text     = 'Filter:'
$lblFilter.Location = New-Object System.Drawing.Point(12, 195)
$lblFilter.Size     = New-Object System.Drawing.Size(44, 18)
$form.Controls.Add($lblFilter)

$txtFilter          = New-Object System.Windows.Forms.TextBox
$txtFilter.Location = New-Object System.Drawing.Point(58, 192)
$txtFilter.Size     = New-Object System.Drawing.Size(300, 24)
$form.Controls.Add($txtFilter)

$lblCount           = New-Object System.Windows.Forms.Label
$lblCount.Text      = 'No results yet.'
$lblCount.Location  = New-Object System.Drawing.Point(370, 195)
$lblCount.Size      = New-Object System.Drawing.Size(622, 18)
$lblCount.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblCount)

$grid                        = New-Object System.Windows.Forms.DataGridView
$grid.Location               = New-Object System.Drawing.Point(12, 222)
$grid.Size                   = New-Object System.Drawing.Size(980, 330)
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
    $grid.Columns[$idx].Width    = $GridWidths[$col]
    $grid.Columns[$idx].SortMode = 'Automatic'
}
$form.Controls.Add($grid)

$lblDetail          = New-Object System.Windows.Forms.Label
$lblDetail.Text     = 'Record found for the selected row:'
$lblDetail.Location = New-Object System.Drawing.Point(12, 560)
$lblDetail.Size     = New-Object System.Drawing.Size(980, 18)
$form.Controls.Add($lblDetail)

$txtDetail            = New-Object System.Windows.Forms.TextBox
$txtDetail.Multiline  = $true
$txtDetail.ReadOnly   = $true
$txtDetail.ScrollBars = 'Vertical'
$txtDetail.WordWrap   = $true
$txtDetail.Location   = New-Object System.Drawing.Point(12, 580)
$txtDetail.Size       = New-Object System.Drawing.Size(980, 150)
$txtDetail.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtDetail)
#endregion

#region UI helpers -----------------------------------------------------------
function Update-Grid {
    <#
        Repaints the grid from $script:Rows through the filter box. The rows
        themselves are never filtered -- Export writes all of them.
    #>
    $filter = "$($txtFilter.Text)".Trim()
    $grid.SuspendLayout()
    try {
        $grid.Rows.Clear()
        $shown = 0
        foreach ($r in @($script:Rows)) {
            if ($filter) {
                $hit = $false
                foreach ($col in $GridColumns) { if ("$($r.$col)" -like "*$filter*") { $hit = $true; break } }
                if (-not $hit) { continue }
            }
            $values = @()
            foreach ($col in $GridColumns) { $values += "$($r.$col)" }
            $idx = $grid.Rows.Add($values)
            $grid.Rows[$idx].Tag = $r
            if ($r.Result -eq 'FAIL')     { $grid.Rows[$idx].DefaultCellStyle.ForeColor = $script:ColorFail }
            elseif ($r.Result -eq 'WARN') { $grid.Rows[$idx].DefaultCellStyle.ForeColor = $script:ColorWarn }
            $shown++
        }
    } finally {
        $grid.ResumeLayout()
    }
    $grid.ClearSelection()
    $total = @($script:Rows).Count
    if ($total -eq 0) {
        $lblCount.Text = 'No results yet.'
    } elseif ($shown -eq $total) {
        $lblCount.Text = "$total row(s)."
    } else {
        $lblCount.Text = "$shown of $total row(s) shown."
    }
}

function Set-DomainCombo {
    $cboDomain.BeginUpdate()
    try {
        $cboDomain.Items.Clear()
        if (@($script:Domains).Count -gt 0) { [void]$cboDomain.Items.Add($AllDomainsItem) }
        foreach ($d in @($script:Domains)) { [void]$cboDomain.Items.Add("$($d.DomainName)") }
    } finally {
        $cboDomain.EndUpdate()
    }
    $cboDomain.Enabled = ($cboDomain.Items.Count -gt 0)
    if ($cboDomain.Items.Count -gt 0) { $cboDomain.SelectedIndex = 0 }
}
#endregion

#region Event handlers -------------------------------------------------------
$btnOpen.Add_Click({
    try { Open-ReportFolder -Path $ReportDir }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Open reports folder' }
})

$txtFilter.Add_TextChanged({
    try { Update-Grid }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Filter results' }
})

$grid.Add_SelectionChanged({
    try {
        if ($grid.SelectedRows.Count -eq 0) { $txtDetail.Text = ''; return }
        $r = $grid.SelectedRows[0].Tag
        if ($null -eq $r) { $txtDetail.Text = ''; return }
        $txtDetail.Text = "$($r.Domain)  --  $($r.Check)  --  $($r.Result)" + [Environment]::NewLine + "$($r.Detail)"
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Show record detail' }
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
        $btnRun.Enabled    = $false
        $btnExport.Enabled = $false
        $script:Domains = @()
        $script:Dkim    = @{}
        $script:Rows    = @()
        Update-Grid
        Set-DomainCombo
        $lblStatus.Text = 'Reading accepted domains and DKIM signing configs...'

        # OFF THE UI THREAD when the process allows it (Start-ToolJob in
        # Common.ps1 explains the test and the fallback). The block travels
        # to the worker as text, so it shapes its own output into plain
        # strings -- the worker sees Common.ps1 and nothing else here.
        Start-ToolJob -Form $form -Text 'Reading accepted domains and DKIM configs ...' -Script {
            param()
            $pDom = @{ ErrorAction = 'Stop' }
            if ((Get-CmdletParameterNames -Name 'Get-AcceptedDomain') -contains 'ResultSize') { $pDom['ResultSize'] = 'Unlimited' }
            $domains = @(Invoke-EXOCommand -Name 'Get-AcceptedDomain' -Parameters $pDom)

            $dkim = @()
            $dkimError = ''
            try {
                $dkim = @(Invoke-EXOCommand -Name 'Get-DkimSigningConfig' -Parameters @{ ErrorAction = 'Stop' })
            } catch { $dkimError = "$($_.Exception.Message)" }

            $outDomains = @()
            foreach ($d in $domains) {
                $outDomains += [PSCustomObject]@{
                    DomainName = "$($d.DomainName)".Trim().ToLower()
                    DomainType = "$($d.DomainType)"
                    Default    = "$($d.Default)"
                }
            }
            $outDkim = @()
            foreach ($k in $dkim) {
                # Keyed by the domain the config is for. Name is the domain on
                # every build seen; Domain is read as a fallback where present.
                $dom = "$($k.Name)".Trim().ToLower()
                if (-not $dom -and $k.PSObject.Properties['Domain']) { $dom = "$($k.Domain)".Trim().ToLower() }
                $outDkim += [PSCustomObject]@{
                    Domain         = $dom
                    Enabled        = "$($k.Enabled)"
                    Status         = "$($k.Status)"
                    Selector1CNAME = "$($k.Selector1CNAME)"
                    Selector2CNAME = "$($k.Selector2CNAME)"
                }
            }
            [PSCustomObject]@{ Domains = $outDomains; Dkim = $outDkim; DkimError = $dkimError }
        } -OnDone {
            param($r, $st)
            try {
                if ($r.Failed) {
                    $lblStatus.Text = "Load failed: $($r.Error)"
                    return
                }
                $data = @($r.Output) | Select-Object -First 1
                $script:Domains = @($data.Domains | Sort-Object { "$($_.DomainName)" })
                $script:Dkim = @{}
                foreach ($k in @($data.Dkim)) { if ($k.Domain) { $script:Dkim[$k.Domain] = $k } }
                Set-DomainCombo
                $btnRun.Enabled = ($script:Domains.Count -gt 0)
                $msg = "$($script:Domains.Count) accepted domain(s), $($script:Dkim.Count) DKIM config(s) loaded."
                if ($data.DkimError) { $msg += " Get-DkimSigningConfig failed: $($data.DkimError)" }
                if ($script:Domains.Count -eq 0) { $msg = 'No accepted domains were returned.' }
                $lblStatus.Text = $msg
            } finally {
                $btnLoad.Enabled = $true
            }
        }
    } catch {
        $lblStatus.Text = 'Load failed.'
        $btnLoad.Enabled = $true
        Show-HandlerError -ErrorRecord $_ -Context 'Load accepted domains'
    }
})

$btnRun.Add_Click({
    try {
        if (@($script:Domains).Count -eq 0) { $lblStatus.Text = 'Load domains first.'; return }

        $pick = "$($cboDomain.SelectedItem)"
        $targets = @()
        if (-not $pick -or $pick -eq $AllDomainsItem) { $targets = @($script:Domains | ForEach-Object { "$($_.DomainName)" }) }
        else { $targets = @($pick) }
        if ($targets.Count -eq 0) { $lblStatus.Text = 'Nothing to check.'; return }

        $server = Get-SelectedResolver
        $dnsOk  = Test-DnsClientAvailable

        # INLINE, NOT Start-ToolJob. Each DNS query is a fraction of a second
        # and the reporter narrates per domain; there is no cloud call here.
        $btnRun.Enabled    = $false
        $btnLoad.Enabled   = $false
        $btnExport.Enabled = $false
        $prev = $form.Cursor
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $script:Rows = @()
        $grid.Rows.Clear()
        $txtDetail.Text = ''
        try {
            $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity 'Domain mail auth check'
            $all = New-Object System.Collections.Generic.List[object]
            $i = 0
            foreach ($d in $targets) {
                $i++
                & $progress "Checking $i of $($targets.Count): $d" ([int](100 * ($i - 1) / $targets.Count))
                $cfg = $null
                if ($script:Dkim.ContainsKey($d)) { $cfg = $script:Dkim[$d] }
                foreach ($row in @(Get-DomainCheckRows -Domain $d -DkimConfig $cfg -Server $server -DnsAvailable $dnsOk)) { $all.Add($row) }
            }
            Complete-ProgressReporter -Activity 'Domain mail auth check' -Bar $pbar

            $script:Rows = @($all)
            Update-Grid

            $fail = @($script:Rows | Where-Object { $_.Result -eq 'FAIL' }).Count
            $warn = @($script:Rows | Where-Object { $_.Result -eq 'WARN' }).Count
            $via  = if ($server) { "resolver $server" } else { 'system DNS' }
            $lblStatus.Text = "Done. $($targets.Count) domain(s) via $via -- $fail FAIL, $warn WARN."
            $btnExport.Enabled = ($script:Rows.Count -gt 0)
        } finally {
            Complete-ProgressReporter -Activity 'Domain mail auth check'
            $form.Cursor     = $prev
            $btnRun.Enabled  = $true
            $btnLoad.Enabled = $true
        }
    } catch {
        $lblStatus.Text = 'Check failed.'
        Show-HandlerError -ErrorRecord $_ -Context 'Domain mail auth check'
    }
})

$btnExport.Add_Click({
    try {
        $rows = @($script:Rows)
        if ($rows.Count -eq 0) { $lblStatus.Text = 'Nothing to export. Run checks first.'; return }

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir ("{0}_{1}.csv" -f $ReportPrefix, $stamp)
        $rows | Select-Object -Property $GridColumns | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $files.SelectPath($csv)
        $lblStatus.Text = "Exported $($rows.Count) row(s) to $csv"
        $fail = @($rows | Where-Object { $_.Result -eq 'FAIL' }).Count
        Write-CompletedLog -Action 'Report' -Target $csv -Result 'Success' `
            -Detail "Domain mail auth check: $($rows.Count) row(s), $fail FAIL."
    } catch {
        $lblStatus.Text = 'Export failed.'
        Show-HandlerError -ErrorRecord $_ -Context 'Export domain mail auth check'
    }
})
#endregion

if (Test-DnsClientAvailable) {
    $lblDnsNote.Text = 'Resolve-DnsName available. DNS checks use the resolver chosen above.'
} else {
    $lblDnsNote.Text = 'Resolve-DnsName NOT available on this machine -- DNS checks will be skipped and say so.'
    $lblDnsNote.ForeColor = $script:ColorWarn
}

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Domain Mail Auth Check' | Out-Null

[void]$form.ShowDialog()
