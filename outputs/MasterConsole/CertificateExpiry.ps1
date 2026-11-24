#Requires -Version 5.1
<#
    CertificateExpiry.ps1
    Certificate expiry across every on-premises Exchange server.

    Read-only. On-premises Exchange only. Reasoning lives in
    ScriptNotes/CertificateExpiry.ps1.notes.md.

    Author   : Glenn Armbruster
    Created  : 2026-08-03
    Revised  : 2026-08-03
#>
# TOOLKIT-HELP-BEGIN
# Reports the certificates on every on-premises Exchange server and how long
# each has left.
#
# Worth running monthly. An expired certificate breaks hybrid quietly -- the
# first symptom is usually somebody complaining that free/busy has stopped
# working between on-premises and the cloud, which does not look like a
# certificate problem.
#
# Certificates in use are separated from ones that are just sitting there, since
# only the first kind will break something.
#
# Read-only. On-premises Exchange only.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Configuration ----------------------------------------------------------
$ReportDir = Get-ReportPath

$script:DefaultWarnDays     = 60
$script:DefaultCriticalDays = 30
#endregion

#region Data -------------------------------------------------------------------
function Get-ExchangeServerList {
    <#
        Server names for the loop.

        Properties are read defensively. Microsoft documents no output
        properties for Get-ExchangeServer at all -- Name, Fqdn and
        AdminDisplayVersion are real in practice but not stated anywhere on the
        reference page, so this asks the object rather than assuming.
    #>
    $out = New-Object System.Collections.ArrayList
    foreach ($s in @(Get-ExchangeServer -ErrorAction Stop)) {
        $name = if ($s.PSObject.Properties['Name']) { "$($s.Name)" } else { "$s" }
        $fqdn = if ($s.PSObject.Properties['Fqdn'] -and $s.Fqdn) { "$($s.Fqdn)" } else { $name }
        $ver  = if ($s.PSObject.Properties['AdminDisplayVersion']) { "$($s.AdminDisplayVersion)" } else { '' }
        if (-not $name) { continue }
        [void]$out.Add([PSCustomObject]@{ Name = $name; Fqdn = $fqdn; Version = $ver })
    }
    return @($out.ToArray())
}

function Get-CertificateRow {
    <#
        One row per certificate per server.

        The server is carried from the loop variable. The certificate object has
        NO documented server property -- Identity is ServerFQDN\Thumbprint, so
        the only alternatives were parsing that string or assuming a property
        Microsoft never documented.

        Expiry is computed from NotAfter and NEVER from Status. There is no
        'Expired' status value: an expired certificate reports 'DateInvalid'.
        A filter on Status -eq 'Expired' returns nothing at all, which is the
        worst possible failure for a report whose entire job is finding expiry.
    #>
    param([Parameter(Mandatory)]$Server, [Parameter(Mandatory)][int]$WarnDays, [Parameter(Mandatory)][int]$CriticalDays)

    $rows = New-Object System.Collections.ArrayList
    $now  = Get-Date

    $certs = @(Get-ExchangeCertificate -Server $Server.Name -ErrorAction Stop)

    foreach ($c in $certs) {
        $notAfter = $null
        if ($c.PSObject.Properties['NotAfter']) { $notAfter = $c.NotAfter }

        $days = $null
        if ($notAfter -is [datetime]) { $days = [int][Math]::Floor(($notAfter - $now).TotalDays) }

        # Services is documented as displaying None/Federation/IIS/IMAP/POP/
        # SMTP/UM/UMCallRouter, but the docs never say whether it is a flags
        # enum, an array or a string. ToString() is the one form that behaves
        # the same for all three.
        $services = ''
        if ($c.PSObject.Properties['Services']) { $services = "$($c.Services)" }
        $bound = ($services -and $services -ne 'None')

        $severity =
            if ($null -eq $days)          { 'Unknown' }
            elseif ($days -lt 0)          { 'EXPIRED' }
            elseif ($days -le $CriticalDays) { 'CRITICAL' }
            elseif ($days -le $WarnDays)     { 'WARNING' }
            else                          { 'OK' }

        [void]$rows.Add([PSCustomObject]@{
            Severity     = $severity
            Server       = $Server.Name
            DaysToExpiry = $days
            NotAfter     = if ($notAfter -is [datetime]) { $notAfter.ToString('yyyy-MM-dd') } else { '' }
            Services     = $services
            InUse        = $bound
            Subject      = if ($c.PSObject.Properties['Subject']) { "$($c.Subject)" } else { '' }
            Status       = if ($c.PSObject.Properties['Status']) { "$($c.Status)" } else { '' }
            SelfSigned   = if ($c.PSObject.Properties['IsSelfSigned']) { [bool]$c.IsSelfSigned } else { $null }
            Issuer       = if ($c.PSObject.Properties['Issuer']) { "$($c.Issuer)" } else { '' }
            Domains      = if ($c.PSObject.Properties['CertificateDomains']) { (@($c.CertificateDomains) -join '; ') } else { '' }
            Thumbprint   = if ($c.PSObject.Properties['Thumbprint']) { "$($c.Thumbprint)" } else { '' }
            NotBefore    = if ($c.PSObject.Properties['NotBefore'] -and $c.NotBefore -is [datetime]) { $c.NotBefore.ToString('yyyy-MM-dd') } else { '' }
        })
    }
    return @($rows.ToArray())
}

function Test-ServerOsAccess {
    <#
        Can this account read the OS layer on each server?

        Nothing here needs it -- certificates come from Exchange RBAC alone.
        It is asked because the FULL health report does need it (disk space on
        the log volume is the most common Exchange outage), and the answer is
        cheaper to collect here than to guess at later.

        Read-only: one WMI query for logical disks, which is exactly what the
        disk check would use. Failure is an answer, not an error.
    #>
    param([Parameter(Mandatory)]$Servers, [scriptblock]$Progress)

    $rows = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($s in $Servers) {
        $i++
        if ($Progress) { & $Progress "Probing OS access on $($s.Name) ($i of $($Servers.Count))..." }
        $ok = $false; $detail = ''
        try {
            $null = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $s.Fqdn -ErrorAction Stop |
                        Select-Object -First 1
            $ok = $true; $detail = 'WMI readable -- disk and service checks are possible.'
        } catch {
            $detail = "$($_.Exception.Message)"
        }
        [void]$rows.Add([PSCustomObject]@{ Server = $s.Name; OsAccess = $ok; Detail = $detail })
    }
    return @($rows.ToArray())
}

function Get-CertificateReport {
    param([int]$WarnDays, [int]$CriticalDays, [switch]$ProbeOs, [scriptblock]$Progress)

    $failed = @()
    & $Progress 'Enumerating Exchange servers...' 2

    $servers = @()
    try { $servers = Get-ExchangeServerList } catch {
        throw "Could not enumerate Exchange servers: $($_.Exception.Message)"
    }
    if ($servers.Count -eq 0) { throw 'Get-ExchangeServer returned nothing.' }

    $all = New-Object System.Collections.ArrayList
    $n = 0
    foreach ($s in $servers) {
        $n++
        & $Progress "Certificates on $($s.Name) ($n of $($servers.Count))..." ([int](5 + 85.0 * $n / $servers.Count))
        try {
            foreach ($r in (Get-CertificateRow -Server $s -WarnDays $WarnDays -CriticalDays $CriticalDays)) {
                [void]$all.Add($r)
            }
        } catch {
            # A server that cannot be read is reported, not fatal. One
            # unreachable server must not hide the expiry dates on the rest.
            $failed += "$($s.Name): $($_.Exception.Message)"
        }
    }

    $osRows = @()
    if ($ProbeOs) {
        & $Progress 'Probing OS access...' 92
        $osRows = Test-ServerOsAccess -Servers $servers -Progress { param($m) & $Progress $m 95 }
    }

    $sorted = @($all | Sort-Object `
        @{ Expression = { switch ($_.Severity) { 'EXPIRED' {0} 'CRITICAL' {1} 'WARNING' {2} 'Unknown' {3} default {4} } } },
        @{ Expression = { if ($null -eq $_.DaysToExpiry) { [int]::MaxValue } else { $_.DaysToExpiry } } },
        Server)

    return [PSCustomObject]@{
        Rows    = $sorted
        Servers = $servers
        Os      = $osRows
        Failed  = $failed
    }
}

function Format-CertificateSummary {
    param($Result, [int]$WarnDays, [int]$CriticalDays)

    $rows = @($Result.Rows)
    $sb = New-Object System.Text.StringBuilder

    [void]$sb.AppendLine("Servers      : $(@($Result.Servers).Count)")
    [void]$sb.AppendLine("Certificates : $($rows.Count)")
    [void]$sb.AppendLine("Thresholds   : critical <= $CriticalDays days, warning <= $WarnDays days")
    [void]$sb.AppendLine('')

    $inUse = @($rows | Where-Object { $_.InUse })
    foreach ($sev in @('EXPIRED', 'CRITICAL', 'WARNING')) {
        $hits = @($rows | Where-Object { $_.Severity -eq $sev })
        $bound = @($hits | Where-Object { $_.InUse })
        if ($hits.Count -eq 0) { continue }
        [void]$sb.AppendLine("--- $sev ($($hits.Count), of which $($bound.Count) bound to a service) ---")
        foreach ($h in $hits) {
            $d = if ($null -eq $h.DaysToExpiry) { '   ?' } else { '{0,4}' -f $h.DaysToExpiry }
            $svc = if ($h.InUse) { $h.Services } else { '(not in use)' }
            [void]$sb.AppendLine(("  {0,-16} {1} days  {2}  {3}" -f $h.Server, $d, $h.NotAfter, $svc))
            [void]$sb.AppendLine("      $($h.Subject)")
        }
        [void]$sb.AppendLine('')
    }

    $ok = @($rows | Where-Object { $_.Severity -eq 'OK' })
    [void]$sb.AppendLine("Not expiring soon : $($ok.Count)")
    [void]$sb.AppendLine("Bound to a service: $($inUse.Count) of $($rows.Count)")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('A certificate with Services = None is not doing anything -- self-signed leftovers')
    [void]$sb.AppendLine('  and the IIS Web Management certificate show up that way. Expiry only matters')
    [void]$sb.AppendLine('  for the ones bound to a service, so those are counted separately above.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Status values come from Exchange verbatim. Note there is no "Expired" status --')
    [void]$sb.AppendLine('  an expired certificate reports DateInvalid. This report computes expiry from')
    [void]$sb.AppendLine('  NotAfter instead, which is why it can be trusted.')

    if (@($Result.Os).Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('--- OS-level access probe (for the future health report) ---')
        foreach ($o in $Result.Os) {
            $mark = if ($o.OsAccess) { 'YES' } else { 'no ' }
            [void]$sb.AppendLine("  $mark  $($o.Server)  --  $($o.Detail)")
        }
        $yes = @($Result.Os | Where-Object { $_.OsAccess }).Count
        [void]$sb.AppendLine("  $yes of $(@($Result.Os).Count) server(s) allow disk and service checks.")
    }

    if (@($Result.Failed).Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('*** INCOMPLETE -- these servers could not be read: ***')
        foreach ($f in $Result.Failed) { [void]$sb.AppendLine("  $f") }
    }
    return $sb.ToString()
}
#endregion

#region WinForms UI ------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Report -- Exchange Certificate Expiry'
$form.Size            = New-Object System.Drawing.Size(880, 720)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.StartPosition   = 'Manual'
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Certificate expiry on every on-premises Exchange server. Read-only -- nothing is changed.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(840, 18)
$form.Controls.Add($lblInfo)

$lblWarn          = New-Object System.Windows.Forms.Label
$lblWarn.Tag      = 'NoTheme'
$lblWarn.Location = New-Object System.Drawing.Point(12, 32)
$lblWarn.Size     = New-Object System.Drawing.Size(840, 18)
$lblWarn.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblWarn)

$emsState = Test-EmsAvailable
if ($emsState.Available) {
    $lblWarn.Text      = 'Connected: on-premises Exchange.'
    $lblWarn.ForeColor = [System.Drawing.Color]::FromArgb(0, 128, 0)
} else {
    $lblWarn.Text      = 'NOT connected to on-premises Exchange -- use Connect ExSE. Nothing can run.'
    $lblWarn.ForeColor = [System.Drawing.Color]::Red
}

$lblCrit          = New-Object System.Windows.Forms.Label
$lblCrit.Text     = 'Critical (days):'
$lblCrit.Location = New-Object System.Drawing.Point(12, 60)
$lblCrit.Size     = New-Object System.Drawing.Size(90, 20)
$form.Controls.Add($lblCrit)

$numCrit          = New-Object System.Windows.Forms.NumericUpDown
$numCrit.Location = New-Object System.Drawing.Point(104, 58)
$numCrit.Size     = New-Object System.Drawing.Size(60, 22)
$numCrit.Minimum  = 1
$numCrit.Maximum  = 365
$numCrit.Value    = $script:DefaultCriticalDays
$form.Controls.Add($numCrit)

$lblWarnD          = New-Object System.Windows.Forms.Label
$lblWarnD.Text     = 'Warning (days):'
$lblWarnD.Location = New-Object System.Drawing.Point(180, 60)
$lblWarnD.Size     = New-Object System.Drawing.Size(94, 20)
$form.Controls.Add($lblWarnD)

$numWarn          = New-Object System.Windows.Forms.NumericUpDown
$numWarn.Location = New-Object System.Drawing.Point(276, 58)
$numWarn.Size     = New-Object System.Drawing.Size(60, 22)
$numWarn.Minimum  = 1
$numWarn.Maximum  = 730
$numWarn.Value    = $script:DefaultWarnDays
$form.Controls.Add($numWarn)

$chkOs          = New-Object System.Windows.Forms.CheckBox
$chkOs.Text     = 'Also probe OS access (for the future health report)'
$chkOs.Location = New-Object System.Drawing.Point(356, 59)
$chkOs.Size     = New-Object System.Drawing.Size(330, 22)
$form.Controls.Add($chkOs)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run Report'
$btnRun.Location = New-Object System.Drawing.Point(12, 88)
$btnRun.Size     = New-Object System.Drawing.Size(120, 30)
$btnRun.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRun)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Reports Folder'
$btnOpen.Location = New-Object System.Drawing.Point(140, 88)
$btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnOpen)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(742, 88)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 164)
$lblStatus.Size     = New-Object System.Drawing.Size(840, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 184 -Width 840 -Height 12

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Location   = New-Object System.Drawing.Point(12, 202)
$txtOut.Size       = New-Object System.Drawing.Size(840, 470)
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Both'
$txtOut.WordWrap   = $false
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'CertificateExpiry' `
            -X 12 -Y 124 -Width 840 -StatusLabel $lblStatus

$btnOpen.Add_Click({
    try { Open-ReportFolder -Path $ReportDir }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Open reports folder' }
})
$btnClose.Add_Click({ $form.Close() })

$btnRun.Add_Click({
    $btnRun.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = ''
    try {
        $ems = Test-EmsAvailable
        if (-not $ems.Available) {
            $txtOut.Text = "Not connected to on-premises Exchange.`r`n`r`n$($ems.Detail)"
            $lblStatus.Text = 'Not connected.'
            return
        }

        $crit = [int]$numCrit.Value
        $warn = [int]$numWarn.Value
        if ($warn -lt $crit) {
            [void][System.Windows.Forms.MessageBox]::Show(
                "Warning ($warn days) is lower than Critical ($crit days).`n`nWarning should be the earlier heads-up.",
                'Thresholds', [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity 'Certificate expiry'
        $report = Get-CertificateReport -WarnDays $warn -CriticalDays $crit `
                    -ProbeOs:$chkOs.Checked -Progress $progress
        Complete-ProgressReporter -Activity 'Certificate expiry' -Bar $pbar

        $summary = Format-CertificateSummary -Result $report -WarnDays $warn -CriticalDays $crit

        if (@($report.Rows).Count -eq 0) {
            $txtOut.Text = $summary + [Environment]::NewLine + 'No certificates returned -- no CSV was written.'
            $lblStatus.Text = 'No certificates found.'
            return
        }

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv = Join-Path $ReportDir "CertificateExpiry_$stamp.csv"
        $report.Rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
        $files.SelectPath($csv)

        $txtOut.Text = $summary + [Environment]::NewLine + 'Saved to:' + [Environment]::NewLine + "  $csv"

        $bad = @($report.Rows | Where-Object { $_.InUse -and $_.Severity -in @('EXPIRED', 'CRITICAL') }).Count
        $lblStatus.Text = if ($bad -gt 0) { "$bad in-use certificate(s) expired or critical." } else { 'Done.' }

        $note = if (@($report.Failed).Count) { ' -- PARTIAL' } else { '' }
        $logResult = if ($note) { 'Partial' } else { 'Success' }
        Write-CompletedLog -Action 'Run' -Target 'Exchange certificates' -Result $logResult `
            -Detail "$(@($report.Rows).Count) certificate(s) across $(@($report.Servers).Count) server(s); $bad in-use expired/critical$note" `
            -Environment 'HQ'
    } catch {
        Show-HandlerError -ErrorRecord $_ -Context 'Certificate expiry report'
        $lblStatus.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Certificate expiry report'
    } finally {
        $form.Cursor = $prev
        Complete-ProgressReporter -Activity 'Certificate expiry'
        $btnRun.Enabled = $true
    }
})

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled, and BringToFront'd inside the helper because WinForms
# draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Cert Check' | Out-Null

[void]$form.ShowDialog()
#endregion
