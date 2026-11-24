#Requires -Version 5.1
<#
    Protocols.ps1
    Reports client-access protocol state for every mailbox and exports to CSV. Author : Glenn Armbruster Revised : 2026-07-28

    Reasoning: ScriptNotes/Protocols.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-07-31
#>
# TOOLKIT-HELP-BEGIN
# Reports which client protocols are switched on for every mailbox -- POP, IMAP,
# ActiveSync, Outlook on the web and the rest -- and writes a CSV.
#
# Useful for security review: legacy protocols are the ones most often used in
# password-spray attacks, and this shows where they are still enabled.
#
# Read-only. It changes nothing.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Configuration --------------------------------------------------------
$ReportDir = Get-ReportPath
$Protocols = @('OWA', 'ActiveSync', 'POP', 'IMAP', 'MAPI', 'EWS')
#endregion

#region Data -----------------------------------------------------------------
function Get-EwsOrgDefault {
    try {
        $org = Invoke-EXOCommand -Name 'Get-OrganizationConfig' -Parameters @{ ErrorAction = 'Stop' }
        if ($null -ne $org.EwsEnabled) { return [bool]$org.EwsEnabled }
    } catch { }
    return $true
}

function Get-CmdletParameterNames {
    param([string]$Name)
    try {
        $cmd = Get-EXOCommand -Name $Name
        if ($cmd.Parameters) { return @($cmd.Parameters.Keys) }
    } catch { }
    try {
        $cmd = Get-Command -Name $Name -ErrorAction Stop
        if ($cmd.Parameters) { return @($cmd.Parameters.Keys) }
    } catch { }
    return @()
}

function Invoke-CasQuery {
    param(
        [string]$Cmdlet,
        [string[]]$Props,
        [string]$RecipientType,
        [string]$Filter,
        [scriptblock]$Progress
    )
    $supported = Get-CmdletParameterNames -Name $Cmdlet

    $remaining = @($Props)
    for ($attempt = 1; $attempt -le 5; $attempt++) {
    # Only pass parameters the installed cmdlet declares -- the module surface
    # varies by version and a guessed parameter fails the whole call.
        $params = @{ ErrorAction = 'Stop' }
        if ($supported -contains 'ResultSize') { $params['ResultSize'] = 'Unlimited' }
        if ($supported -contains 'Properties') { $params['Properties'] = $remaining }
        if ($supported -contains 'PageSize')   { $params['PageSize']   = 500 }
        if ($RecipientType -and $supported -contains 'RecipientTypeDetails') {
            $params['RecipientTypeDetails'] = $RecipientType
        }
        if ($Filter -and $supported -contains 'Filter') { $params['Filter'] = $Filter }
        try {
            return @(Invoke-EXOCommand -Name $Cmdlet -Parameters $params)
        } catch {
            $msg = "$($_.Exception.Message)"

            $m = [regex]::Match($msg, 'InvalidProperties\s*=\s*(.+?)\s*$')
            if ($m.Success) {
                $bad = $m.Groups[1].Value -split '\s*,\s*' | Where-Object { $_ }
                $remaining = @($remaining | Where-Object { $bad -notcontains $_ })
                if (-not $remaining) { throw }
                & $Progress ("Dropped unsupported propert{0}: {1}. Retrying..." -f
                             $(if ($bad.Count -eq 1) { 'y' } else { 'ies' }), ($bad -join ', '))
                continue
            }

            $transientPattern = 'underlying connection was closed|operation has timed out|' +
                                'unexpected error occurred on a receive|task was canceled|' +
                                'Unable to read data from the transport|forcibly closed'
            if (($msg -match $transientPattern) -and $attempt -lt 5) {
                $wait = [math]::Pow(2, $attempt)
                & $Progress "Connection dropped. Retrying in $wait s (attempt $($attempt + 1) of 5)..."
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }
    throw "Could not query $Cmdlet after $attempt attempts."
}

function Get-ProtocolReport {
    param([scriptblock]$Progress)

    $cmdName = 'Get-EXOCASMailbox'
    if (-not (Get-Command $cmdName -ErrorAction SilentlyContinue)) {
        $cmdName = 'Get-CASMailbox'
    }

    # RecipientTypeDetails is a FILTER parameter on Get-EXOCasMailbox, not a
    # returnable property -- asking for it fails the call outright.
    $props = @('PrimarySmtpAddress', 'DisplayName',
               'OWAEnabled', 'ActiveSyncEnabled', 'PopEnabled', 'ImapEnabled',
               'MapiEnabled', 'EwsEnabled')

    $types = @('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox',
               'LinkedMailbox', 'TeamMailbox', 'DiscoveryMailbox')

    $collected = New-Object System.Collections.Generic.List[object]
    $skipped   = @()
    $script:Subdivided = @()

    # Slice ONLY if the cmdlet accepts the filter. Otherwise it is dropped and
    # every pass returns every mailbox -- duplicates stamped with the wrong type.
    $supported = Get-CmdletParameterNames -Name $cmdName
    $useTypes  = ($supported -contains 'RecipientTypeDetails')

    if ($useTypes) {
        $canFilter = ($supported -contains 'Filter')
        $letters   = @([char[]](97..122) | ForEach-Object { "$_" }) + @('0','1','2','3','4','5','6','7','8','9')

        $i = 0
        foreach ($type in $types) {
            $i++
            & $Progress "Querying $type ($i of $($types.Count))..." ([int](100 * ($i - 1) / $types.Count))
            $got = $null
            try {
                $got = @(Invoke-CasQuery -Cmdlet $cmdName -Props $props `
                                         -RecipientType $type -Progress $Progress)
            } catch {
                $firstErr = "$($_.Exception.Message)"

                if (-not $canFilter) {
                    $skipped += "$type ($firstErr)"
                    & $Progress "$type failed and cannot be subdivided -- skipping."
                    continue
                }

                & $Progress "$type too large -- splitting by first letter..."
                $script:Subdivided += $type
                $got  = New-Object System.Collections.Generic.List[object]
                $lost = @()
                $n = 0
                foreach ($ch in $letters) {
                    $n++
                    $pctBase = 100 * ($i - 1) / $types.Count
                    $pctIn    = (100.0 / $types.Count) * ($n / $letters.Count)
                    & $Progress "$type : slice $ch ($n of $($letters.Count)) -- $($collected.Count + $got.Count) so far..." ([int]($pctBase + $pctIn))
                    try {
                        $part = @(Invoke-CasQuery -Cmdlet $cmdName -Props $props `
                                                  -RecipientType $type `
                                                  -Filter "Alias -like '$ch*'" -Progress $Progress)
                        foreach ($p in $part) { $got.Add($p) }
                    } catch {
                        $lost += $ch
                    }
                }
                if ($lost.Count -gt 0) {
                    $skipped += "$type -- slices [$($lost -join ',')] failed"
                }
                if ($got.Count -eq 0) { $skipped += "$type ($firstErr)"; continue }
            }

            foreach ($s in $got) {
                Add-Member -InputObject $s -NotePropertyName '_Type' `
                           -NotePropertyValue $type -Force
                $collected.Add($s)
            }
            & $Progress "$type : $($got.Count). Running total $($collected.Count)."
        }
    } else {
        & $Progress 'Querying all mailboxes...'
        $all = @(Invoke-CasQuery -Cmdlet $cmdName -Props $props -Progress $Progress)
        foreach ($s in $all) { $collected.Add($s) }
    }

    & $Progress 'Reading organisation EWS default...'
    $ewsDefault = Get-EwsOrgDefault

    & $Progress 'Removing any duplicates...'
    $seen  = @{}
    $uniq  = New-Object System.Collections.Generic.List[object]
    $dupes = 0
    foreach ($m in $collected) {
        $k = "$($m.PrimarySmtpAddress)".ToLower()
        if (-not $k) { continue }
        if ($seen.ContainsKey($k)) { $dupes++; continue }
        $seen[$k] = $true
        $uniq.Add($m)
    }
    $collected = $uniq

    & $Progress 'Building report...'
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($m in $collected) {
        $ews = if ($null -eq $m.EwsEnabled) { $ewsDefault } else { [bool]$m.EwsEnabled }
        $rows.Add([PSCustomObject][ordered]@{
            PrimarySmtpAddress = "$($m.PrimarySmtpAddress)"
            OWA                = [bool]$m.OWAEnabled
            ActiveSync         = [bool]$m.ActiveSyncEnabled
            POP                = [bool]$m.PopEnabled
            IMAP               = [bool]$m.ImapEnabled
            MAPI               = [bool]$m.MapiEnabled
            EWS                = $ews
            RecipientType      = "$($m._Type)"
            DisplayName        = "$($m.DisplayName)"
        })
    }
    return [PSCustomObject]@{
        Rows       = $rows
        Cmdlet     = $cmdName
        EwsDefault = $ewsDefault
        Skipped    = $skipped
        Duplicates = $dupes
        Subdivided = $script:Subdivided
    }
}

function Format-ReportSummary {
    param($Result)
    $rows = $Result.Rows
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Mailboxes  : $($rows.Count)")
    [void]$sb.AppendLine("Source     : $($Result.Cmdlet)")
    [void]$sb.AppendLine("EWS default: $($Result.EwsDefault)  (used where the mailbox has no explicit setting)")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Enabled counts:')
    foreach ($p in $Protocols) {
        $on = @($rows | Where-Object { $_.$p }).Count
        $pct = if ($rows.Count) { [math]::Round(100 * $on / $rows.Count) } else { 0 }
        [void]$sb.AppendLine(("  {0,-11}{1,6} of {2}  ({3}%)" -f $p, $on, $rows.Count, $pct))
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('By recipient type:')
    $rows | Group-Object RecipientType | Sort-Object Count -Descending | ForEach-Object {
        [void]$sb.AppendLine(("  {0,-22}{1,6}" -f $_.Name, $_.Count))
    }
    if ($Result.Subdivided -and $Result.Subdivided.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("Subdivided by alias letter (too large for one call): $($Result.Subdivided -join ', ')")
        [void]$sb.AppendLine('  Slicing keys on Alias a-z/0-9. Sanity-check the total against the tenant;')
        [void]$sb.AppendLine('  an alias starting with another character would not be picked up.')
    }
    if ($Result.Duplicates -gt 0) {
        [void]$sb.AppendLine("Duplicate rows removed: $($Result.Duplicates)")
    }
    if ($Result.Skipped -and $Result.Skipped.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('*** INCOMPLETE -- these types could not be read: ***')
        foreach ($s in $Result.Skipped) { [void]$sb.AppendLine("  $s") }
    }
    return $sb.ToString()
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Report -- Mailbox Protocols'
$form.Size            = New-Object System.Drawing.Size(760, 656)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Reports OWA, ActiveSync, POP, IMAP, MAPI and EWS for every mailbox in the connected tenant.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblInfo)

$lblTenant           = New-Object System.Windows.Forms.Label
$lblTenant.Tag       = 'NoTheme'
$lblTenant.Location  = New-Object System.Drawing.Point(12, 34)
$lblTenant.Size      = New-Object System.Drawing.Size(720, 18)
$lblTenant.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblTenant)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run Report'
$btnRun.Location = New-Object System.Drawing.Point(12, 60)
$btnRun.Size     = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnRun)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Reports Folder'
$btnOpen.Location = New-Object System.Drawing.Point(140, 60)
$btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnOpen)


$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(622, 60)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$lblStatus           = New-Object System.Windows.Forms.Label
$lblStatus.Text      = 'Idle.'
$lblStatus.Location  = New-Object System.Drawing.Point(12, 134)
$lblStatus.Size      = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 154 -Width 720 -Height 12

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Location   = New-Object System.Drawing.Point(12, 172)
$txtOut.Size       = New-Object System.Drawing.Size(720, 424)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'Protocols' `
                -X 12 -Y 96 -Width 720 -StatusLabel $lblStatus


$btnOpen.Add_Click({ Open-ReportFolder -Path $ReportDir })

$btnRun.Add_Click({
    $btnRun.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = ''
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar `
                        -Activity 'Mailbox protocol report'
        $report = Get-ProtocolReport -Progress $progress

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        if ($report.Rows.Count -eq 0) {
            $lblStatus.Text = 'No mailboxes returned -- nothing written.'
            $txtOut.Text = (Format-ReportSummary -Result $report) +
                           [Environment]::NewLine +
                           'No CSV was written because the query returned no rows.'
            return
        }

        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir "Protocols_$stamp.csv"
        $report.Rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $files.SelectPath($csv)
        Complete-ProgressReporter -Activity 'Mailbox protocol report' -Bar $pbar
        $lblStatus.Text  = "Done. $($report.Rows.Count) row(s) written."

        $txtOut.Text = (Format-ReportSummary -Result $report) +
                       [Environment]::NewLine +
                       "Saved to:" + [Environment]::NewLine + "  $csv"

        $logNote = if ($report.Skipped -and $report.Skipped.Count) { ' -- PARTIAL' } else { '' }
        # PARTIAL is the keyword the log is scanned for; Result is the same
        # fact in a column a filter can see.
        if ($logNote) { $logResult = 'Partial' } else { $logResult = 'Success' }
        Write-CompletedLog -Task "Protocol report -- $($report.Rows.Count) mailboxes$logNote" -Result $logResult
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Protocol report'
    } finally {
        Complete-ProgressReporter -Activity 'Mailbox protocol report'
        $form.Cursor    = $prev
        $btnRun.Enabled = $true
    }
})

$d = Get-ConnectedTenantDomain
if ($d) {
    $lblTenant.Text      = "Tenant: $d"
    $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(0, 128, 0)
} else {
    $lblTenant.Text      = 'Tenant: not connected -- connect to Exchange Online first.'
    $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(200, 0, 0)
}

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Protocols' | Out-Null
[void]$form.ShowDialog()
#endregion
