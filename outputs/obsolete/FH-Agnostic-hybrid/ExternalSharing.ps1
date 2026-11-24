#Requires -Version 5.1
<#
    ExternalSharing.ps1
    Reports the external sharing posture of every SharePoint site and the guest accounts that actually have access, and exports to CSV. Author : Glenn Arm

    Reasoning: ScriptNotes/ExternalSharing.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-07-31
#>
# TOOLKIT-HELP-BEGIN
# Reports how each SharePoint site is configured for external sharing, and which
# guest accounts actually have access.
#
# Both halves matter and neither is enough alone: a site permitting external
# sharing with no guests on it is a setting, while a guest with access to a site
# is a person who can read your files.
#
# Writes a CSV. Read-only.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$DomainController
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

#region Configuration --------------------------------------------------------
$ReportDir = Get-ReportPath

# Get-SPOExternalUser refuses a PageSize above 50, so paging is not optional.
$ExtPageSize = 50

$ExtMaxPerSite = 2000

$RiskOrder = @{
    'ExternalUserAndGuestSharing' = 0   # anonymous links allowed
    'ExternalUserSharing'         = 1   # authenticated guests only
    'ExistingExternalUserSharing' = 2   # only guests already in the directory
    'Disabled'                    = 3   # internal only
}
#endregion

#region Data -----------------------------------------------------------------
function Get-ExternalSharingReport {
    param([scriptblock]$Progress, [switch]$IncludePersonalSites)

    if (-not (Get-Command -Name 'Get-SPOSite' -ErrorAction SilentlyContinue)) {
        throw 'Get-SPOSite is not available. Import Microsoft.Online.SharePoint.PowerShell and run Connect-SPOService first.'
    }

    & $Progress 'Reading the tenant sharing setting...' 2

    $tenantCap = ''
    try {
        $t = Get-SPOTenant -ErrorAction Stop
        $tenantCap = "$($t.SharingCapability)"
    } catch { }

    & $Progress 'Enumerating sites...' 5

    $declared = Get-CmdletParameterNames -Name 'Get-SPOSite'
    $params = @{ ErrorAction = 'Stop' }
    if ($declared -contains 'Limit')               { $params['Limit'] = 'ALL' }
    if ($declared -contains 'IncludePersonalSite') { $params['IncludePersonalSite'] = [bool]$IncludePersonalSites }
    $cappedRisk = ($declared -notcontains 'Limit')

    $sites = @(Get-SPOSite @params)
    if (-not $IncludePersonalSites) {
        $sites = @($sites | Where-Object { "$($_.Url)" -notmatch '/personal/' })
    }

    $total = @($sites).Count
    if ($total -eq 0) { throw 'No sites were returned. Confirm the account has SharePoint administrator rights.' }

    # Get-SPOSite does NOT populate SharingCapability when -Limit or -Filter is
    # used -- it returns a DEFAULT. Enumerate for URLs, then re-fetch each site
    # with -Identity for its real posture. See the notes.
    $canDetail = ($declared -contains 'Identity')
    if (-not $canDetail) {
        throw 'Get-SPOSite does not declare -Identity on this module build, so per-site sharing settings cannot be read accurately. Refusing to report default values as findings.'
    }

    $rows        = New-Object System.Collections.Generic.List[object]
    $siteFailed  = @()
    $extFailed   = @()
    $extCapped   = @()
    $canExternal = [bool](Get-Command -Name 'Get-SPOExternalUser' -ErrorAction SilentlyContinue)
    $n = 0

    foreach ($s in $sites) {
        $n++
        if ($n % 5 -eq 0 -or $n -eq $total) {
            & $Progress "Reading site $n of $total -- $($rows.Count) row(s)..." (8 + [int](88.0 * $n / $total))
        }

        $url = "$($s.Url)"
        try {
            $d = Get-SPOSite -Identity $url -ErrorAction Stop
        } catch {
            $siteFailed += "$url : $($_.Exception.Message)"
            continue
        }

        $cap = "$($d.SharingCapability)"

        $guests   = @()
        $guestCap = $false
        if ($canExternal -and $cap -and $cap -ne 'Disabled') {
            $r = Get-SiteExternalUser -SiteUrl $url
            if ($r.Error) { $extFailed += "$url : $($r.Error)" }
            $guests   = @($r.Users)
            $guestCap = $r.Capped
            if ($guestCap) { $extCapped += $url }
        }

        $rows.Add([PSCustomObject][ordered]@{
            Url                 = $url
            Title               = "$($d.Title)"
            SharingCapability   = $cap
            RiskRank            = if ($RiskOrder.ContainsKey($cap)) { $RiskOrder[$cap] } else { 9 }
            AnonymousLinksAllowed = ($cap -eq 'ExternalUserAndGuestSharing')
            GuestCount          = @($guests).Count
            GuestDomains        = ((@($guests) | ForEach-Object {
                                     if ("$_" -match '@(.+)$') { $Matches[1] } else { '' }
                                   } | Where-Object { $_ } | Sort-Object -Unique) -join '; ')
            GuestAddresses      = ((@($guests) | Sort-Object) -join '; ')
            GuestListTruncated  = $guestCap
            Owner               = "$($d.Owner)"
            Template            = "$($d.Template)"
            IsTeamsConnected    = if ($null -ne $d.IsTeamsConnected) { [bool]$d.IsTeamsConnected } else { $null }
            LockState           = "$($d.LockState)"
            LastModified        = if ($d.LastContentModifiedDate) { $d.LastContentModifiedDate.ToString('yyyy-MM-dd') } else { '' }
        })
    }

    $sorted = @($rows | Sort-Object RiskRank, @{ Expression = 'GuestCount'; Descending = $true }, Url)

    return [PSCustomObject]@{
        Rows              = $sorted
        SiteTotal         = $total
        TenantCapability  = $tenantCap
        SiteFailed        = $siteFailed
        ExternalFailed    = $extFailed
        ExternalCapped    = $extCapped
        CappedRisk        = $cappedRisk
        ExternalAvailable = $canExternal
        PersonalIncluded  = [bool]$IncludePersonalSites
    }
}

function Get-SiteExternalUser {
    param([Parameter(Mandatory)][string]$SiteUrl)

    $out    = New-Object System.Collections.Generic.List[string]
    $pos    = 0
    $err    = ''
    $capped = $false

    $declared = Get-CmdletParameterNames -Name 'Get-SPOExternalUser'
    $hasPaging = ($declared -contains 'Position') -and ($declared -contains 'PageSize')

    while ($true) {
        $p = @{ SiteUrl = $SiteUrl; ErrorAction = 'Stop' }
        if ($hasPaging) {
            $p['Position'] = $pos
            $p['PageSize'] = $ExtPageSize
        }
        try {
            $batch = @(Get-SPOExternalUser @p)
        } catch {
            if ($out.Count -eq 0 -and $pos -eq 0) { $err = '' } else { $err = $_.Exception.Message }
            break
        }

        if (@($batch).Count -eq 0) { break }
        foreach ($u in $batch) {
            $addr = if ($u.Email) { "$($u.Email)" } else { "$($u.DisplayName)" }
            if ($addr) { [void]$out.Add($addr) }
        }

        if (-not $hasPaging) { break }
        if (@($batch).Count -lt $ExtPageSize) { break }
        $pos += $ExtPageSize
        if ($out.Count -ge $ExtMaxPerSite) { $capped = $true; break }
    }

    return [PSCustomObject]@{
        Users  = @($out | Sort-Object -Unique)
        Error  = $err
        Capped = $capped
    }
}

function Format-ReportSummary {
    param($Result)
    $rows = @($Result.Rows)
    $sb = New-Object System.Text.StringBuilder

    [void]$sb.AppendLine("Tenant sharing setting         : $(if ($Result.TenantCapability) { $Result.TenantCapability } else { '(could not read)' })")
    [void]$sb.AppendLine("Sites examined                 : $($Result.SiteTotal)")
    [void]$sb.AppendLine("Personal sites (OneDrive)      : $(if ($Result.PersonalIncluded) { 'included' } else { 'excluded' })")
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine('By sharing capability (worst first):')
    $rows | Group-Object SharingCapability | Sort-Object { if ($RiskOrder.ContainsKey($_.Name)) { $RiskOrder[$_.Name] } else { 9 } } | ForEach-Object {
        [void]$sb.AppendLine(("  {0,-30}{1,6}" -f $_.Name, $_.Count))
    }

    $anon = @($rows | Where-Object { $_.AnonymousLinksAllowed })
    $withGuests = @($rows | Where-Object { $_.GuestCount -gt 0 })
    $guestTotal = ($rows | Measure-Object GuestCount -Sum).Sum

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("Sites allowing anonymous links : $($anon.Count)")
    [void]$sb.AppendLine("Sites with guests inside       : $($withGuests.Count)")
    [void]$sb.AppendLine("Guest entries counted          : $([int]$guestTotal)")

    $domains = @{}
    foreach ($r in $rows) {
        foreach ($d in ($r.GuestDomains -split ';')) {
            $d = $d.Trim()
            if (-not $d) { continue }
            if ($domains.ContainsKey($d)) { $domains[$d] = $domains[$d] + 1 }
            else                          { $domains[$d] = 1 }
        }
    }
    if ($domains.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("Guest domains seen ($($domains.Count) distinct), by number of sites:")
        $domains.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 15 | ForEach-Object {
            [void]$sb.AppendLine(("  {0,-38}{1,4} site(s)" -f $_.Key, $_.Value))
        }
        if ($domains.Count -gt 15) { [void]$sb.AppendLine("  ... and $($domains.Count - 15) more in the CSV.") }
    }

    if ($withGuests.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Most guests:')
        foreach ($r in ($rows | Sort-Object GuestCount -Descending | Select-Object -First 10)) {
            if ($r.GuestCount -eq 0) { break }
            [void]$sb.AppendLine(("  {0,4} guest(s)  {1,-30} {2}" -f $r.GuestCount, $r.SharingCapability, $r.Url))
        }
    }

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('--- Scope, and why ---')
    [void]$sb.AppendLine('This reports the sharing GATE per site and WHO is already inside it. It')
    [void]$sb.AppendLine('  does not enumerate individual sharing links: a link lives on an item, so')
    [void]$sb.AppendLine('  finding them all means walking every file in every library in the tenant.')
    [void]$sb.AppendLine('  For one specific site, use the Purview audit log or the Defender sharing')
    [void]$sb.AppendLine('  reports -- both are buttons in the M365 Reports launcher.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('SharingCapability is read one site at a time on purpose. A bulk')
    [void]$sb.AppendLine('  enumeration with -Limit does NOT populate it -- Microsoft returns a')
    [void]$sb.AppendLine('  default value instead -- so the fast version of this report would be')
    [void]$sb.AppendLine('  uniformly wrong and would not error.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('ExternalUserAndGuestSharing means anonymous "Anyone with the link" is')
    [void]$sb.AppendLine('  permitted. ExternalUserSharing requires the guest to sign in.')
    [void]$sb.AppendLine('  ExistingExternalUserSharing allows only guests already in the directory.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('A site can show 0 guests and still allow anonymous links -- an anonymous')
    [void]$sb.AppendLine('  link creates no guest account. Those rows are the gate standing open')
    [void]$sb.AppendLine('  with nobody recorded as having walked through it.')

    if (-not $Result.ExternalAvailable) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('*** Get-SPOExternalUser was not available -- guest columns are empty. ***')
    }
    if ($Result.ExternalCapped.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("*** PARTIAL -- guest list truncated at $ExtMaxPerSite for $($Result.ExternalCapped.Count) site(s). ***")
    }
    if ($Result.ExternalFailed.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("NOTE -- guest enumeration failed for $($Result.ExternalFailed.Count) site(s):")
        foreach ($f in ($Result.ExternalFailed | Select-Object -First 5)) { [void]$sb.AppendLine("  $f") }
    }
    if ($Result.CappedRisk) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('*** Get-SPOSite does not declare -Limit here; the site list may be ***')
        [void]$sb.AppendLine('*** capped at 200. Treat as INCOMPLETE.                            ***')
    }
    if ($Result.SiteFailed.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("*** INCOMPLETE -- $($Result.SiteFailed.Count) site(s) could not be read: ***")
        foreach ($f in ($Result.SiteFailed | Select-Object -First 10)) { [void]$sb.AppendLine("  $f") }
    }
    return $sb.ToString()
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Report -- External Sharing'
$form.Size            = New-Object System.Drawing.Size(760, 660)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Per-site external sharing setting, plus the guest accounts that already have access.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblInfo)

$lblSrc           = New-Object System.Windows.Forms.Label
$lblSrc.Text      = 'Reads each site individually -- a bulk read returns a default sharing value, not the real one. Allow a few minutes.'
$lblSrc.Location  = New-Object System.Drawing.Point(12, 32)
$lblSrc.Size      = New-Object System.Drawing.Size(720, 18)
$lblSrc.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblSrc)

$chkPersonal          = New-Object System.Windows.Forms.CheckBox
$chkPersonal.Text     = 'Include personal sites (OneDrive) -- much slower on a large tenant'
$chkPersonal.Location = New-Object System.Drawing.Point(12, 54)
$chkPersonal.Size     = New-Object System.Drawing.Size(420, 20)
$chkPersonal.Checked  = $false
$form.Controls.Add($chkPersonal)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run Report'
$btnRun.Location = New-Object System.Drawing.Point(12, 80)
$btnRun.Size     = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnRun)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Reports Folder'
$btnOpen.Location = New-Object System.Drawing.Point(140, 80)
$btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnOpen)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(622, 80)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 154)
$lblStatus.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 174 -Width 720 -Height 12

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Location   = New-Object System.Drawing.Point(12, 192)
$txtOut.Size       = New-Object System.Drawing.Size(720, 406)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'ExternalSharing' `
                -X 12 -Y 116 -Width 720 -StatusLabel $lblStatus

$btnOpen.Add_Click({ Open-ReportFolder -Path $ReportDir })

$btnRun.Add_Click({
    $btnRun.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = ''
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar `
                        -Activity 'External sharing'
        $report = Get-ExternalSharingReport -Progress $progress -IncludePersonalSites:$chkPersonal.Checked

        Complete-ProgressReporter -Activity 'External sharing' -Bar $pbar

        if (@($report.Rows).Count -eq 0) {
            $lblStatus.Text = 'No sites could be reported -- nothing written.'
            $txtOut.Text = (Format-ReportSummary -Result $report) +
                           [Environment]::NewLine +
                           'No CSV was written because there were no rows.'
            return
        }

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir "ExternalSharing_$stamp.csv"
        $report.Rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $files.SelectPath($csv)
        $lblStatus.Text = "Done. $(@($report.Rows).Count) row(s) written."

        $txtOut.Text = (Format-ReportSummary -Result $report) +
                       [Environment]::NewLine +
                       'Saved to:' + [Environment]::NewLine + "  $csv"

        $note = if ($report.SiteFailed.Count -or $report.CappedRisk -or
                    $report.ExternalCapped.Count -or -not $report.ExternalAvailable) { ' -- PARTIAL' } else { '' }
        # PARTIAL is the keyword the log is scanned for; Result is the same
        # fact in a column a filter can see.
        if ($note) { $logResult = 'Partial' } else { $logResult = 'Success' }
        Write-CompletedLog -Task "External sharing -- $(@($report.Rows).Count) sites$note" -Result $logResult
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'External sharing'
    } finally {
        Complete-ProgressReporter -Activity 'External sharing'
        $form.Cursor    = $prev
        $btnRun.Enabled = $true
    }
})

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'External Sharing' | Out-Null
[void]$form.ShowDialog()
#endregion
