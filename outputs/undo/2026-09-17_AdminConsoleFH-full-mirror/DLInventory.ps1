#Requires -Version 5.1
# TOOLKIT-VARIANT: Cloud
# ---------------------------------------------------------------------------
# CLOUD VERSION, AND NOW THE ONLY VERSION.
#
# AD-DLInventory.ps1 was the on-premises twin. It has been PARKED to
# DisabledTools\, so there is no longer an AD- variant beside this file and
# nothing here needs a domain controller.
#
# The header used to say the AD- variant was "beside it", which stopped being
# true the moment the file moved. A comment that names a sibling is a
# reference like any other -- it goes stale silently, and the next person
# reads it as current.
# -------------------------------------------------------------------------
<#
    DLInventory.ps1
    Inventories mail-enabled groups from Exchange Online for migration and
    clean-up planning, and exports to CSV.

    Cloud-only build. Covers distribution groups, mail-enabled security groups
    and Microsoft 365 groups.

    Reasoning: ScriptNotes/DLInventory.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-03
#>
# TOOLKIT-HELP-BEGIN
# Inventories the mail-enabled groups in on-premises Active Directory, for
# migration planning. Writes a CSV.
#
# Built for the distribution list move to the cloud: what exists, how big it is,
# and who manages it, in one file you can sort and plan against.
#
# Read-only.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress
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

#region Configuration --------------------------------------------------------
$ReportDir = Get-ReportPath

$StaleDays = 365
#endregion

#region Data -----------------------------------------------------------------

function Get-DLInventory {
    <#
        Everything comes from Exchange Online, so there is no on-prem/cloud
        cmdlet collision to avoid -- every call is routed through the cloud
        session explicitly.

        Two cmdlets, because the cloud splits what AD kept in one place:
        Get-DistributionGroup returns classic distribution groups and
        mail-enabled security groups; Get-UnifiedGroup returns Microsoft 365
        groups, which Get-DistributionGroup does not see at all. Leaving the
        second one out would under-report the estate and quietly make the
        migration plan wrong.

        Membership counts cost one call per group. There is no cloud
        equivalent of reading the member attribute off a bulk directory query,
        so this is a scan and is reported as one.
    #>
    param([scriptblock]$Progress)

    $rows       = New-Object System.Collections.ArrayList
    $failed     = @()
    $ownerCache = @{}

    & $Progress 'Querying Exchange Online for distribution groups...' 2

    $groups = @()
    try {
        $groups += @(Invoke-EXOCommand -Name 'Get-DistributionGroup' -Parameters @{
            ResultSize = 'Unlimited'; ErrorAction = 'Stop' })
    } catch {
        $failed += "Distribution groups ($($_.Exception.Message))"
    }

    & $Progress 'Querying Exchange Online for Microsoft 365 groups...' 5
    $unified = @()
    try {
        $unified = @(Invoke-EXOCommand -Name 'Get-UnifiedGroup' -Parameters @{
            ResultSize = 'Unlimited'; ErrorAction = 'Stop' })
    } catch {
        $failed += "Microsoft 365 groups ($($_.Exception.Message))"
    }

    if ($groups.Count -eq 0 -and $unified.Count -eq 0 -and $failed.Count -gt 0) {
        throw "Could not enumerate any mail-enabled groups: $($failed -join '; ')"
    }

    $total = $groups.Count + $unified.Count
    $n = 0

    function Resolve-OwnerList {
        param($ManagedBy, [hashtable]$Cache)
        $names = @()
        foreach ($m in @($ManagedBy)) {
            $key = "$m"
            if ($key -eq '') { continue }
            if ($Cache.ContainsKey($key)) { $names += $Cache[$key]; continue }
            $resolved = $key
            try {
                $r = Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{
                    Identity = $key; ErrorAction = 'Stop' }
                if ($r.PrimarySmtpAddress) { $resolved = "$($r.PrimarySmtpAddress)" }
            } catch { }
            $Cache[$key] = $resolved
            $names += $resolved
        }
        return ($names -join '; ')
    }

    foreach ($g in $groups) {
        $n++
        if ($n % 10 -eq 0 -or $n -eq $total) {
            $pct = [int](5 + 93.0 * ($n / [Math]::Max(1, $total)))
            & $Progress "$n of $total groups..." $pct
        }

        $direct = 0
        $nested = 0
        try {
            $members = @(Invoke-EXOCommand -Name 'Get-DistributionGroupMember' -Parameters @{
                Identity = "$($g.Identity)"; ResultSize = 'Unlimited'; ErrorAction = 'Stop' })
            $direct = $members.Count
            foreach ($m in $members) {
                if ("$($m.RecipientType)" -match 'Group') { $nested++ }
            }
        } catch {
            $failed += "Members of $($g.PrimarySmtpAddress) ($($_.Exception.Message))"
            $direct = $null
        }

        $isSecurity  = ("$($g.RecipientTypeDetails)" -eq 'MailUniversalSecurityGroup')
        $changedDays = if ($g.WhenChanged) { [int]((Get-Date) - $g.WhenChanged).TotalDays } else { $null }

        [void]$rows.Add([PSCustomObject][ordered]@{
            Name                = "$($g.Name)"
            DisplayName         = "$($g.DisplayName)"
            PrimarySmtpAddress  = "$($g.PrimarySmtpAddress)"
            GroupType           = if ($isSecurity) { 'MailSecurity' } else { 'Distribution' }
            GroupCategory       = if ($isSecurity) { 'Security' } else { 'Distribution' }
            RecipientTypeDetails = "$($g.RecipientTypeDetails)"
            DirectMembers       = $direct
            NestedGroups        = $nested
            IsEmpty             = ($direct -eq 0)
            IsSingleMember      = ($direct -eq 1)
            HasOwner            = ([string]::IsNullOrWhiteSpace("$(@($g.ManagedBy) -join '')") -eq $false)
            ManagedBy           = (Resolve-OwnerList -ManagedBy $g.ManagedBy -Cache $ownerCache)
            MailEnabledSecurity = $isSecurity
            HiddenFromGAL       = [bool]$g.HiddenFromAddressListsEnabled
            RequireSenderAuth   = [bool]$g.RequireSenderAuthenticationEnabled
            IsDirSynced         = [bool]$g.IsDirSynced
            WhenCreated         = if ($g.WhenCreated) { $g.WhenCreated.ToString('yyyy-MM-dd') } else { '' }
            WhenChanged         = if ($g.WhenChanged) { $g.WhenChanged.ToString('yyyy-MM-dd') } else { '' }
            DaysSinceChange     = $changedDays
            LooksStale          = ($null -ne $changedDays -and $changedDays -gt $StaleDays)
            Description         = ((@($g.Description) -join ' ') -replace '\s+', ' ').Trim()
        })
    }

    foreach ($u in $unified) {
        $n++
        if ($n % 10 -eq 0 -or $n -eq $total) {
            $pct = [int](5 + 93.0 * ($n / [Math]::Max(1, $total)))
            & $Progress "$n of $total groups..." $pct
        }

        $changedDays = if ($u.WhenChanged) { [int]((Get-Date) - $u.WhenChanged).TotalDays } else { $null }

        [void]$rows.Add([PSCustomObject][ordered]@{
            Name                = "$($u.Name)"
            DisplayName         = "$($u.DisplayName)"
            PrimarySmtpAddress  = "$($u.PrimarySmtpAddress)"
            GroupType           = 'Microsoft365'
            GroupCategory       = 'Microsoft365'
            RecipientTypeDetails = "$($u.RecipientTypeDetails)"
            DirectMembers       = [int]$u.GroupMemberCount
            NestedGroups        = 0
            IsEmpty             = ([int]$u.GroupMemberCount -eq 0)
            IsSingleMember      = ([int]$u.GroupMemberCount -eq 1)
            HasOwner            = ([int]$u.GroupExternalMemberCount -ge 0 -and [int]$u.ManagedByDetails.Count -gt 0)
            ManagedBy           = ((@($u.ManagedByDetails) | ForEach-Object { "$_" }) -join '; ')
            MailEnabledSecurity = $false
            HiddenFromGAL       = [bool]$u.HiddenFromAddressListsEnabled
            RequireSenderAuth   = [bool]$u.RequireSenderAuthenticationEnabled
            IsDirSynced         = [bool]$u.IsDirSynced
            WhenCreated         = if ($u.WhenCreated) { $u.WhenCreated.ToString('yyyy-MM-dd') } else { '' }
            WhenChanged         = if ($u.WhenChanged) { $u.WhenChanged.ToString('yyyy-MM-dd') } else { '' }
            DaysSinceChange     = $changedDays
            LooksStale          = ($null -ne $changedDays -and $changedDays -gt $StaleDays)
            Description         = ((@($u.Notes) -join ' ') -replace '\s+', ' ').Trim()
        })
    }

    $sorted = @($rows | Sort-Object `
        @{ Expression = { $_.NestedGroups }; Descending = $true },
        @{ Expression = { $_.HasOwner } },
        @{ Expression = { $_.DirectMembers }; Descending = $true })

    return [PSCustomObject]@{ Rows = $sorted; Failed = $failed }
}

function Format-ReportSummary {
    param($Result)
    $rows = @($Result.Rows)
    $sb = New-Object System.Text.StringBuilder

    [void]$sb.AppendLine("Mail-enabled groups found : $($rows.Count)")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('By type:')
    $rows | Group-Object GroupType | Sort-Object Name | ForEach-Object {
        [void]$sb.AppendLine(("  {0,-16}{1,6}" -f $_.Name, $_.Count))
    }

    $empty   = @($rows | Where-Object { $_.IsEmpty })
    $single  = @($rows | Where-Object { $_.IsSingleMember })
    $noOwner = @($rows | Where-Object { -not $_.HasOwner })
    $nested  = @($rows | Where-Object { $_.NestedGroups -gt 0 })
    $stale   = @($rows | Where-Object { $_.LooksStale })
    $hidden  = @($rows | Where-Object { $_.HiddenFromGAL })
    $mesg    = @($rows | Where-Object { $_.MailEnabledSecurity })
    $synced  = @($rows | Where-Object { $_.IsDirSynced })

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('--- Migration planning ---')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("Retire rather than migrate:")
    [void]$sb.AppendLine(("  Empty (no members)          {0,6}" -f $empty.Count))
    [void]$sb.AppendLine(("  Single member only          {0,6}" -f $single.Count))
    [void]$sb.AppendLine(("  Unchanged > $StaleDays days        {0,6}" -f $stale.Count))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Need a decision before migrating:')
    [void]$sb.AppendLine(("  No owner (ManagedBy empty)  {0,6}   <-- hardest: nobody to ask" -f $noOwner.Count))
    [void]$sb.AppendLine(("  Contain nested groups       {0,6}" -f $nested.Count))
    [void]$sb.AppendLine(("  Mail-enabled security       {0,6}   membership AND permissions" -f $mesg.Count))
    [void]$sb.AppendLine(("  Hidden from the GAL         {0,6}" -f $hidden.Count))
    [void]$sb.AppendLine(("  Directory-synced            {0,6}   <-- cannot be edited in the cloud" -f $synced.Count))

    $clean = @($rows | Where-Object {
        -not $_.IsEmpty -and -not $_.IsSingleMember -and $_.HasOwner -and
        $_.NestedGroups -eq 0 -and -not $_.MailEnabledSecurity })
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine(("Straightforward migrations   {0,6}" -f $clean.Count))

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Note on staleness: DaysSinceChange is when the OBJECT last changed,')
    [void]$sb.AppendLine('  not when the group was last EMAILED. Exchange holds no last-used')
    [void]$sb.AppendLine('  data and message trace only reaches back ~10 days, so a group can')
    [void]$sb.AppendLine('  look stale and still be in daily use. Confirm before retiring one.')
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Note on directory-synced groups: a group with IsDirSynced True is')
    [void]$sb.AppendLine('  mastered on-premises. It appears here because it is mail-enabled in')
    [void]$sb.AppendLine('  the cloud, but membership and most properties can only be changed at')
    [void]$sb.AppendLine('  the source. In a genuinely cloud-only tenant this count should be 0;')
    [void]$sb.AppendLine('  anything else means a directory sync is still running somewhere.')

    if ($Result.Failed -and $Result.Failed.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('*** INCOMPLETE -- these could not be read: ***')
        foreach ($f in $Result.Failed) { [void]$sb.AppendLine("  $f") }
    }
    return $sb.ToString()
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Report -- Distribution List Inventory'
$form.Size            = New-Object System.Drawing.Size(760, 676)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Inventories mail-enabled groups from Exchange Online, with migration-readiness indicators.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblInfo)

$lblSrc           = New-Object System.Windows.Forms.Label
$lblSrc.Text      = 'Source: Exchange Online -- distribution groups, mail-enabled security groups and Microsoft 365 groups. One membership call per group, so a large tenant takes a while.'
$lblSrc.Location  = New-Object System.Drawing.Point(12, 32)
$lblSrc.Size      = New-Object System.Drawing.Size(720, 32)
$lblSrc.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblSrc)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run Report'
$btnRun.Location = New-Object System.Drawing.Point(12, 72)
$btnRun.Size     = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnRun)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Reports Folder'
$btnOpen.Location = New-Object System.Drawing.Point(140, 72)
$btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnOpen)


$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(622, 72)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 148)
$lblStatus.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 168 -Width 720 -Height 12

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Location   = New-Object System.Drawing.Point(12, 186)
$txtOut.Size       = New-Object System.Drawing.Size(720, 428)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'DLInventory' `
                -X 12 -Y 108 -Width 720 -StatusLabel $lblStatus


$btnOpen.Add_Click({ Open-ReportFolder -Path $ReportDir })

$btnRun.Add_Click({
    $btnRun.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = ''
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar `
                        -Activity 'DL inventory'
        $report = Get-DLInventory -Progress $progress

        Complete-ProgressReporter -Activity 'DL inventory' -Bar $pbar

        if (@($report.Rows).Count -eq 0) {
            $lblStatus.Text = 'No mail-enabled groups found -- nothing written.'
            $txtOut.Text = (Format-ReportSummary -Result $report) +
                           [Environment]::NewLine +
                           'No CSV was written because there were no rows.'
            return
        }

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir "DLInventory_$stamp.csv"
        $report.Rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $files.SelectPath($csv)
        $lblStatus.Text  = "Done. $(@($report.Rows).Count) group(s) written."

        $txtOut.Text = (Format-ReportSummary -Result $report) +
                       [Environment]::NewLine +
                       'Saved to:' + [Environment]::NewLine + "  $csv"

        $note = if ($report.Failed.Count) { ' -- PARTIAL' } else { '' }
        # PARTIAL is the keyword the log is scanned for; Result is the same
        # fact in a column a filter can see.
        if ($note) { $logResult = 'Partial' } else { $logResult = 'Success' }
        Write-CompletedLog -Task "DL inventory -- $(@($report.Rows).Count) groups$note" -Result $logResult
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'DL inventory'
    } finally {
        Complete-ProgressReporter -Activity 'DL inventory'
        $form.Cursor    = $prev
        $btnRun.Enabled = $true
    }
})

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'DL Inventory' | Out-Null
[void]$form.ShowDialog()
#endregion
