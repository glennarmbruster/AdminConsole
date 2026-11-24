#Requires -Version 5.1
<#
    Inventory-AllGroups.ps1
    Every mail-enabled group in the forest, three fields, written to CSV.

    THE STARTING DATASET FOR THE SOURCE-ANCHOR PROJECT. Not an answer -- a
    source. Slice it in Excel, save the slices, and the scripts that come next
    run against those files rather than re-querying the directory.

    WHY EVERYTHING RATHER THAN A SCOPED SLICE. The first version of this
    filtered to one directory and enriched every row with member counts, nested
    group detection and resolved owners. Exchange has to search the whole
    forest to answer Get-DistributionGroup regardless, so the filtering saved
    nothing at query time -- it only threw away rows that had already been
    fetched, and cost an AD read per group for the enrichment.

    So: one pass, no filter, no Active Directory calls at all. The expensive
    part was never the number of rows returned.

    Read-only. On-premises Exchange only. Exchange Online is never touched --
    these groups are all still source-anchored on-premises, which is the whole
    premise of the project.

    Author   : Glenn Armbruster
    Revised  : 2026-08-13
#>
# TOOLKIT-HELP-BEGIN
# Lists EVERY mail-enabled group in the forest and writes a CSV to the reports
# folder. Three columns: Name, Primary Email, Recipient Type.
#
# This is the starting point for the distribution list move to Microsoft 365.
# It deliberately does not filter and does not decide anything -- it just gets
# the whole estate into one file you can sort, split and plan against.
#
# Both directories are included. Recipient Type is what separates the kinds of
# group:
#
#   MailUniversalDistributionGroup   ordinary distribution group
#   MailUniversalSecurityGroup       mail-enabled security group
#   MailNonUniversalGroup            legacy, from before universal scope
#   RoomList                         room lists
#   DynamicDistributionGroup         membership is a filter, not a list
#
# No member counts, no owners, no nesting. Those cost a directory call per
# group and this is meant to be quick.
#
# Read-only. Nothing is changed by running this.
# TOOLKIT-HELP-END
# TOOLKIT-HELP-RUNS: on-premises Exchange. Active Directory and Exchange Online are not queried.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Three levels up: Projects\<Tenant>\DLMigration\ -> the toolkit root.
# Resolve-Path rather than a literal join so the path in any error message is
# the real one rather than one with '..' in the middle of it.
$ToolkitRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..')).Path
. (Join-Path $ToolkitRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
#
# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost.
if (-not (Test-ConsoleLaunch)) { return }

$ReportDir = Get-ReportPath
$Prefix    = 'Inventory-AllGroups'
$Title     = 'All Mail-Enabled Groups'

#region Data -----------------------------------------------------------------

function Get-AllGroupInventory {
    <#
        Two cmdlets, because Exchange splits what you would think of as one
        thing. Get-DistributionGroup returns distribution groups, mail-enabled
        security groups, legacy non-universal groups and room lists.
        Get-DynamicDistributionGroup returns dynamic groups and is a SEPARATE
        call -- they are not in the first result set at all, and leaving the
        second call out would under-report the estate without any sign that it
        had.

        Every call goes through Invoke-ExchangeCommand -World OnPrem.
        Get-DistributionGroup exists in BOTH Exchanges and the last module
        imported wins the bare name, so an operator with a cloud session open
        would otherwise get a cloud inventory under an on-premises caption.
        Routing it explicitly turns a wrong answer into a refusal.
    #>
    param([scriptblock]$Progress)

    $failed = @()

    # SCOPE FIRST, QUERY SECOND.
    #
    # Exchange's default recipient scope is the DOMAIN hosting Exchange, not
    # the forest. In a multi-domain forest that means this returns a different
    # set depending on what else has run in this process -- anything that
    # called Set-ADSessionScope has already widened it, anything that has not
    # leaves it narrow. A report whose row count depends on which button you
    # pressed first cannot be reconciled against anything.
    #
    # Called with no candidate controller, so it goes forest-wide and pins
    # nothing. Pinning would be actively wrong here: this report deliberately
    # wants every domain, and a pinned controller is a controller in one of
    # them.
    #
    # A failure is recorded rather than thrown. The report still runs and comes
    # back marked PARTIAL saying the scope could not be confirmed. Continuing
    # silently on an unknown scope is the one outcome that must not happen.
    & $Progress 'Setting the directory scope to the whole forest...' 2
    try {
        $scope = Set-ADSessionScope
        if ($scope.Fallback) { $failed += "Scope note: $($scope.Detail)" }
    } catch {
        $failed += ("Could not set the recipient scope, so this list may be missing groups " +
                    "from other domains in the forest ($($_.Exception.Message))")
    }

    $rows = New-Object System.Collections.Generic.List[object]

    # -ResultSize Unlimited IS LOAD-BEARING. The default is 1000 and it
    # truncates SILENTLY -- no warning, no error, just a short list that looks
    # complete. On an estate this size that is not a hypothetical.
    & $Progress 'Querying on-premises Exchange for mail-enabled groups...' 10
    try {
        foreach ($g in @(Invoke-ExchangeCommand -World OnPrem -Command 'Get-DistributionGroup' -Params @{
                    ResultSize = 'Unlimited'; ErrorAction = 'Stop' })) {
            $rows.Add([PSCustomObject][ordered]@{
                Name                 = "$($g.Name)"
                PrimarySmtpAddress   = "$($g.PrimarySmtpAddress)"
                RecipientTypeDetails = "$($g.RecipientTypeDetails)"
            })
        }
    } catch {
        $failed += "Distribution groups ($($_.Exception.Message))"
    }

    & $Progress "Querying dynamic distribution groups ($($rows.Count) so far)..." 70
    try {
        foreach ($d in @(Invoke-ExchangeCommand -World OnPrem -Command 'Get-DynamicDistributionGroup' -Params @{
                    ResultSize = 'Unlimited'; ErrorAction = 'Stop' })) {
            # RecipientTypeDetails is stamped rather than read. This cmdlet has
            # no -RecipientTypeDetails parameter and everything it returns is a
            # dynamic group by definition.
            $rows.Add([PSCustomObject][ordered]@{
                Name                 = "$($d.Name)"
                PrimarySmtpAddress   = "$($d.PrimarySmtpAddress)"
                RecipientTypeDetails = 'DynamicDistributionGroup'
            })
        }
    } catch {
        $failed += "Dynamic distribution groups ($($_.Exception.Message))"
    }

    if ($rows.Count -eq 0 -and $failed.Count -gt 0) {
        throw "Could not read any groups: $($failed -join '; ')"
    }

    & $Progress "Sorting $($rows.Count) group(s)..." 95

    return [PSCustomObject]@{
        Rows   = @($rows | Sort-Object Name)
        Failed = $failed
    }
}

#endregion

#region Window ---------------------------------------------------------------

$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "Report -- $Title"
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.ClientSize      = New-Object System.Drawing.Size(1004, 720)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Every mail-enabled group in the forest. Name, primary email, and what kind of group it is.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(980, 18)
$form.Controls.Add($lblInfo)

$lblSrc           = New-Object System.Windows.Forms.Label
$lblSrc.Text      = 'Source: on-premises Exchange. No Active Directory calls, no Exchange Online.'
$lblSrc.Location  = New-Object System.Drawing.Point(12, 32)
$lblSrc.Size      = New-Object System.Drawing.Size(980, 18)
$lblSrc.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblSrc)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run Report'
$btnRun.Location = New-Object System.Drawing.Point(12, 58)
$btnRun.Size     = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnRun)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Reports Folder'
$btnOpen.Location = New-Object System.Drawing.Point(140, 58)
$btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnOpen)

$lblFind          = New-Object System.Windows.Forms.Label
$lblFind.Text     = 'Find:'
$lblFind.Location = New-Object System.Drawing.Point(310, 64)
$lblFind.Size     = New-Object System.Drawing.Size(34, 18)
$form.Controls.Add($lblFind)

$txtFind          = New-Object System.Windows.Forms.TextBox
$txtFind.Location = New-Object System.Drawing.Point(346, 61)
$txtFind.Size     = New-Object System.Drawing.Size(240, 22)
$form.Controls.Add($txtFind)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(882, 58)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Anchor   = 'Top,Right'
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle -- click Run Report.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 132)
$lblStatus.Size     = New-Object System.Drawing.Size(980, 18)
$lblStatus.Anchor   = 'Top,Left,Right'
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 152 -Width 980 -Height 12

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location                = New-Object System.Drawing.Point(12, 172)
$grid.Size                    = New-Object System.Drawing.Size(980, 500)
$grid.Anchor                  = 'Top,Left,Right,Bottom'
$grid.ReadOnly                = $true
$grid.AllowUserToAddRows      = $false
$grid.AllowUserToDeleteRows   = $false
$grid.AllowUserToOrderColumns = $true
$grid.RowHeadersVisible       = $false
$grid.SelectionMode           = 'FullRowSelect'
$grid.MultiSelect             = $true
$grid.AutoSizeColumnsMode     = 'None'
$form.Controls.Add($grid)

$lblCount          = New-Object System.Windows.Forms.Label
$lblCount.Location = New-Object System.Drawing.Point(12, 684)
$lblCount.Size     = New-Object System.Drawing.Size(980, 20)
$lblCount.Anchor   = 'Left,Right,Bottom'
$form.Controls.Add($lblCount)

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix $Prefix `
                -X 12 -Y 94 -Width 980 -StatusLabel $lblStatus

# A DataTable BOUND TO THE GRID, NOT ROWS ADDED IN A LOOP.
#
# DataGridView.Rows.Add() is a per-row round trip through the control, and this
# report can return five figures. Binding hands the whole set over once, and
# filtering becomes DefaultView.RowFilter -- which the DataTable does in memory
# without touching the control at all.
#
# The alternative, rebuilding the grid on every keystroke in the Find box, is
# the version that makes a report feel broken on a large estate.
$table = New-Object System.Data.DataTable
[void]$table.Columns.Add('Name',                 [string])
[void]$table.Columns.Add('PrimarySmtpAddress',   [string])
[void]$table.Columns.Add('RecipientTypeDetails', [string])
$grid.DataSource = $table

$grid.Columns['Name'].HeaderText                 = 'Name'
$grid.Columns['Name'].Width                      = 320
$grid.Columns['PrimarySmtpAddress'].HeaderText   = 'Primary Email'
$grid.Columns['PrimarySmtpAddress'].Width        = 340
$grid.Columns['RecipientTypeDetails'].HeaderText = 'Recipient Type'
$grid.Columns['RecipientTypeDetails'].Width      = 260

function Update-RowFilter {
    <#
        Doubling the single quotes is REQUIRED, not defensive.

        RowFilter is an expression language, and an apostrophe in the search
        text -- a group called O'Brien's Team, or a typo -- terminates the
        string literal early and throws a syntax error on a keystroke. Doubling
        is how a DataTable expression escapes a quote.
    #>
    $find = "$($txtFind.Text)".Trim()
    if (-not $find) {
        $table.DefaultView.RowFilter = ''
    } else {
        $safe = $find.Replace("'", "''")
        $table.DefaultView.RowFilter =
            "Name LIKE '%$safe%' OR PrimarySmtpAddress LIKE '%$safe%' OR RecipientTypeDetails LIKE '%$safe%'"
    }

    $shown = $table.DefaultView.Count
    $total = $table.Rows.Count
    $lblCount.Text = if ($shown -eq $total) {
        "$total group(s)."
    } else {
        "$shown of $total group(s) shown -- filtered by '$find'."
    }
}

$txtFind.Add_TextChanged({
    try { Update-RowFilter } catch { $lblCount.Text = "Filter error: $($_.Exception.Message)" }
})

$btnOpen.Add_Click({ Open-ReportFolder -Path $ReportDir })

$btnRun.Add_Click({
    $btnRun.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity $Title
        $result   = Get-AllGroupInventory -Progress $progress
        Complete-ProgressReporter -Activity $Title -Bar $pbar

        $rows = @($result.Rows)

        $table.BeginLoadData()
        $table.Rows.Clear()
        foreach ($r in $rows) {
            [void]$table.Rows.Add($r.Name, $r.PrimarySmtpAddress, $r.RecipientTypeDetails)
        }
        $table.EndLoadData()
        Update-RowFilter

        $partial = (@($result.Failed).Count -gt 0)

        if ($rows.Count -eq 0) {
            $lblStatus.Text = 'No groups found -- nothing written.'
            return
        }

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir "$($Prefix)_$stamp.csv"
        $rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $files.SelectPath($csv)

        # A per-type breakdown on the status line, because the first question
        # after "how many" is always "how many of each".
        $byType = ($rows | Group-Object RecipientTypeDetails | Sort-Object Name |
                    ForEach-Object { "$($_.Name) $($_.Count)" }) -join ',  '

        $suffix = if ($partial) { '  ***PARTIAL***' } else { '' }
        $lblStatus.Text = "Done. $($rows.Count) group(s) -> $(Split-Path -Leaf $csv)$suffix"
        $lblCount.Text  = "$($rows.Count) group(s).    $byType"

        if ($partial) {
            [System.Windows.Forms.MessageBox]::Show(
                ("The list is INCOMPLETE. The following did not answer:" +
                 [Environment]::NewLine + [Environment]::NewLine +
                 ($result.Failed -join [Environment]::NewLine)),
                'Partial result', 'OK', 'Warning') | Out-Null
        }

        $logResult = if ($partial) { 'Partial' } else { 'Success' }
        Write-CompletedLog -Action 'Report' -Target $Prefix -Result $logResult `
            -Detail "$Title -- $($rows.Count) group(s) written to $(Split-Path -Leaf $csv)"
    } catch {
        $lblStatus.Text = 'Failed.'
        [System.Windows.Forms.MessageBox]::Show(
            (Format-ErrorDetail -ErrorRecord $_ -Context $Title),
            'Failed', 'OK', 'Error') | Out-Null
        Write-FailedLog -Action 'Report' -Target $Prefix -ErrorRecord $_
    } finally {
        Complete-ProgressReporter -Activity $Title
        $form.Cursor    = $prev
        $btnRun.Enabled = $true
    }
})

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
# AFTER the size is final -- Set-FormResizable takes the current size as the
# minimum, so calling it earlier would pin the window to whatever it was
# part-way through layout.
Set-FormResizable -Form $form
Set-ConsoleTheme -Form $form
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title $Title | Out-Null

[void]$form.ShowDialog()
#endregion
