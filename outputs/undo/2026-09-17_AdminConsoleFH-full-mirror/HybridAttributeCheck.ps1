#Requires -Version 5.1
# TOOLKIT-VARIANT: Dual
#
# RUNS AGAINST ON-PREMISES AND EXCHANGE ONLINE from one window in the
# console's own process. Nothing here is called unqualified: every
# on-premises call goes through Invoke-OnPremCommand and every cloud call
# through Invoke-EXOCommand, so "which Exchange answered" is decided by the
# code and not by import order. Connect both before starting; the window
# says which one is missing.
<#
    HybridAttributeCheck.ps1
    Compares the on-premises remote mailbox stub with the cloud mailbox, attribute by attribute, and says where they disagree.

    Author   : Glenn Armbruster
    Revised  : 2026-09-12
#>
# TOOLKIT-HELP-RUNS: reads Get-RemoteMailbox on-premises and Get-Mailbox / Get-EXOMailbox in the cloud for each identity you list. Changes nothing.
# TOOLKIT-HELP-BEGIN
# Side-by-side check of the attributes a hybrid mailbox needs to agree on.
#
# In a hybrid the on-premises remote mailbox stub and the Exchange Online
# mailbox are two objects describing one person. When they drift apart mail
# usually keeps flowing, so nobody notices until a move fails, an archive
# stops being reachable, or a reply to an old message bounces. This tool puts
# the two objects next to each other and marks every line OK, MISMATCH or
# MISSING.
#
# Paste one or more identities (address or alias, one per line), or let the
# console fill in the selected mailbox, and click Check. For each identity:
#
#   * Exists              -- is there a stub on-premises, is there a mailbox
#                            in the cloud
#   * ExchangeGuid        -- must match, or the mailbox cannot be moved
#   * ArchiveGuid         -- must match when there is an archive
#   * RemoteRoutingAddress-- the stub's routing address must be one of the
#                            cloud mailbox's addresses, or on-premises mail
#                            for this person goes nowhere
#   * PrimarySmtpAddress  -- the two primaries should be the same
#   * EmailAddresses      -- smtp proxy addresses only on one side are listed
#   * RecipientTypeDetails-- RemoteUserMailbox pairs with UserMailbox, and so
#                            on for Shared, Room and Equipment
#
# A read that fails for a reason other than "not found" is an ERROR row with
# the message, so a dropped session never reads as a missing mailbox.
#
# Read-only. Needs both Exchange Online and on-premises Exchange. Export
# writes the grid to CSV.
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
$ReportPrefix = 'HybridAttributeCheck'

# Grid column order, shared by the grid and the CSV.
$GridColumns = @('Identity', 'Check', 'OnPrem', 'Cloud', 'Status')
$GridWidths  = @{ Identity = 190; Check = 150; OnPrem = 260; Cloud = 260; Status = 80 }

# The properties the cloud read asks for. Get-EXOMailbox returns a minimum
# set unless told otherwise, and asking for a property that is already in
# the minimum set costs nothing.
$CloudProperties = @(
    'DisplayName', 'PrimarySmtpAddress', 'EmailAddresses', 'ExchangeGuid',
    'ArchiveGuid', 'RecipientTypeDetails'
)

# File-scope state. Kept in $script: at file scope only -- never read from
# anything Common.ps1 owns.
$script:Rows = @()      # every result row from the last run

# Row colours. Falls back to plain red when the theme does not define one.
$script:ColorBad = $null
try {
    $theme = Get-ConsoleTheme
    if ($theme) { $script:ColorBad = ConvertTo-ThemeColor "$($theme.DangerColor)" }
} catch { }
if (-not $script:ColorBad) { $script:ColorBad = [System.Drawing.Color]::FromArgb(214, 69, 69) }
#endregion

#region Data -----------------------------------------------------------------
function Test-NotFoundError {
    <#
        Is this exception Exchange saying "no such object", as opposed to
        "I could not answer"? The two must produce different rows: MISSING
        is a finding, ERROR is a reason to reconnect and run again.
    #>
    param($ErrorRecord)
    $m = "$($ErrorRecord.Exception.Message)"
    return [bool]($m -match '(?i)(couldn.t be found|cannot be found|could not be found|wasn.t found|was not found|not found)')
}

function Get-SmtpAddressSet {
    <#
        The smtp: proxy addresses of a recipient as a lower-cased HashSet.
        Only smtp entries -- X500, SIP and SPO addresses are not mail routing
        and would make every comparison a mismatch.
    #>
    param($Object)
    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    if ($null -eq $Object) { return $set }
    foreach ($a in @($Object.EmailAddresses)) {
        $t = "$a"
        if ($t -match '^(?i)smtp:(.+)$') {
            $addr = $Matches[1].Trim()
            if ($addr) { [void]$set.Add($addr) }
        }
    }
    return $set
}

function ConvertTo-GuidText {
    param($Value)
    $g = "$Value".Trim().ToLower()
    if (-not $g -or $g -eq [Guid]::Empty.ToString()) { return '' }
    return $g
}

function Get-OnPremStub {
    <#
        One remote mailbox from on-premises Exchange, or $null when there is
        none. Anything other than "not found" is rethrown -- the caller turns
        it into an ERROR row.
    #>
    param([string]$Identity)
    try {
        $r = @(Invoke-OnPremCommand -Name 'Get-RemoteMailbox' -Parameters @{
            Identity = $Identity; ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' })
    } catch {
        if (Test-NotFoundError -ErrorRecord $_) { return $null }
        throw
    }
    if ($r.Count -eq 0) { return $null }
    if ($r.Count -gt 1) { throw "'$Identity' matched $($r.Count) remote mailboxes on-premises. Use the full address." }
    return $r[0]
}

function Get-CloudMailbox {
    <#
        One mailbox from Exchange Online, or $null when there is none.

        Get-EXOMailbox when the module offers it (REST, no session to keep
        warm), Get-Mailbox otherwise. Both go through Invoke-EXOCommand
        because Get-Mailbox exists on-premises too, and this tool holds both
        sessions at once.
    #>
    param([string]$Identity)
    $useExo = [bool](Get-Command -Name 'Get-EXOMailbox' -ErrorAction SilentlyContinue)
    try {
        if ($useExo) {
            $r = @(Invoke-EXOCommand -Name 'Get-EXOMailbox' -Parameters @{
                Identity = $Identity; Properties = $CloudProperties
                ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' })
        } else {
            $r = @(Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
                Identity = $Identity; ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' })
        }
    } catch {
        if (Test-NotFoundError -ErrorRecord $_) { return $null }
        throw
    }
    if ($r.Count -eq 0) { return $null }
    if ($r.Count -gt 1) { throw "'$Identity' matched $($r.Count) mailboxes in the cloud. Use the full address." }
    return $r[0]
}

function New-CheckRow {
    param([string]$Identity, [string]$Check, [string]$OnPrem, [string]$Cloud, [string]$Status)
    return [PSCustomObject][ordered]@{
        Identity = $Identity
        Check    = $Check
        OnPrem   = $OnPrem
        Cloud    = $Cloud
        Status   = $Status
    }
}

function Compare-HybridMailbox {
    <#
        Every check for one identity, as rows. Reads both sides first; a
        failure on either side is one ERROR row and nothing else for that
        identity, because comparing against a read that failed is guessing.
    #>
    param([string]$Identity)

    $rows = New-Object System.Collections.Generic.List[object]

    $stub = $null; $mbx = $null
    try { $stub = Get-OnPremStub -Identity $Identity }
    catch { $rows.Add((New-CheckRow $Identity 'Exists' "ERROR: $($_.Exception.Message)" '' 'ERROR')); return @($rows) }
    try { $mbx = Get-CloudMailbox -Identity $Identity }
    catch { $rows.Add((New-CheckRow $Identity 'Exists' '' "ERROR: $($_.Exception.Message)" 'ERROR')); return @($rows) }

    # --- Exists ---
    $opText = if ($stub) { "RemoteMailbox: $($stub.PrimarySmtpAddress)" } else { '(no remote mailbox on-premises)' }
    $clText = if ($mbx)  { "Mailbox: $($mbx.PrimarySmtpAddress)" }        else { '(no mailbox in the cloud)' }
    $status = if ($stub -and $mbx) { 'OK' } else { 'MISSING' }
    $rows.Add((New-CheckRow $Identity 'Exists' $opText $clText $status))
    if (-not ($stub -and $mbx)) {
        # One side only: show the facts that side has, so the operator can
        # see what exists without a second lookup. No comparison is possible.
        $side = if ($stub) { $stub } else { $mbx }
        $op = if ($stub) { "$($side.ExchangeGuid)" } else { '' }
        $cl = if ($mbx)  { "$($side.ExchangeGuid)" } else { '' }
        $rows.Add((New-CheckRow $Identity 'ExchangeGuid' $op $cl 'MISSING'))
        $op = if ($stub) { "$($side.RecipientTypeDetails)" } else { '' }
        $cl = if ($mbx)  { "$($side.RecipientTypeDetails)" } else { '' }
        $rows.Add((New-CheckRow $Identity 'RecipientTypeDetails' $op $cl 'MISSING'))
        return @($rows)
    }

    # --- ExchangeGuid ---
    $opGuid = ConvertTo-GuidText $stub.ExchangeGuid
    $clGuid = ConvertTo-GuidText $mbx.ExchangeGuid
    $status = if (-not $opGuid -or -not $clGuid) { 'MISSING' } elseif ($opGuid -eq $clGuid) { 'OK' } else { 'MISMATCH' }
    $rows.Add((New-CheckRow $Identity 'ExchangeGuid' $(if ($opGuid) { $opGuid } else { '(empty)' }) $(if ($clGuid) { $clGuid } else { '(empty)' }) $status))

    # --- ArchiveGuid ---
    $opArc = ConvertTo-GuidText $stub.ArchiveGuid
    $clArc = ConvertTo-GuidText $mbx.ArchiveGuid
    if (-not $opArc -and -not $clArc) {
        $rows.Add((New-CheckRow $Identity 'ArchiveGuid' '(no archive)' '(no archive)' 'OK'))
    } elseif ($opArc -and $clArc) {
        $status = if ($opArc -eq $clArc) { 'OK' } else { 'MISMATCH' }
        $rows.Add((New-CheckRow $Identity 'ArchiveGuid' $opArc $clArc $status))
    } else {
        # An archive one side knows about and the other does not. The cloud
        # side having one the stub lacks is the common shape after an archive
        # is enabled online without the stub being updated.
        $rows.Add((New-CheckRow $Identity 'ArchiveGuid' $(if ($opArc) { $opArc } else { '(empty)' }) $(if ($clArc) { $clArc } else { '(empty)' }) 'MISMATCH'))
    }

    # --- RemoteRoutingAddress ---
    $rra = "$($stub.RemoteRoutingAddress)".Trim()
    if ($rra -match '^(?i)smtp:(.+)$') { $rra = $Matches[1].Trim() }
    $cloudSet = Get-SmtpAddressSet -Object $mbx
    $cloudPrimary = "$($mbx.PrimarySmtpAddress)".Trim()
    if (-not $rra) {
        $rows.Add((New-CheckRow $Identity 'RemoteRoutingAddress' '(empty)' $cloudPrimary 'MISSING'))
    } else {
        # The routing address is where on-premises sends this person's mail.
        # It has to be an address the cloud mailbox actually answers to; the
        # cloud column shows the cloud address in the same domain if there is
        # one, so a near-miss is visible.
        $rraDomain = ''
        if ($rra -match '@(.+)$') { $rraDomain = $Matches[1].ToLower() }
        $sameDomain = @($cloudSet | Where-Object { $_ -match '@(.+)$' -and $Matches[1].ToLower() -eq $rraDomain })
        $clText = if ($sameDomain.Count -gt 0) { ($sameDomain -join '; ') } else { "$cloudPrimary (no address in the routing domain)" }
        $status = if ($cloudSet.Contains($rra)) { 'OK' } else { 'MISMATCH' }
        $rows.Add((New-CheckRow $Identity 'RemoteRoutingAddress' $rra $clText $status))
    }

    # --- PrimarySmtpAddress ---
    $opPrimary = "$($stub.PrimarySmtpAddress)".Trim()
    $status = if (-not $opPrimary -or -not $cloudPrimary) { 'MISSING' } elseif ($opPrimary -ieq $cloudPrimary) { 'OK' } else { 'MISMATCH' }
    $rows.Add((New-CheckRow $Identity 'PrimarySmtpAddress' $opPrimary $cloudPrimary $status))

    # --- EmailAddresses (smtp only) ---
    $opSet = Get-SmtpAddressSet -Object $stub
    $onlyOp = @($opSet    | Where-Object { -not $cloudSet.Contains($_) } | Sort-Object)
    $onlyCl = @($cloudSet | Where-Object { -not $opSet.Contains($_) }    | Sort-Object)
    $opText = "$($opSet.Count) smtp address(es)"
    $clText = "$($cloudSet.Count) smtp address(es)"
    if ($onlyOp.Count -gt 0) { $opText += '; only on-prem: ' + ($onlyOp -join '; ') }
    if ($onlyCl.Count -gt 0) { $clText += '; only cloud: '   + ($onlyCl -join '; ') }
    $status = if ($onlyOp.Count -eq 0 -and $onlyCl.Count -eq 0) { 'OK' } else { 'MISMATCH' }
    $rows.Add((New-CheckRow $Identity 'EmailAddresses (smtp)' $opText $clText $status))

    # --- RecipientTypeDetails ---
    # RemoteUserMailbox <-> UserMailbox, RemoteSharedMailbox <-> SharedMailbox,
    # RemoteRoomMailbox <-> RoomMailbox, RemoteEquipmentMailbox <->
    # EquipmentMailbox. Dropping the 'Remote' prefix is the whole mapping.
    $opType = "$($stub.RecipientTypeDetails)".Trim()
    $clType = "$($mbx.RecipientTypeDetails)".Trim()
    $opBase = $opType -replace '^(?i)Remote', ''
    $status = if (-not $opType -or -not $clType) { 'MISSING' } elseif ($opBase -ieq $clType) { 'OK' } else { 'MISMATCH' }
    $rows.Add((New-CheckRow $Identity 'RecipientTypeDetails' $opType $clType $status))

    return @($rows)
}

function Get-IdentityList {
    <#
        The textbox, one identity per line, trimmed, blanks and duplicates
        dropped. Order is kept so the grid reads in the order typed.
    #>
    param([string]$Text)
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $out  = New-Object System.Collections.Generic.List[string]
    foreach ($line in ("$Text" -split "`r?`n")) {
        $t = "$line".Trim()
        if (-not $t) { continue }
        if ($seen.Add($t)) { $out.Add($t) }
    }
    return @($out)
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Hybrid Attribute Check -- on-premises stub vs cloud mailbox'
$form.Size            = New-Object System.Drawing.Size(980, 760)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Compares Get-RemoteMailbox (on-premises) with the Exchange Online mailbox, attribute by attribute. Read-only.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(940, 18)
$form.Controls.Add($lblInfo)

$lblIds          = New-Object System.Windows.Forms.Label
$lblIds.Text     = 'Identities  (address or alias, one per line):'
$lblIds.Location = New-Object System.Drawing.Point(12, 40)
$lblIds.Size     = New-Object System.Drawing.Size(560, 18)
$form.Controls.Add($lblIds)

$txtIds            = New-Object System.Windows.Forms.TextBox
$txtIds.Multiline  = $true
$txtIds.ScrollBars = 'Vertical'
$txtIds.WordWrap   = $false
$txtIds.AcceptsReturn = $true
$txtIds.Location   = New-Object System.Drawing.Point(12, 60)
$txtIds.Size       = New-Object System.Drawing.Size(560, 110)
$txtIds.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtIds.Text       = "$SmtpAddress"
$form.Controls.Add($txtIds)

$btnCheck          = New-Object System.Windows.Forms.Button
$btnCheck.Text     = 'Check'
$btnCheck.Location = New-Object System.Drawing.Point(588, 60)
$btnCheck.Size     = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnCheck)

$btnExport          = New-Object System.Windows.Forms.Button
$btnExport.Text     = 'Export CSV'
$btnExport.Location = New-Object System.Drawing.Point(588, 98)
$btnExport.Size     = New-Object System.Drawing.Size(120, 30)
$btnExport.Enabled  = $false
$form.Controls.Add($btnExport)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Reports Folder'
$btnOpen.Location = New-Object System.Drawing.Point(588, 136)
$btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnOpen)

$chkShowOk          = New-Object System.Windows.Forms.CheckBox
$chkShowOk.Text     = 'Show OK rows'
$chkShowOk.Checked  = $true
$chkShowOk.Location = New-Object System.Drawing.Point(720, 102)
$chkShowOk.Size     = New-Object System.Drawing.Size(130, 22)
$form.Controls.Add($chkShowOk)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(842, 60)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle. List identities and click Check.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 178)
$lblStatus.Size     = New-Object System.Drawing.Size(940, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 198 -Width 940 -Height 12

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix $ReportPrefix `
                -X 12 -Y 216 -Width 940 -StatusLabel $lblStatus

$lblFilter          = New-Object System.Windows.Forms.Label
$lblFilter.Text     = 'Filter:'
$lblFilter.Location = New-Object System.Drawing.Point(12, 257)
$lblFilter.Size     = New-Object System.Drawing.Size(44, 18)
$form.Controls.Add($lblFilter)

$txtFilter          = New-Object System.Windows.Forms.TextBox
$txtFilter.Location = New-Object System.Drawing.Point(58, 254)
$txtFilter.Size     = New-Object System.Drawing.Size(300, 24)
$form.Controls.Add($txtFilter)

$lblCount           = New-Object System.Windows.Forms.Label
$lblCount.Text      = 'No results yet.'
$lblCount.Location  = New-Object System.Drawing.Point(370, 257)
$lblCount.Size      = New-Object System.Drawing.Size(582, 18)
$lblCount.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblCount)

$grid                        = New-Object System.Windows.Forms.DataGridView
$grid.Location               = New-Object System.Drawing.Point(12, 284)
$grid.Size                   = New-Object System.Drawing.Size(940, 420)
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
#endregion

#region UI helpers -----------------------------------------------------------
function Update-Grid {
    <#
        Repaints the grid from $script:Rows through the filter box and the
        OK checkbox. The rows themselves are never filtered -- Export writes
        all of them.
    #>
    $filter = "$($txtFilter.Text)".Trim()
    $showOk = $chkShowOk.Checked
    $grid.SuspendLayout()
    try {
        $grid.Rows.Clear()
        $shown = 0
        foreach ($r in @($script:Rows)) {
            if (-not $showOk -and $r.Status -eq 'OK') { continue }
            if ($filter) {
                $hit = $false
                foreach ($col in $GridColumns) { if ("$($r.$col)" -like "*$filter*") { $hit = $true; break } }
                if (-not $hit) { continue }
            }
            $values = @()
            foreach ($col in $GridColumns) { $values += "$($r.$col)" }
            $idx = $grid.Rows.Add($values)
            $grid.Rows[$idx].Tag = $r
            if ($r.Status -ne 'OK') {
                $grid.Rows[$idx].DefaultCellStyle.ForeColor = $script:ColorBad
            }
            $shown++
        }
    } finally {
        $grid.ResumeLayout()
    }
    $total = @($script:Rows).Count
    if ($total -eq 0) {
        $lblCount.Text = 'No results yet.'
    } elseif ($shown -eq $total) {
        $lblCount.Text = "$total row(s)."
    } else {
        $lblCount.Text = "$shown of $total row(s) shown."
    }
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

$chkShowOk.Add_CheckedChanged({
    try { Update-Grid }
    catch { Show-HandlerError -ErrorRecord $_ -Context 'Toggle OK rows' }
})

$btnCheck.Add_Click({
    try {
        $ids = @(Get-IdentityList -Text $txtIds.Text)
        if ($ids.Count -eq 0) {
            $lblStatus.Text = 'Type at least one identity, one per line.'
            return
        }

        # BOTH SIDES ARE REQUIRED, AND THE MESSAGE SAYS WHICH ONE IS MISSING.
        # "Not connected" on a tool that needs two connections is a guess.
        $exo = Test-ExoConnectionState
        if (-not $exo.Connected) {
            Show-ToolDialog(
                "Not connected to Exchange Online. Use the Connect button on the console first.`r`n`r`n$($exo.Detail)",
                'Exchange Online required', 'OK', 'Warning') | Out-Null
            return
        }
        $ems = Test-EmsAvailable
        if (-not $ems.Available) {
            Show-ToolDialog(
                "No on-premises Exchange session in this process. Connect OnPrem from the console first.`r`n`r`n$($ems.Detail)",
                'On-premises Exchange required', 'OK', 'Warning') | Out-Null
            return
        }

        # INLINE, NOT Start-ToolJob. Half of every identity is an on-premises
        # read, and an implicit-remoting session cannot be shared with the
        # background worker (see the note above Start-ToolJob in Common.ps1).
        # The progress reporter keeps the window honest instead.
        $btnCheck.Enabled  = $false
        $btnExport.Enabled = $false
        $prev = $form.Cursor
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $script:Rows = @()
        $grid.Rows.Clear()
        try {
            $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity 'Hybrid attribute check'
            $all = New-Object System.Collections.Generic.List[object]
            $i = 0
            foreach ($id in $ids) {
                $i++
                & $progress "Checking $i of $($ids.Count): $id" ([int](100 * ($i - 1) / $ids.Count))
                foreach ($r in @(Compare-HybridMailbox -Identity $id)) { $all.Add($r) }
            }
            Complete-ProgressReporter -Activity 'Hybrid attribute check' -Bar $pbar

            $script:Rows = @($all)
            Update-Grid

            $bad = @($script:Rows | Where-Object { $_.Status -ne 'OK' }).Count
            $lblStatus.Text = "Done. $($ids.Count) identit$(if ($ids.Count -eq 1) { 'y' } else { 'ies' }) checked, $($script:Rows.Count) row(s), $bad not OK."
            $btnExport.Enabled = ($script:Rows.Count -gt 0)
        } finally {
            Complete-ProgressReporter -Activity 'Hybrid attribute check'
            $form.Cursor      = $prev
            $btnCheck.Enabled = $true
        }
    } catch {
        $lblStatus.Text = 'Check failed.'
        Show-HandlerError -ErrorRecord $_ -Context 'Hybrid attribute check'
    }
})

$btnExport.Add_Click({
    try {
        $rows = @($script:Rows)
        if ($rows.Count -eq 0) { $lblStatus.Text = 'Nothing to export. Run a check first.'; return }

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir ("{0}_{1}.csv" -f $ReportPrefix, $stamp)
        $rows | Select-Object -Property $GridColumns | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $files.SelectPath($csv)
        $lblStatus.Text = "Exported $($rows.Count) row(s) to $csv"
        $bad = @($rows | Where-Object { $_.Status -ne 'OK' }).Count
        Write-CompletedLog -Action 'Report' -Target $csv -Result 'Success' `
            -Detail "Hybrid attribute check: $($rows.Count) row(s), $bad not OK."
    } catch {
        $lblStatus.Text = 'Export failed.'
        Show-HandlerError -ErrorRecord $_ -Context 'Export hybrid attribute check'
    }
})
#endregion

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Hybrid Attribute Check' | Out-Null

[void]$form.ShowDialog()
