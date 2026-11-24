#Requires -Version 5.1
<#
    _Shared.ps1  --  helpers common to every step of the SMBMigration project.

    NOT a tool and NOT a fork of Common.ps1. Every step dot-sources the one
    Common.ps1 at the toolkit root first, then this file for the handful of
    things only this project needs: the instruction popup, the newest-export
    picker, and the list join/split used to round-trip multi-valued
    attributes through CSV.

    The underscore keeps it at the top of the folder and out of the numbered
    sequence. It has no TOOLKIT-HELP block on purpose: it is never launched.

    Author   : Glenn Armbruster
    Revised  : 2026-10-10  (SMBMigration: batch files, M-d tag, soft gates)
#>

# Multi-valued attributes travel through CSV as one cell, ';' separated.
# ';' rather than ',' so the CSV stays a CSV; nothing exported here can
# legally contain ';' (SMTP addresses, X500 DNs, GUIDs, names).
$script:ListSep = ';'

function ConvertTo-CsvList {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    return (@($Value | ForEach-Object { "$_" } | Where-Object { $_ -ne '' }) -join $script:ListSep)
}

function ConvertFrom-CsvList {
    param([string]$Text)
    if (-not "$Text".Trim()) { return @() }
    return @("$Text" -split [regex]::Escape($script:ListSep) | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Get-MigrationTimestamp { return (Get-Date -Format 'yyyy-MM-dd_HHmm') }

function Get-NewestExport {
    <#
        The most recent file matching a pattern in the project folder, e.g.
        'Migration-DL-ONPREM_*.csv'. Newest by LastWriteTime, not by name, so
        a file copied back in from Exports-Complete\ still wins if it is the
        latest. Returns $null when there is none; the caller decides whether
        that is fatal.
    #>
    param(
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][string]$Pattern
    )
    return (Get-ChildItem -LiteralPath $Folder -Filter $Pattern -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1)
}

# ---- THE BATCH -----------------------------------------------------------
# Everything in this project hangs off one batch id, 'M-d-NNN' (month, day,
# three-digit number), which Step 1 mints. The input list is
# SMBMigration-<id>.csv; every stage writes SMBMigration-<id>-<Stage>.csv
# beside it. So a batch is one glob and nothing is picked up by "newest".
$script:BatchPrefix = 'SMBMigration-'
$script:TagPattern  = '^SMB-\d{1,2}-\d{1,2}-'

function Get-SmbTag { param([string]$BatchId)
    # The display-name prefix is the batch's month-day: 'SMB-10-10-'.
    if ($BatchId -match '^(\d{1,2}-\d{1,2})-') { return ('SMB-' + $Matches[1] + '-') }
    return ('SMB-' + (Get-Date -Format 'M-d') + '-')
}
function Test-SmbTagged  { param([string]$DisplayName) return ("$DisplayName" -match $script:TagPattern) }
function Remove-SmbTag   { param([string]$DisplayName) return ("$DisplayName" -replace $script:TagPattern, '') }

function Get-SmbBatchFile {
    param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$BatchId, [string]$Stage = '')
    $n = $script:BatchPrefix + $BatchId + $(if ($Stage) { "-$Stage" } else { '' }) + '.csv'
    return (Join-Path $Folder $n)
}

function Get-SmbBatchList {
    # Every batch id with an input file in the folder, newest first.
    param([Parameter(Mandatory)][string]$Folder)
    $ids = @()
    foreach ($f in @(Get-ChildItem -LiteralPath $Folder -Filter ($script:BatchPrefix + '*.csv') -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)) {
        if ($f.BaseName -match ('^' + [regex]::Escape($script:BatchPrefix) + '(\d{1,2}-\d{1,2}-\d{3})$')) { $ids += $Matches[1] }
    }
    return @($ids)
}

function Select-SmbBatch {
    <#
        Which batch is this step working on? A small dialog listing every
        batch input file in the folder, newest first and preselected.
        Returns the id ('10-10-001') or '' when cancelled.
    #>
    param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$Title)
    $ids = @(Get-SmbBatchList -Folder $Folder)
    if ($ids.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No batch file ($($script:BatchPrefix)<M-d-NNN>.csv) exists in $Folder. Run Step 1 first.", $Title, 'OK', 'Warning') | Out-Null
        return ''
    }
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $Title; $f.StartPosition = 'CenterScreen'; $f.FormBorderStyle = 'FixedDialog'
    $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
    $f.ClientSize = New-Object System.Drawing.Size(420, 110)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = 'Batch:'; $l.Location = New-Object System.Drawing.Point(12, 16); $l.Size = New-Object System.Drawing.Size(60, 20)
    $f.Controls.Add($l)
    $c = New-Object System.Windows.Forms.ComboBox
    $c.DropDownStyle = 'DropDownList'; $c.Location = New-Object System.Drawing.Point(76, 12); $c.Size = New-Object System.Drawing.Size(330, 24)
    foreach ($i in $ids) { [void]$c.Items.Add($i) }
    $c.SelectedIndex = 0
    $f.Controls.Add($c)
    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'OK'; $ok.Size = New-Object System.Drawing.Size(90, 28); $ok.Location = New-Object System.Drawing.Point(218, 66)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK; $f.Controls.Add($ok)
    $ca = New-Object System.Windows.Forms.Button
    $ca.Text = 'Cancel'; $ca.Size = New-Object System.Drawing.Size(90, 28); $ca.Location = New-Object System.Drawing.Point(316, 66)
    $ca.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $f.Controls.Add($ca)
    $f.AcceptButton = $ok; $f.CancelButton = $ca
    Set-ConsoleTheme -Form $f
    if ($f.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { Write-Host 'Cancelled.' -ForegroundColor Yellow; return '' }
    return "$($c.SelectedItem)"
}

function Read-SmbBatch {
    # The addresses in the batch input file: column PrimarySmtpAddress,
    # trimmed, de-blanked, unique.
    param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$BatchId)
    $p = Get-SmbBatchFile -Folder $Folder -BatchId $BatchId
    if (-not (Test-Path -LiteralPath $p)) { throw "Batch file not found: $p" }
    $rows = @(Import-Csv -LiteralPath $p)
    $col = 'PrimarySmtpAddress'
    if ($rows.Count -and -not $rows[0].PSObject.Properties[$col]) {
        # Tolerate the old one-column shape.
        $col = @($rows[0].PSObject.Properties.Name)[0]
    }
    $out = @($rows | ForEach-Object { "$($_.$col)".Trim() } | Where-Object { $_ } | Select-Object -Unique)
    if ($out.Count -eq 0) { throw "No addresses in $([System.IO.Path]::GetFileName($p))." }
    return $out
}

function Confirm-SmbGate {
    <#
        A SOFT gate: says what is wrong, offers to go on anyway. Returns
        $true to proceed. Every gate in this project is soft by request --
        the operator knows things the script cannot see.
    #>
    param([Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][string]$Problem)
    $r = [System.Windows.Forms.MessageBox]::Show(
        ($Problem + "`r`n`r`n" + 'Continue anyway?'), $Title,
        [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    if ($r -eq [System.Windows.Forms.DialogResult]::Yes) { Write-Host "OVERRIDE: $Problem" -ForegroundColor Yellow; return $true }
    Write-Host 'Stopped at the gate.' -ForegroundColor Yellow
    return $false
}

function Show-StepInstructions {
    <#
        The instruction popup every step opens with. Returns $true on OK.
        Height is measured from the text rather than guessed (ARCHITECTURE 14l).
    #>
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][string]$ScriptPath
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false; $form.MinimizeBox = $false
    $form.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.AutoSize = $true
    $label.MaximumSize = New-Object System.Drawing.Size(560, 0)
    $label.Location = New-Object System.Drawing.Point(12, 12)
    $form.Controls.Add($label)

    $sz = [System.Windows.Forms.TextRenderer]::MeasureText($Text, $label.Font, (New-Object System.Drawing.Size(560, 0)), [System.Windows.Forms.TextFormatFlags]::WordBreak)
    $btnY = 12 + $sz.Height + 16
    $form.ClientSize = New-Object System.Drawing.Size(584, ($btnY + 40))

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'OK'; $ok.Size = New-Object System.Drawing.Size(90, 28)
    $ok.Location = New-Object System.Drawing.Point(384, $btnY)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($ok)

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'; $cancel.Size = New-Object System.Drawing.Size(90, 28)
    $cancel.Location = New-Object System.Drawing.Point(482, $btnY)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancel)
    $form.AcceptButton = $ok; $form.CancelButton = $cancel

    Set-ConsoleTheme -Form $form
    Add-ToolHelpButton -Form $form -ScriptPath $ScriptPath -Title $Title | Out-Null
    $r = $form.ShowDialog()
    if ($r -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Host 'Cancelled.' -ForegroundColor Yellow
        return $false
    }
    return $true
}

function Write-ExportDone {
    param([Parameter(Mandatory)][string]$Path, [int]$Rows)
    Write-Host ''
    Write-Host "Wrote $Rows row(s): $Path" -ForegroundColor Green
}

function Enter-MigrationOnPremScope {
    <#
        Forest-wide recipient scope for the on-premises steps.

        FOUND 2026-10-04, before the first real run. The Exchange shell's
        default recipient scope is the domain the Exchange server lives in,
        and the jump box's AD module defaults to the domain the box is joined
        to. The objects this project moves live in a CHILD domain, so without
        this every Get-RemoteMailbox / Get-ADObject in the on-premises steps
        would answer "not found" for an object that is sitting right there.

        Set-ADSessionScope (Common.ps1) sets -ViewEntireForest and never
        throws. The AD module has no session scope at all -- each call takes
        -Server, which Get-MigrationAdServer supplies.
    #>
    try {
        $r = Set-ADSessionScope
        Write-Host "Scope: $($r.Detail)" -ForegroundColor DarkGray
    } catch {
        Write-Host "WARNING: could not widen the Exchange recipient scope: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Get-MigrationAdServer {
    <#
        What to pass as -Server (AD module) or -DomainController (Exchange)
        for an object with this distinguished name. The directory is matched
        from the DN (Resolve-DistinguishedNameDirectory, longest suffix
        wins); a controller comes from Tenant.json or DCLocator
        (Resolve-DirectoryDomainController). Falls back to the directory's
        DOMAIN name, which the AD module accepts and locates a DC for; ''
        means "nothing matched -- pass nothing".
    #>
    param([Parameter(Mandatory)][string]$DistinguishedName)
    $d = $null
    try { $d = Resolve-DistinguishedNameDirectory -DistinguishedName $DistinguishedName } catch { }
    if (-not $d) { return '' }
    $dc = ''
    try { $dc = Resolve-DirectoryDomainController -DirectoryId "$($d.Id)" } catch { }
    if ("$dc".Trim()) { return "$dc".Trim() }
    return "$($d.Domain)".Trim()
}

function Get-RecipientSmtp {
    <#
        Turn whatever a Get-* cmdlet returns for a recipient reference
        (ManagedBy, GrantSendOnBehalfTo, AcceptMessagesOnlyFrom ...) into a
        primary SMTP address. On-premises those come back as canonical names
        (contoso.com/Users/Bob); in the cloud as display-ish names. Either way
        Get-Recipient -Identity resolves them, and the address is the only
        form that survives a move between directories.

        -Resolver is the Get-Recipient to use: a scriptblock so the same
        function serves an on-premises step (bare cmdlet) and a cloud step
        (Invoke-EXOCommand).
    #>
    param([object]$Refs, [Parameter(Mandatory)][scriptblock]$Resolver)
    $out = @()
    foreach ($ref in @($Refs)) {
        $id = "$ref"
        if (-not $id) { continue }
        $r = $null
        try { $r = & $Resolver $id } catch { }
        if ($r -and "$($r.PrimarySmtpAddress)") { $out += "$($r.PrimarySmtpAddress)" } else { $out += $id }
    }
    return $out
}
