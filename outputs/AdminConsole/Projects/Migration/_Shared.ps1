#Requires -Version 5.1
<#
    _Shared.ps1  --  helpers common to every step of the Migration project.

    NOT a tool and NOT a fork of Common.ps1. Every step dot-sources the one
    Common.ps1 at the toolkit root first, then this file for the handful of
    things only this project needs: the instruction popup, the newest-export
    picker, and the list join/split used to round-trip multi-valued
    attributes through CSV.

    The underscore keeps it at the top of the folder and out of the numbered
    sequence. It has no TOOLKIT-HELP block on purpose: it is never launched.

    Author   : Glenn Armbruster
    Revised  : 2026-09-08
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

function Read-InputList {
    <#
        Opens the input CSV for review, waits, then returns its Name column
        trimmed and de-blanked. Same shape as DLtoSharedMbx step 1: the review
        pause is the only gate between "I edited the wrong file" and a run.
    #>
    param([Parameter(Mandatory)][string]$Path, [string]$What = 'objects')
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Input file not found: $Path  --  it ships with the project as a one-column template (Name)."
    }
    Write-Host "Input file: $Path"
    Invoke-Item -LiteralPath $Path
    Write-Host ''
    Read-Host "Review $([System.IO.Path]::GetFileName($Path)) -- one $What per row under 'Name' -- save, then press ENTER" | Out-Null
    $rows = @(Import-Csv -LiteralPath $Path)
    $names = @($rows | ForEach-Object { "$($_.Name)".Trim() } | Where-Object { $_ })
    if ($names.Count -eq 0) { throw "No rows in $([System.IO.Path]::GetFileName($Path)). Nothing to do." }
    Write-Host "$($names.Count) $what listed." -ForegroundColor Cyan
    return $names
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
        this every Get-RemoteMailbox / Get-ADObject in Steps 1, 3, 3R and 6
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

# THE DISPLAY-NAME TAG (2026-10-05). The admin center's Deleted users page
# searches on display name, so a batch tagged 'SMB-10-10-2026-' is one
# search, one select-all, one Restore -- instead of 456 of them. Step 1
# writes it on-premises AFTER exporting the clean name; Step 2 refuses to
# call the batch ready until the tag has synced; Step 4 (restamp) puts the
# clean name back from the exports. Date is the day Step 1 ran.
$script:TagPattern = '^SMB-\d{1,2}-\d{1,2}-\d{4}-'
function Get-MigrationTag { return ('SMB-' + (Get-Date -Format 'M-d-yyyy') + '-') }
function Test-MigrationTagged { param([string]$DisplayName) return ("$DisplayName" -match $script:TagPattern) }
function Remove-MigrationTag { param([string]$DisplayName) return ("$DisplayName" -replace $script:TagPattern, '') }

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
