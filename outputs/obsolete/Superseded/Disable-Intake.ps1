#Requires -Version 5.1
<#
    Disable-Intake.ps1
    Intake form for termination / disable requests.

    Author   : Glenn Armbruster
    Revised  : 2026-07-28
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Parsing ------------------------------------------------------------

function Normalize-Value {
    param([string]$value)
    if (-not $value) { return "" }
    $trimmed = $value.Trim()
    $emptyTokens = @("NA","N/A","NONE","NULL","-","NO")
    if ($emptyTokens -contains $trimmed.ToUpper()) { return "" }
    if ($trimmed -match "@") { return $trimmed.Split("@")[0] }
    return $trimmed
}

function Normalize-Line {
    param([string]$text)
    if (-not $text) { return "" }
    $clean = $text
    $clean = $clean.Replace([char]0xA0, ' ')
    $clean = $clean.Replace([char]0x2007, ' ')
    $clean = $clean.Replace([char]0x202F, ' ')
    $clean = $clean.Replace([string][char]0x00AD, "")
    return $clean.Trim()
}

function Normalize-Manager {
    param([string]$value)
    if ([string]::IsNullOrWhiteSpace($value)) { return "" }
    $v = $value.Trim()
    if ($v -match '@') { $v = $v.Split('@')[0] }
    $v = $v -replace '\s+', '.'
    if ($v -notmatch '^[A-Za-z0-9._\-]+$') { return $null }
    return $v
}

function Validate-LogonName {
    param([string]$value)
    if ($value -match '^[A-Za-z0-9]+$') { return $true }
    return $false
}

function Test-Identity {
    param([string]$value)
    if ([string]::IsNullOrWhiteSpace($value)) { return $false }
    $v = $value.Trim()
    if ($v -match '^[A-Za-z0-9]+$') { return $true }
    if ($v -match '^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$') { return $true }
    return $false
}

function Show-ParseError {
    param([string]$Message)
    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        "Input Error",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Get-NormalizedLines {
    param([string]$Text)
    $raw = $Text -split "\r?\n"
    return ($raw | ForEach-Object { Normalize-Line $_ } | Where-Object { $_ -ne "" })
}

function New-Record {
    param([string]$Type)
    return [pscustomobject]@{
        Type     = $Type
        MB       = ""
        Delegate = ""
        OD       = ""
        Employee = ""
        LWD      = ""
    }
}

function Parse-HQ {
    param([string]$Text)
    $lines = Get-NormalizedLines -Text $Text
    if ($lines.Count -eq 0) {
        return [pscustomobject]@{ Ok = $true; Records = @() }
    }
    $records = @()
    $cur     = $null
    foreach ($line in $lines) {
        if ($line -match '^\s*([^:]+?)\s*:\s*(.*)$') {
            $label = $Matches[1].Trim().ToLower()
            $val   = $Matches[2].Trim()
            switch ($label) {
                'employee name' {
                    if ($cur) { $records += $cur }
                    $cur = New-Record 'HQ'
                    $cur.Employee = $val
                }
                'upn' {
                    if (-not $cur -or $cur.MB -ne '') {
                        if ($cur) { $records += $cur }
                        $cur = New-Record 'HQ'
                    }
                    $cur.MB = $val
                }
                'last working day' {
                    if ($cur) { $cur.LWD = $val }
                }
                default { }
            }
        }
    }
    if ($cur) { $records += $cur }
    if ($records.Count -eq 0) {
        Show-ParseError "The HQ box has text but no 'Employee Name:' or 'UPN:' entries were found."
        return [pscustomobject]@{ Ok = $false; Records = @() }
    }
    foreach ($r in $records) {
        if ([string]::IsNullOrWhiteSpace($r.MB)) {
            $who = if ($r.Employee) { $r.Employee } else { "(unnamed HQ block)" }
            Show-ParseError "An HQ block is missing its UPN:`n`n$who"
            return [pscustomobject]@{ Ok = $false; Records = @() }
        }
        if (-not (Test-Identity $r.MB)) {
            Show-ParseError "Invalid HQ UPN / identity:`n`n$($r.MB)"
            return [pscustomobject]@{ Ok = $false; Records = @() }
        }
    }
    return [pscustomobject]@{ Ok = $true; Records = $records }
}

function Parse-Franchise {
    param([string]$Text)
    $lines = Get-NormalizedLines -Text $Text
    if ($lines.Count -eq 0) {
        return [pscustomobject]@{ Ok = $true; Records = @() }
    }
    $records = @()
    $cur     = $null
    foreach ($line in $lines) {
        if ($line -match '^\s*([^:]+?)\s*:\s*(.*)$') {
            $label = $Matches[1].Trim().ToLower()
            $val   = $Matches[2].Trim()
            switch ($label) {
                'logon name' {
                    if ($cur) { $records += $cur }
                    $cur = New-Record 'Franchise'
                    $cur.MB = $val
                }
                'delegate' {
                    if ($cur) { $cur.Delegate = Normalize-Value $val }
                }
                { $_ -in @('onedrive','one drive') } {
                    if ($cur) { $cur.OD = Normalize-Value $val }
                }
                default { }
            }
        }
    }
    if ($cur) { $records += $cur }
    if ($records.Count -eq 0) {
        Show-ParseError "The Franchise box has text but no 'Logon Name:' entries were found."
        return [pscustomobject]@{ Ok = $false; Records = @() }
    }
    foreach ($r in $records) {
        if (-not (Validate-LogonName $r.MB)) {
            Show-ParseError "Franchise Logon Name must be alphanumeric (a sAMAccountName, not an email):`n`n$($r.MB)"
            return [pscustomobject]@{ Ok = $false; Records = @() }
        }
        if ($r.Delegate -ne '' -and $r.Delegate -notmatch '^[A-Za-z0-9._\-]+$') {
            Show-ParseError "A delegate value does not look valid:`n`n$($r.Delegate)"
            return [pscustomobject]@{ Ok = $false; Records = @() }
        }
    }
    return [pscustomobject]@{ Ok = $true; Records = $records }
}

#endregion

#region Actions ------------------------------------------------------------

<#
    The disable-request folder, resolved rather than hardcoded.

    This was 'E:\scripts\m365\disabledrequests\', which assumes a drive letter
    the machine may not have. Set the real location in ToolManager
    ("Settings..."), which writes Settings.json.

    Get-DisableRequestPath comes from Common.ps1. This script does NOT
    dot-source it -- it inherits the function because AdminConsole launches
    tools with '&' in the same process. That works from the console and not
    from a bare PowerShell window, so the fallback below covers a standalone
    run rather than failing with a confusing "command not found".
#>
$DisableRequestFolder = if (Get-Command -Name 'Get-DisableRequestPath' -ErrorAction SilentlyContinue) {
    Get-DisableRequestPath
} else {
    Join-Path $PSScriptRoot 'DisableRequests'
}

$HqScriptPath        = Join-Path $DisableRequestFolder 'disable-new-hq.ps1'
$FranchiseScriptPath = Join-Path $DisableRequestFolder 'disable-new.ps1'
$CsvPath             = Join-Path $DisableRequestFolder 'disablelist.csv'

# A wrong folder must be obvious. Without this the queue simply reads as empty
# and the two action buttons fail one at a time, which looks like a permissions
# problem rather than a path problem.
if (-not (Test-Path -LiteralPath $DisableRequestFolder)) {
    Write-Warning "Disable-request folder not found: '$DisableRequestFolder'. Set it in ToolManager -> Settings..."
}

function Format-Plan {
    param($Fran, $HQ)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("This intake will process:")
    [void]$sb.AppendLine("")
    if ($Fran.Count -gt 0) {
        [void]$sb.AppendLine("FRANCHISE  ($($Fran.Count))  ->  disable-new.ps1")
        foreach ($r in $Fran) {
            $del = if ($r.Delegate) { "$($r.Delegate)@expresspros.com" } else { "(none)" }
            $od  = if ($r.OD)       { "$($r.OD)@expresspros.com" }       else { "(none)" }
            [void]$sb.AppendLine("   $($r.MB)")
            [void]$sb.AppendLine("      delegate=$del   onedrive=$od")
        }
        [void]$sb.AppendLine("")
    }
    if ($HQ.Count -gt 0) {
        [void]$sb.AppendLine("HQ  ($($HQ.Count))  ->  disable-new-hq.ps1")
        foreach ($r in $HQ) {
            $name = if ($r.Employee) { "   ($($r.Employee))" } else { "" }
            $del  = if ($r.Delegate) { "$($r.Delegate)@expresspros.com" } else { "(none)" }
            $od   = if ($r.OD)       { "$($r.OD)@expresspros.com" }       else { "(none)" }
            [void]$sb.AppendLine("   $($r.MB)$name")
            [void]$sb.AppendLine("      delegate=$del   onedrive=$od")
        }
        [void]$sb.AppendLine("")
    }
    if ($Fran.Count -gt 0 -and $HQ.Count -gt 0) {
        [void]$sb.AppendLine("Run order: Franchise first, then HQ (separate runs).")
        [void]$sb.AppendLine("")
    }
    [void]$sb.AppendLine("Proceed?")
    return $sb.ToString()
}

function Write-DisableCsv {
    param($Records)
    $out = @("MB,Delegate,OD")
    foreach ($r in $Records) {
        $out += "$($r.MB),$($r.Delegate),$($r.OD)"
    }
    $out | Set-Content -Path $CsvPath -Encoding UTF8
    return $out
}

#endregion

#region WinForms UI --------------------------------------------------------

$script:FranRecords = @()
$script:HqRecords   = @()

$form = New-Object System.Windows.Forms.Form
$form.Text          = "Disabled User Intake"
$form.Size          = New-Object System.Drawing.Size(525, 475)
$form.StartPosition = "CenterScreen"

$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Text        = "Make sure you are connected to on-prem Exchange, Exchange Online, and SharePoint Online before running."
$lblWarn.MaximumSize = New-Object System.Drawing.Size(495, 0)
$lblWarn.AutoSize    = $true
$lblWarn.Location    = New-Object System.Drawing.Point(12, 10)
$lblWarn.Font        = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$lblWarn.ForeColor   = [System.Drawing.Color]::Red
$form.Controls.Add($lblWarn)

$lblHQ          = New-Object System.Windows.Forms.Label
$lblHQ.Text     = "HQ Request (paste SNOW Description):"
$lblHQ.AutoSize = $true
$lblHQ.Location = New-Object System.Drawing.Point(12, 58)
$form.Controls.Add($lblHQ)

$txtHQ             = New-Object System.Windows.Forms.TextBox
$txtHQ.Multiline   = $true
$txtHQ.ScrollBars  = "Vertical"
$txtHQ.Font        = New-Object System.Drawing.Font("Consolas", 9)
$txtHQ.Size        = New-Object System.Drawing.Size(486, 88)
$txtHQ.Location    = New-Object System.Drawing.Point(12, 76)
$form.Controls.Add($txtHQ)

$lblHQMgr          = New-Object System.Windows.Forms.Label
$lblHQMgr.Text     = "HQ Manager:"
$lblHQMgr.AutoSize = $true
$lblHQMgr.Location = New-Object System.Drawing.Point(12, 176)
$form.Controls.Add($lblHQMgr)

$txtHQMgr          = New-Object System.Windows.Forms.TextBox
$txtHQMgr.Size     = New-Object System.Drawing.Size(200, 22)
$txtHQMgr.Location = New-Object System.Drawing.Point(100, 173)
$form.Controls.Add($txtHQMgr)

$lblFran          = New-Object System.Windows.Forms.Label
$lblFran.Text     = "Franchise requests (Logon Name / Delegate / OneDrive) - paste one or more:"
$lblFran.AutoSize = $true
$lblFran.Location = New-Object System.Drawing.Point(12, 206)
$form.Controls.Add($lblFran)

$txtFran           = New-Object System.Windows.Forms.TextBox
$txtFran.Multiline = $true
$txtFran.ScrollBars = "Vertical"
$txtFran.Font      = New-Object System.Drawing.Font("Consolas", 9)
$txtFran.Size      = New-Object System.Drawing.Size(486, 140)
$txtFran.Location  = New-Object System.Drawing.Point(12, 224)
$form.Controls.Add($txtFran)

$lblNote           = New-Object System.Windows.Forms.Label
$lblNote.Text      = "Either section may be left empty - the other is still processed."
$lblNote.AutoSize  = $true
$lblNote.Location  = New-Object System.Drawing.Point(12, 372)
$lblNote.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblNote)

$okButton          = New-Object System.Windows.Forms.Button
$okButton.Text     = "OK"
$okButton.Location = New-Object System.Drawing.Point(333, 398)
$okButton.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtHQ.Text) -and [string]::IsNullOrWhiteSpace($txtFran.Text)) {
        Show-ParseError "Both sections are empty - nothing to process."
        return
    }

    $rHQ = Parse-HQ -Text $txtHQ.Text
    if (-not $rHQ.Ok) { return }

    $mgrLocal = Normalize-Manager -value $txtHQMgr.Text
    if ($null -eq $mgrLocal) {
        Show-ParseError "HQ Manager value is not valid: '$($txtHQMgr.Text.Trim())'`nEnter as 'bob smith' or 'bob.smith'."
        return
    }
    foreach ($r in $rHQ.Records) {
        $r.Delegate = $mgrLocal
        $r.OD       = $mgrLocal
    }

    $rFran = Parse-Franchise -Text $txtFran.Text
    if (-not $rFran.Ok) { return }

    $hq   = @($rHQ.Records)
    $fran = @($rFran.Records)

    if ($hq.Count -eq 0 -and $fran.Count -eq 0) {
        Show-ParseError "Nothing recognizable was found in either section."
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        (Format-Plan -Fran $fran -HQ $hq),
        "Confirm Disable Intake",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $script:HqRecords   = $hq
    $script:FranRecords = $fran
    $form.Tag = "OK"
    $form.Close()
})
$form.Controls.Add($okButton)

$cancelButton          = New-Object System.Windows.Forms.Button
$cancelButton.Text     = "Cancel"
$cancelButton.Location = New-Object System.Drawing.Point(418, 398)
$cancelButton.Add_Click({ $form.Tag = "Cancel"; $form.Close() })
$form.Controls.Add($cancelButton)

$form.StartPosition = 'Manual'
$wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$y  = [Math]::Max($wa.Top, $wa.Bottom - $form.Height)
$form.Location = New-Object System.Drawing.Point($wa.Left, $y)
if (Get-Command Set-ConsoleTheme -ErrorAction SilentlyContinue) {
    Set-ConsoleTheme -Form $form
}
$form.ShowDialog() | Out-Null

#endregion

if ($form.Tag -ne "OK") { return }

if ($script:FranRecords.Count -gt 0) {
    $csv = Write-DisableCsv -Records $script:FranRecords
    Write-Host "===================================="
    Write-Host " disablelist.csv generated (Franchise)"
    Write-Host "===================================="
    Write-Host "Path: $CsvPath"
    Write-Host ""
    $csv | ForEach-Object { Write-Host $_ }
    Write-Host "`nRunning Franchise disable..."
& $FranchiseScriptPath
}

if ($script:HqRecords.Count -gt 0) {
    $csv = Write-DisableCsv -Records $script:HqRecords
    Write-Host "===================================="
    Write-Host " disablelist.csv generated (HQ)"
    Write-Host "===================================="
    Write-Host "Path: $CsvPath"
    Write-Host ""
    $csv | ForEach-Object { Write-Host $_ }
    Write-Host "`nRunning HQ disable..."
& $HqScriptPath
}

Write-Host "`nIntake complete."
exit
