#Requires -Version 5.1
<#
    Test-Toolkit.ps1
    Static self-test for a toolkit folder. Run it before every ship.

    Reasoning: ScriptNotes/Test-Toolkit.ps1.notes.md

    Author   : Glenn Armbruster
    Created  : 2026-08-03
    Revised  : 2026-08-03
#>
# TOOLKIT-HELP-BEGIN
# Checks the toolkit's own scripts for problems, without running any of them.
#
# Looks for syntax errors, wrong file encoding, scripts the console would call
# with parameters they do not accept, and a set of known traps that have caused
# real bugs here before.
#
# Point it at any toolkit folder. Run it before shipping a copy to another
# tenant.
#
# It reads scripts. It does not execute them and connects to nothing.
# TOOLKIT-HELP-END

param(
    [string]$Path,
    [switch]$Quiet,
    [ValidateSet('Error', 'Warn', 'Never')][string]$FailOn = 'Error'
)

. (Join-Path $PSScriptRoot 'Common.ps1')

if (-not $Path) { $Path = $PSScriptRoot }
$Path = (Resolve-Path -LiteralPath $Path).Path

#region What gets tested --------------------------------------------------------
function Get-TestTargetFile {
    <#
        Every .ps1 under the folder, INCLUDING subfolders.

        The -Recurse is the whole point. Without it this scanned the root only,
        so the eleven scripts under Projects\ were never checked -- six of them
        with no BOM and three carrying non-ASCII, exactly the errors this tool
        exists to catch. It reported "no errors" over files it had not opened.

        The give-away was there to be read: the Tester said 64 files while the
        version manifest covered 76. Two counts of the same toolkit that
        disagree mean one of them is wrong about what the toolkit is.

        One function, three callers. It was three copies of the same
        Get-ChildItem line, which is why the -Recurse had to be missing from
        all three at once -- and why a fix could have been applied to one and
        not the others.
    #>
    param([Parameter(Mandatory)][string]$Folder)
    return @(Get-ChildItem -LiteralPath $Folder -Filter '*.ps1' -File -Recurse -ErrorAction SilentlyContinue |
             Sort-Object FullName)
}
#endregion

#region Findings ---------------------------------------------------------------
$script:Findings = New-Object System.Collections.ArrayList

function Get-FindingDisplayName {
    <#
        Leaf name at the root; path relative to the root below it.

        This became necessary the moment the scan started recursing: a bare
        leaf name cannot distinguish a root script from one two folders down,
        and the whole value of a finding is knowing which file to open.
    #>
    param([string]$File)
    $root = "$script:TestRoot".TrimEnd('\')
    if ($root -and $File.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
        $rel = $File.Substring($root.Length).TrimStart('\')
        if ($rel) { return $rel }
    }
    return (Split-Path -Leaf $File)
}

function Add-Finding {
    param(
        [ValidateSet('ERROR', 'WARN', 'INFO')][string]$Severity,
        [string]$File, [int]$Line, [string]$Rule, [string]$Message
    )
    [void]$script:Findings.Add([PSCustomObject]@{
        Severity = $Severity
        File     = (Get-FindingDisplayName -File $File)
        Line     = $Line
        Rule     = $Rule
        Message  = $Message
        FullPath = $File
    })
}
#endregion

#region T0 -- parse ------------------------------------------------------------
function Test-Parse {
    <#
        PowerShell's own parser, without executing anything.

        This is the check that matters most and the one no amount of brace
        counting substitutes for: it is the difference between "the braces
        balance" and "this file is valid PowerShell". An unterminated block
        comment once left a script dead for days while a brace count said it
        was fine, because a stray { inside the runaway comment happened to
        offset the orphaned }.

        Note the phrasing above. This paragraph originally spelled out the
        close-comment marker literally -- and that ENDED THIS COMMENT, turning
        the rest of the sentence into code. The self-test failed with "the term
        'once' is not recognized". A block comment cannot contain its own
        terminator and there is no escape for it, so describe it in words.
    #>
    param([string]$File)
    $tokens = $null; $errors = $null
    try {
        $null = [System.Management.Automation.Language.Parser]::ParseFile($File, [ref]$tokens, [ref]$errors)
    } catch {
        Add-Finding -Severity ERROR -File $File -Line 1 -Rule 'SYN000' -Message "Parser threw: $($_.Exception.Message)"
        return $null
    }
    foreach ($e in @($errors)) {
        Add-Finding -Severity ERROR -File $File -Line $e.Extent.StartLineNumber -Rule 'SYN001' -Message $e.Message
    }
    return $tokens
}
#endregion

#region T1 -- encoding ---------------------------------------------------------
function Test-Encoding {
    param([string]$File)
    $bytes = [System.IO.File]::ReadAllBytes($File)
    $bom = @(0xEF, 0xBB, 0xBF)

    $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq $bom[0] -and $bytes[1] -eq $bom[1] -and $bytes[2] -eq $bom[2])
    if (-not $hasBom) { Add-Finding -Severity ERROR -File $File -Line 1 -Rule 'ENC001' -Message 'No UTF-8 BOM. PowerShell 5.1 falls back to the ANSI codepage, so this file is a different file on a machine with a different locale.' }
    if ($bytes.Length -ge 6 -and $hasBom -and $bytes[3] -eq $bom[0] -and $bytes[4] -eq $bom[1] -and $bytes[5] -eq $bom[2]) {
        Add-Finding -Severity ERROR -File $File -Line 1 -Rule 'ENC002' -Message 'Double BOM. Leaves a zero-width character inside line 1, which breaks #Requires.'
    }

    $lf = 0; $crlf = 0
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        if ($bytes[$i] -eq 10) { $lf++; if ($i -gt 0 -and $bytes[$i-1] -eq 13) { $crlf++ } }
    }
    if ($lf -ne $crlf) { Add-Finding -Severity ERROR -File $File -Line 1 -Rule 'ENC003' -Message "$($lf - $crlf) LF-only line ending(s). The file renders as one line in the ISE." }

    # U+FFFD means the bytes were already mangled and saved that way -- the
    # original characters are gone, not merely misread.
    $text = [System.IO.File]::ReadAllText($File)
    $n = 0
    foreach ($line in ($text -split "`n")) {
        $n++
        if ($line.IndexOf([char]0xFFFD) -ge 0) {
            Add-Finding -Severity ERROR -File $File -Line $n -Rule 'ENC004' -Message 'Replacement character (U+FFFD) -- this file has already lost characters to a bad encoding round-trip.'
        } elseif ($line -match '[^\x00-\x7E]') {
            Add-Finding -Severity WARN -File $File -Line $n -Rule 'ENC005' -Message 'Non-ASCII character. The toolkit convention is ASCII only; smart quotes and dashes from Word are the usual source.'
        }
    }
}
#endregion

#region T2 -- contract ---------------------------------------------------------
function Get-TopLevelParameterName {
    param([string]$File)
    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($File, [ref]$tokens, [ref]$errors)
    if (-not $ast -or -not $ast.ParamBlock) { return $null }   # $null = no param block at all
    return @($ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
}

function Test-Contract {
    param([string]$Folder)

    $toolsPath = Join-Path $Folder 'Config\Tools.json'
    if (-not (Test-Path -LiteralPath $toolsPath)) {
        Add-Finding -Severity WARN -File $toolsPath -Line 1 -Rule 'CFG001' -Message 'No Config\Tools.json -- contract checks skipped.'
        return
    }

    $tools = @()
    try { $tools = @((Get-Content -LiteralPath $toolsPath -Raw | ConvertFrom-Json).Tools) } catch { }
    if ($tools.Count -eq 0) {
        try { $tools = @(Get-Content -LiteralPath $toolsPath -Raw | ConvertFrom-Json) } catch {
            Add-Finding -Severity ERROR -File $toolsPath -Line 1 -Rule 'CFG002' -Message "Tools.json could not be parsed: $($_.Exception.Message)"
            return
        }
    }

    # What AdminConsole.ps1 hands every tool.
    $consoleParams = @('SmtpAddress', 'Environment', 'DomainController')

    foreach ($t in $tools) {
        $rel = "$($t.Script)"
        $full = Join-Path $Folder ($rel -replace '^\.\\', '')
        if (-not (Test-Path -LiteralPath $full)) {
            Add-Finding -Severity ERROR -File $toolsPath -Line 1 -Rule 'CFG003' -Message "'$($t.Name)' points to '$rel', which does not exist."
            continue
        }

        $declared = Get-TopLevelParameterName -File $full
        # $null means no param() block -- extras land in $args and are harmless.
        if ($null -ne $declared) {
            $unknown = @($consoleParams | Where-Object { $declared -notcontains $_ })
            $known   = @($consoleParams | Where-Object { $declared -contains $_ })
            if ($known.Count -gt 0 -and $unknown.Count -gt 0) {
                Add-Finding -Severity ERROR -File $full -Line 1 -Rule 'CFG004' -Message "Declares $($known -join ', ') but not $($unknown -join ', '). A script with a param() block REJECTS parameters it does not declare, so this fails on click unless the console filters."
            }
        }
    }
}

function Test-CommonContract {
    <#
        A name collision with Common.ps1 is only a SHADOW if the file actually
        dot-sources Common.ps1. Several tools here are deliberately standalone
        and carry their own copies because nothing else would define them --
        for those, a local definition is the only definition, and calling it a
        redefinition is false.

        The first version of this rule skipped that test and reported every
        collision. Twelve findings, of which two were real. That ratio is how a
        rule gets ignored, and an ignored rule is worse than a missing one
        because it still occupies the report.

        Worse, acting on it would have been destructive: deleting the seven
        "redefinitions" in Mailbox-Group-Inspector.ps1 -- a standalone tool --
        would have broken it outright.
    #>
    param([string]$Folder, [string[]]$Files)

    $commonPath = Join-Path $Folder 'Common.ps1'
    if (-not (Test-Path -LiteralPath $commonPath)) { return }

    $commonText = Get-Content -LiteralPath $commonPath -Raw
    $commonFns  = @([regex]::Matches($commonText, '(?m)^function\s+([A-Za-z]+-[A-Za-z0-9]+)') | ForEach-Object { $_.Groups[1].Value })

    foreach ($f in $Files) {
        if ((Split-Path -Leaf $f) -ieq 'Common.ps1') { continue }
        $text = Get-Content -LiteralPath $f -Raw

        # Dot-source, not a mention. A comment naming Common.ps1 does not load it.
        $loadsCommon = [regex]::IsMatch($text, '(?m)^\s*\.\s+.*Common\.ps1')

        $own = @([regex]::Matches($text, '(?m)^function\s+([A-Za-z]+-[A-Za-z0-9]+)') | ForEach-Object { $_.Groups[1].Value })

        foreach ($name in $own) {
            if ($commonFns -notcontains $name) { continue }

            if ($loadsCommon) {
                Add-Finding -Severity ERROR -File $f -Line 1 -Rule 'FUN001' -Message "Redefines '$name' from Common.ps1, which this file dot-sources. The local copy wins, so any fix made in Common never reaches this tool."
            } else {
                # Still worth saying once: a standalone copy is legitimate, but
                # it is a second place to maintain, and it WILL drift.
                Add-Finding -Severity INFO -File $f -Line 1 -Rule 'FUN002' -Message "Has its own '$name' and does not dot-source Common.ps1. Legitimate for a standalone tool, but this copy will drift from Common's -- check it when Common's version changes."
            }
        }
    }
}
#endregion

#region T3 -- policy -----------------------------------------------------------
function Test-Policy {
    param([string]$File)

    $raw  = Get-Content -LiteralPath $File -Raw
    $code = (Remove-PowerShellComment -Text $raw) -split "`n"

    for ($i = 0; $i -lt $code.Count; $i++) {
        $line = $code[$i]
        $n    = $i + 1

        if ($line -match '`u\{') {
            Add-Finding -Severity ERROR -File $File -Line $n -Rule 'VER001' -Message '`u{...} is PowerShell 6+. In 5.1 the literal characters end up in the string.'
        }
        if ($line -match '\?\?|\?\.' -and $line -notmatch '["'']') {
            Add-Finding -Severity WARN -File $File -Line $n -Rule 'VER002' -Message 'Possible null-coalescing / null-conditional operator -- PowerShell 7+ only.'
        }
        if ($line -match 'Start-Process' -and $line -match '-ArgumentList' -and $line -notmatch '"') {
            Add-Finding -Severity WARN -File $File -Line $n -Rule 'PRC001' -Message 'Start-Process -ArgumentList does not quote for you. A value containing a space arrives as two arguments.'
        }
        if ($line -match 'ResultSize\s*=?\s*[''"]?\d') {
            $ctx = ($code[[Math]::Max(0, $i-3)..[Math]::Min($code.Count-1, $i+3)]) -join "`n"
            if ($ctx -notmatch 'WarningAction') {
                Add-Finding -Severity WARN -File $File -Line $n -Rule 'EXO002' -Message 'A deliberately limited ResultSize makes Exchange warn "There are more results available" on every run. Add -WarningAction SilentlyContinue.'
            }
        }
    }

    # Handlers that call Exchange or AD directly with no try. A handler that
    # delegates to a function handling its own errors is fine, which is why
    # this looks for the cmdlet rather than for length.
    foreach ($m in [regex]::Matches($code -join "`n", '\.Add_\w+\(\{')) {
        $s = ($code -join "`n")
        $d = 0; $k = $m.Index + $m.Length - 1
        while ($k -lt $s.Length) {
            if ($s[$k] -eq '{') { $d++ }
            elseif ($s[$k] -eq '}') { $d--; if ($d -eq 0) { break } }
            $k++
        }
        $body = $s.Substring($m.Index + $m.Length, [Math]::Max(0, $k - $m.Index - $m.Length))
        $line = ($s.Substring(0, $m.Index) -split "`n").Count
        if ($body -notmatch '\btry\b' -and $body -match '\b(?:Get|Set|New|Remove|Add|Enable|Disable)-(?:Mailbox|Recipient|DistributionGroup|UnifiedGroup|RemoteMailbox|InboxRule|MailboxPermission|RecipientPermission|ADUser|ADGroup)\w*\b') {
            Add-Finding -Severity WARN -File $File -Line $line -Rule 'HDL001' -Message 'Click handler calls Exchange/AD directly with no try/catch. An exception here is caught by .NET, not PowerShell, and the operator sees assembly versions instead of a line number.'
        }
    }

    # Top-level property set before the control is created. Scoped to top level
    # deliberately: inside a function the statement runs at CALL time, and
    # flagging those produced six false positives and no real defects.
    $depth = 0
    $created = @{}
    $uses = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $code.Count; $i++) {
        $line = $code[$i]
        if ($depth -eq 0 -and $line -notmatch '^\s*function\s') {
            if ($line -match '^\s*(\$\w+)\s*=\s*New-Object') {
                if (-not $created.ContainsKey($Matches[1])) { $created[$Matches[1]] = $i + 1 }
            }
            if ($line -match '^\s*(\$\w+)\.(Visible|Text|Enabled|Location|Size|ForeColor|BackColor)\s*=') {
                [void]$uses.Add(@{ Var = $Matches[1]; Prop = $Matches[2]; Line = $i + 1 })
            }
        }
        $depth += ([regex]::Matches($line, '\{')).Count - ([regex]::Matches($line, '\}')).Count
    }
    foreach ($u in $uses) {
        if ($created.ContainsKey($u.Var) -and $created[$u.Var] -gt $u.Line) {
            Add-Finding -Severity ERROR -File $File -Line $u.Line -Rule 'REF001' -Message "$($u.Var).$($u.Prop) is set on line $($u.Line) but $($u.Var) is not created until line $($created[$u.Var])."
        }
    }
}
#endregion

#region Run --------------------------------------------------------------------
function Invoke-ToolkitTest {
    param([string]$Folder, [scriptblock]$Progress)

    $script:Findings.Clear()
    # Root for relative finding names. Set before anything can report.
    $script:TestRoot = $Folder
    $files = @(Get-TestTargetFile -Folder $Folder)

    $i = 0
    foreach ($f in $files) {
        $i++
        if ($Progress) { & $Progress "$($f.Name) ($i of $($files.Count))" ([int](100 * $i / [Math]::Max(1, $files.Count))) }
        Test-Encoding -File $f.FullName
        $null = Test-Parse -File $f.FullName
        try { Test-Policy -File $f.FullName } catch {
            Add-Finding -Severity WARN -File $f.FullName -Line 1 -Rule 'INT001' -Message "Policy checks could not run: $($_.Exception.Message)"
        }
    }

    if ($Progress) { & $Progress 'Contract checks...' 98 }
    Test-Contract -Folder $Folder
    Test-CommonContract -Folder $Folder -Files @($files | ForEach-Object { $_.FullName })

    return @($script:Findings.ToArray())
}

function Format-TestSummary {
    param($Findings, [string]$Folder, [int]$FileCount)

    $e = @($Findings | Where-Object { $_.Severity -eq 'ERROR' })
    $w = @($Findings | Where-Object { $_.Severity -eq 'WARN'  })

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Folder : $Folder")
    [void]$sb.AppendLine("Files  : $FileCount")
    [void]$sb.AppendLine("Result : $($e.Count) error(s), $($w.Count) warning(s)")
    [void]$sb.AppendLine('')

    if ($e.Count -eq 0 -and $w.Count -eq 0) {
        [void]$sb.AppendLine('Clean. Safe to ship.')
        return $sb.ToString()
    }

    foreach ($group in @('ERROR', 'WARN')) {
        $rows = @($Findings | Where-Object { $_.Severity -eq $group } | Sort-Object File, Line)
        if ($rows.Count -eq 0) { continue }
        [void]$sb.AppendLine("--- $group ($($rows.Count)) ---")
        foreach ($r in $rows) {
            [void]$sb.AppendLine(("  {0,-28} {1,5}  {2}  {3}" -f $r.File, $r.Line, $r.Rule, $r.Message))
        }
        [void]$sb.AppendLine('')
    }

    [void]$sb.AppendLine('Rules that need judgement rather than obedience:')
    [void]$sb.AppendLine('  ENC005  non-ASCII is sometimes deliberate (a comment quoting a Word en-dash).')
    [void]$sb.AppendLine('  HDL001  a handler delegating to a function that catches its own errors is fine.')
    [void]$sb.AppendLine('  EXO002  a probe that deliberately asks for one row is correct; silence the warning, do not raise the limit.')
    return $sb.ToString()
}
#endregion

#region Headless ---------------------------------------------------------------
if ($Quiet) {
    $files = @(Get-TestTargetFile -Folder $Path)
    $findings = Invoke-ToolkitTest -Folder $Path
    Write-Host (Format-TestSummary -Findings $findings -Folder $Path -FileCount $files.Count)

    $errCount  = @($findings | Where-Object { $_.Severity -eq 'ERROR' }).Count
    $warnCount = @($findings | Where-Object { $_.Severity -eq 'WARN'  }).Count
    $code = switch ($FailOn) {
        'Error' { if ($errCount -gt 0) { 1 } else { 0 } }
        'Warn'  { if ($errCount + $warnCount -gt 0) { 1 } else { 0 } }
        'Never' { 0 }
    }
    exit $code
}
#endregion

#region WinForms UI ------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ReportDir = Get-ReportPath

$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Toolkit Self-Test'
$form.Size            = New-Object System.Drawing.Size(1000, 720)
$form.StartPosition   = 'Manual'
# Resizable, because the findings are long lines and the useful part is often
# past the right edge. MinimumSize stops it being dragged smaller than the
# button row, which would leave controls overlapping.
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MinimumSize     = New-Object System.Drawing.Size(700, 420)
$form.MaximizeBox     = $true
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Static checks only -- nothing is executed and nothing is connected to. Safe to run any time.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(960, 18)
$lblInfo.Anchor   = 'Top,Left,Right'
$form.Controls.Add($lblInfo)

$lblPath          = New-Object System.Windows.Forms.Label
$lblPath.Text     = 'Folder:'
$lblPath.Location = New-Object System.Drawing.Point(12, 38)
$lblPath.Size     = New-Object System.Drawing.Size(50, 20)
$form.Controls.Add($lblPath)

$txtPath          = New-Object System.Windows.Forms.TextBox
$txtPath.Location = New-Object System.Drawing.Point(64, 35)
$txtPath.Size     = New-Object System.Drawing.Size(700, 22)
$txtPath.Anchor   = 'Top,Left,Right'
$txtPath.Text     = $Path
$txtPath.Font     = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtPath)

$btnBrowse          = New-Object System.Windows.Forms.Button
$btnBrowse.Text     = 'Browse...'
$btnBrowse.Location = New-Object System.Drawing.Point(772, 34)
$btnBrowse.Size     = New-Object System.Drawing.Size(90, 24)
$btnBrowse.Anchor   = 'Top,Right'
$form.Controls.Add($btnBrowse)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run Tests'
$btnRun.Location = New-Object System.Drawing.Point(870, 33)
$btnRun.Size     = New-Object System.Drawing.Size(102, 26)
$btnRun.Anchor   = 'Top,Right'
$btnRun.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnRun)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Text     = 'Idle.'
$lblStatus.Location = New-Object System.Drawing.Point(12, 66)
$lblStatus.Size     = New-Object System.Drawing.Size(960, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 86 -Width 960 -Height 12

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Location   = New-Object System.Drawing.Point(12, 148)
$txtOut.Size       = New-Object System.Drawing.Size(960, 486)
$txtOut.Anchor     = 'Top,Left,Right,Bottom'
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
# WordWrap ON. It was off with a horizontal scrollbar, and a finding whose
# explanation sits past the right edge is a finding nobody reads -- the same
# mistake that hid the write-scope message for three rounds on 2026-08-06.
$txtOut.WordWrap   = $true
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

# ReadOnly does not stop selection, but a button beats explaining that.
$btnCopy          = New-Object System.Windows.Forms.Button
$btnCopy.Text     = 'Copy Results'
$btnCopy.Location = New-Object System.Drawing.Point(12, 642)
$btnCopy.Size     = New-Object System.Drawing.Size(120, 28)
$btnCopy.Anchor   = 'Bottom,Left'
$form.Controls.Add($btnCopy)

$btnOpenCsv          = New-Object System.Windows.Forms.Button
$btnOpenCsv.Text     = 'Open CSV'
$btnOpenCsv.Location = New-Object System.Drawing.Point(140, 642)
$btnOpenCsv.Size     = New-Object System.Drawing.Size(110, 28)
$btnOpenCsv.Anchor   = 'Bottom,Left'
$form.Controls.Add($btnOpenCsv)

$lblCopied           = New-Object System.Windows.Forms.Label
$lblCopied.Tag       = 'NoTheme'
$lblCopied.Location  = New-Object System.Drawing.Point(260, 648)
$lblCopied.Size      = New-Object System.Drawing.Size(500, 18)
$lblCopied.Anchor    = 'Bottom,Left,Right'
$lblCopied.ForeColor = [System.Drawing.Color]::FromArgb(80, 220, 130)
$form.Controls.Add($lblCopied)

$btnCopy.Add_Click({
    if (-not "$($txtOut.Text)".Trim()) { $lblCopied.Text = 'Nothing to copy -- run the tests first.'; return }
    try {
        [System.Windows.Forms.Clipboard]::SetText($txtOut.Text)
        $lblCopied.Text = "Copied $(($txtOut.Text -split "`n").Count) line(s) to the clipboard."
    } catch {
        $lblCopied.Text = "Could not copy: $($_.Exception.Message)"
    }
})

$btnOpenCsv.Add_Click({
    # The CSV is written on every run; this opens the newest without going
    # through the Previous runs dropdown.
    try {
        $newest = Get-ChildItem -LiteralPath $ReportDir -Filter 'ToolkitTest_*.csv' -ErrorAction Stop |
                  Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $newest) { $lblCopied.Text = 'No CSV yet -- run the tests first.'; return }
        Invoke-Item -LiteralPath $newest.FullName
        $lblCopied.Text = "Opened $($newest.Name)"
    } catch {
        $lblCopied.Text = "Could not open the CSV: $($_.Exception.Message)"
    }
})

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'ToolkitTest' `
            -X 12 -Y 106 -Width 960 -StatusLabel $lblStatus

$btnBrowse.Add_Click({
    try {
        $fb = New-Object System.Windows.Forms.FolderBrowserDialog
        $fb.Description = 'Toolkit folder to test'
        if (Test-Path -LiteralPath $txtPath.Text) { $fb.SelectedPath = $txtPath.Text }
        if ($fb.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtPath.Text = $fb.SelectedPath }
    } catch { Show-HandlerError -ErrorRecord $_ -Context 'Browse for toolkit folder' }
})

$btnRun.Add_Click({
    $btnRun.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $folder = "$($txtPath.Text)".Trim()
        if (-not (Test-Path -LiteralPath $folder)) {
            $txtOut.Text = "Folder not found: $folder"
            return
        }

        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar -Activity 'Toolkit self-test'
        $fileList = @(Get-TestTargetFile -Folder $folder)
        $findings = Invoke-ToolkitTest -Folder $folder -Progress $progress
        Complete-ProgressReporter -Activity 'Toolkit self-test' -Bar $pbar

        $summary = Format-TestSummary -Findings $findings -Folder $folder -FileCount $fileList.Count

        # Integrity is reported first because it answers a different question
        # from the rules: not "is this code sound" but "is this the code I
        # shipped". A drifted file can be perfectly valid PowerShell.
        $intro = New-Object System.Text.StringBuilder
        try {
            $vv = Test-ToolkitIntegrity
            [void]$intro.AppendLine('=== VERSION AND INTEGRITY ===')
            if (-not $vv.HasManifest) {
                [void]$intro.AppendLine('No Version.json -- this copy has never been stamped.')
                [void]$intro.AppendLine('Run Publish-Toolkit.ps1 on the MASTER copy to create one.')
            } else {
                [void]$intro.AppendLine("Version  : $($vv.Version)   released $($vv.Released)")
                [void]$intro.AppendLine("Verified : $($vv.Matched) file(s) match the manifest")
                # Only when something already differs. A newer file that hashes
                # the same is identical, and saying so would be a warning about
                # nothing -- see the Stale note in Common.ps1.
                if ($vv.Stale -and ($vv.Modified.Count -or $vv.Missing.Count -or $vv.Extra.Count)) {
                    [void]$intro.AppendLine('NOTE     : the manifest is older than the code, which likely explains the below.')
                    [void]$intro.AppendLine('           Re-run Publish-Toolkit.ps1 on the master.')
                }
                foreach ($set in @(
                    @{ Name = 'MODIFIED (differ from the master)'; Items = $vv.Modified },
                    @{ Name = 'MISSING (in the manifest, not on disk)'; Items = $vv.Missing },
                    @{ Name = 'EXTRA (on disk, not in the manifest)'; Items = $vv.Extra })) {
                    if ($set.Items.Count -gt 0) {
                        [void]$intro.AppendLine('')
                        [void]$intro.AppendLine("$($set.Name): $($set.Items.Count)")
                        foreach ($i in $set.Items) { [void]$intro.AppendLine("    $i") }
                    }
                }
            }
            [void]$intro.AppendLine('')
        } catch {
            [void]$intro.AppendLine('=== VERSION AND INTEGRITY ===')
            [void]$intro.AppendLine("Could not check: $($_.Exception.Message)")
            [void]$intro.AppendLine('')
        }

        $txtOut.Text = $intro.ToString() + $summary

        $errCount = @($findings | Where-Object { $_.Severity -eq 'ERROR' }).Count
        $lblStatus.Text = if ($errCount -gt 0) { "$errCount error(s) -- do not ship." } else { 'No errors.' }

        if ($findings.Count -gt 0) {
            if (-not (Test-Path -LiteralPath $ReportDir)) {
                New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
            }
            $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
            $csv = Join-Path $ReportDir "ToolkitTest_$stamp.csv"
            $findings | Select-Object Severity, File, Line, Rule, Message |
                Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
            $files.SelectPath($csv)
        }

        $result = if ($errCount -gt 0) { 'Failed' } else { 'Success' }
        Write-CompletedLog -Action 'Run' -Target (Split-Path -Leaf $folder) -Result $result `
            -Detail "Toolkit self-test: $($fileList.Count) file(s), $errCount error(s), $(@($findings | Where-Object { $_.Severity -eq 'WARN' }).Count) warning(s)"
    } catch {
        Show-HandlerError -ErrorRecord $_ -Context 'Toolkit self-test'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Toolkit self-test'
    } finally {
        $form.Cursor = $prev
        Complete-ProgressReporter -Activity 'Toolkit self-test'
        $btnRun.Enabled = $true
    }
})

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled, and BringToFront'd inside the helper because WinForms
# draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Tester' | Out-Null

[void]$form.ShowDialog()
#endregion
