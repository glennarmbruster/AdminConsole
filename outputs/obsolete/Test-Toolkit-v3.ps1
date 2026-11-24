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

#region T6 -- tenant leakage ----------------------------------------------------
function Test-TenantLeakage {
    <#
        A tenant name, domain or server in a script that ships to every tenant.

        This is the check the whole refactor depends on. One toolkit serving six
        tenants only works while the code contains no tenant, and "we were
        careful" is not a mechanism -- the toolkit had 241 such values and
        nobody had noticed, including in a suite explicitly named agnostic.

        EX- FILES ARE EXEMPT AND THAT IS THE POINT. Some scripts genuinely
        cannot be generalised: a 3,000-line standalone tool that keeps private
        constants precisely so it runs with no toolkit, or a link to one
        company's ticket system. The prefix is a declaration that the file
        ships to one tenant, and this rule is what makes the declaration mean
        something -- an unprefixed file may not contain tenant data, and a
        prefixed one may.

        Config\Tenant.json is exempt for the same reason: it is the one place
        tenant data is supposed to live.
    #>
    param([string]$Folder, [string[]]$Files)

    # Names are read from Tenant.json, not hardcoded here -- a checker with a
    # baked-in list of tenants would be the very thing it is checking for.
    $names = New-Object System.Collections.Generic.List[string]
    $tj = Join-Path $Folder 'Config\Tenant.json'
    try {
        if (-not (Test-Path -LiteralPath $tj)) { throw "Config\Tenant.json not found." }
        $cfg = Get-Content -LiteralPath $tj -Raw | ConvertFrom-Json
        foreach ($v in @($cfg.TenantName, $cfg.PrimaryDomain, $cfg.OnMicrosoftDomain)) {
            if ($v) { [void]$names.Add("$v") }
        }
        foreach ($d in @($cfg.Directories)) {
            if ($d.Domain) { [void]$names.Add("$($d.Domain)") }
            foreach ($dc in @($d.DomainControllers)) { if ($dc) { [void]$names.Add("$dc") } }
        }

        # Group names count. This list originally covered domains and servers
        # only, so 'E5 License Fran - Full' sat hardcoded in two scripts that
        # ship to every tenant and the rule was structurally unable to see it --
        # it is in Tenant.json under LicenseGroups, which nothing read.
        #
        # Anything NAMED in Tenant.json is tenant data by definition. If a value
        # is worth configuring per tenant, it is worth catching in code.
        foreach ($g in @($cfg.LicenseGroups)) {
            if ($g.Name) { [void]$names.Add("$($g.Name)") }
        }

        # THE UNDERSCORE FORM. OneDrive personal site URLs write a domain with
        # dots replaced by underscores -- 'contoso.com' becomes 'contoso_com'.
        #
        # That is the same tenant data in a disguise no domain rule can see:
        # the configured string does not appear, and the word-boundary check
        # cannot help because underscores make it one long token. A hardcoded
        # tenant domain sat in the URL-builder of MailboxOneDriveAccess.ps1,
        # in shipping code, and passed every check here.
        #
        # Any transformation of a tenant value is still that tenant value.
        $domainish = @($cfg.PrimaryDomain, $cfg.OnMicrosoftDomain)
        foreach ($d in @($cfg.Directories)) { $domainish += $d.Domain }
        foreach ($v in $domainish) {
            if ($v -and "$v".Contains('.')) { [void]$names.Add(("$v" -replace '\.', '_')) }
        }
        if ($cfg.SharePoint) {
            foreach ($p in @('MySiteHost', 'AdminUrl', 'RootUrl')) {
                if ($cfg.SharePoint.PSObject.Properties[$p] -and $cfg.SharePoint.$p) {
                    [void]$names.Add("$($cfg.SharePoint.$p)")
                }
            }
        }
    } catch {
        Add-Finding -Severity ERROR -File $tj -Line 1 -Rule 'TEN003' `
            -Message "Tenant.json could not be read, so TEN002 checked NOTHING this run: $($_.Exception.Message)"
        return
    }

    # A rule that can quietly check nothing is worse than no rule, because the
    # clean run is read as proof. This was an empty catch{} around the block
    # above: any failure left $names empty, TEN002 could never fire, and the
    # output looked like a pass. It now reports the list it is matching on, so
    # "no TEN002 findings" can be told apart from "TEN002 never ran".
    if ($names.Count -eq 0) {
        Add-Finding -Severity ERROR -File $tj -Line 1 -Rule 'TEN003' `
            -Message "Tenant.json parsed but named nothing, so TEN002 checked NOTHING this run."
        return
    }
    Add-Finding -Severity INFO -File $tj -Line 1 -Rule 'TEN000' `
        -Message "TEN002 is matching $($names.Count) name(s) from Tenant.json: $($names -join ', ')"

    # Anything shaped like an internal FQDN or a vanity SharePoint host, plus
    # whatever this deployment's own Tenant.json names.
    $shapes = @(
        '\b[a-z0-9-]+\.(?:franchises|corp|internal|ad|local)\.[a-z.]+\b'
        '\b[a-z0-9-]+-(?:my|admin)\.sharepoint\.com\b'
        '\b[a-z0-9-]+\.mail\.onmicrosoft\.com\b'
        '\b[a-z0-9-]+\.service-now\.com\b'
    )

    # Projects\<Tenant>\ is exempt for the same reason EX- is: the path is a
    # declaration that the file belongs to one tenant and is pulled before the
    # toolkit ships to another. The difference is that EX- was a decision and
    # this was not -- these files simply never produced a finding, and an
    # exemption nobody chose is indistinguishable from a rule that is broken.
    # Counting them makes the exemption something you can see and argue with.
    $projectSkips = 0

    foreach ($f in $Files) {
        $full = if ($f -is [System.IO.FileInfo]) { $f.FullName } else { "$f" }
        $leaf = Split-Path -Leaf $full
        if ($leaf -like 'EX-*') { continue }
        if ($leaf -ieq 'Test-Toolkit.ps1') { continue }

        $rel = $full
        if ($Folder -and $full.StartsWith($Folder, [StringComparison]::OrdinalIgnoreCase)) {
            $rel = $full.Substring($Folder.Length).TrimStart('\', '/')
        }
        if ($rel -match '^Projects[\\/]') { $projectSkips++; continue }

        # Comment-stripped, like Test-Policy. Skipping only lines that START
        # with '#' left block comments in scope, so widening the name list
        # immediately flagged a docstring in Common.ps1 that QUOTES a licence
        # group name while explaining an encoding trap. Documenting a value is
        # not shipping it, and Remove-PowerShellComment preserves line numbers
        # so findings still point at the right line.
        $lines = @()
        try {
            $raw = Get-Content -LiteralPath $full -Raw -ErrorAction Stop
            $lines = @((Remove-PowerShellComment -Text $raw) -split "`n")
        } catch { continue }

        for ($i = 0; $i -lt $lines.Count; $i++) {
            $ln = $lines[$i]
            if ($ln.TrimStart().StartsWith('#')) { continue }

            foreach ($s in $shapes) {
                $m = [regex]::Match($ln, $s, 'IgnoreCase')

                # Skip matches that are really a variable being interpolated.
                # "https://$label-admin.sharepoint.com" BUILDS the host from a
                # configured value -- the template is the correct answer, not a
                # leak. Without this the rule fires on the very function that
                # exists to keep the URL out of the code, and a rule that
                # flags the fix is a rule people switch off.
                if ($m.Success -and $m.Index -gt 0 -and
                    ('$', ')', '}') -contains $ln[$m.Index - 1]) { continue }

                if ($m.Success) {
                    Add-Finding -Severity ERROR -File $full -Line ($i + 1) -Rule 'TEN001' `
                        -Message "Tenant-specific value '$($m.Value)' in a script that ships to every tenant. Move it to Tenant.json, or rename the file EX- if it genuinely cannot be generalised."
                    break
                }
            }
            foreach ($n in $names) {
                if ($n.Length -lt 4) { continue }

                # WORD BOUNDARY, not substring. A TenantName of 'Express'
                # matched 'Expression' -- and @{ Expression = ... } appears in
                # every report that sorts, so the rule produced dozens of
                # findings on correct code. A check that cries wolf is a check
                # that gets switched off, and this is the second time today the
                # same substring trap has been walked into.
                if ($ln -match ('(?<![A-Za-z0-9])' + [regex]::Escape($n) + '(?![A-Za-z0-9])')) {
                    Add-Finding -Severity ERROR -File $full -Line ($i + 1) -Rule 'TEN002' `
                        -Message "'$n' comes from this deployment's Tenant.json and must not appear in code. Read it through Get-TenantConfig."
                    break
                }
            }
        }
    }

    if ($projectSkips -gt 0) {
        Add-Finding -Severity INFO -File (Join-Path $Folder 'Projects') -Line 1 -Rule 'TEN000' `
            -Message "$projectSkips file(s) under Projects\ were not checked for tenant leakage -- that folder is tenant-scoped by design and is pulled before the toolkit ships elsewhere. Confirm it IS pulled."
    }
}
#endregion

#region T7 -- command resolution ------------------------------------------------
function Test-CommandResolution {
    <#
        A call to a function that does not exist.

        Get-MailDomain was called in 23 places across 12 files and defined in
        none of them: the tenant refactor replaced hardcoded domains with calls
        to a helper nobody wrote. It parsed, it hashed, it passed every check
        here, and it threw the moment the console opened.

        It hid because every call site is inside a string -- "...@$(Get-MailDomain)".
        A subexpression is not evaluated until the string is built, so there is
        nothing for the parser to object to. It hid a second time inside a try,
        where Get-SpoAdminUrl used it as a fallback and reported "nothing could
        be derived", which reads as an unset option rather than a missing
        function.

        TWO PASSES, because a string is both hiding place and false friend:

          A. the inside of every $( ), which is code even in a string. This is
             the one that matters.
          B. the line with string literals REMOVED, which is ordinary code.

        Doing only B would miss the bug. Doing neither -- scanning raw lines --
        flags 'Send-As' and 'On-Prem' in message text, and a rule that fires on
        prose is a rule that gets switched off.

        Resolution order is defined-here, then allowlist, then Get-Command.
        Get-Command is what makes this affordable: it knows New-Object and
        Where-Object without a list. The allowlist covers server cmdlets that
        are absent until a session connects -- flagging Get-Mailbox because
        Exchange is not loaded would make the tester useless offline.
    #>
    param([string]$Folder, $Files)

    $defined = @{}
    $paths = @()
    foreach ($f in $Files) {
        $full = if ($f -is [System.IO.FileInfo]) { $f.FullName } else { "$f" }
        $paths += $full
        try { $raw = Get-Content -LiteralPath $full -Raw -ErrorAction Stop } catch { continue }
        foreach ($m in [regex]::Matches($raw, '(?mi)^\s*function\s+([A-Za-z0-9_\-]+)')) {
            $defined[$m.Groups[1].Value.ToLower()] = $true
        }
    }

    # Cmdlets that only exist once a session is connected. Matched on the NOUN
    # so a new verb on a known noun is covered without editing this.
    $allow = '(?i)-(AD[A-Za-z0-9]*|SPO[A-Za-z0-9]*|Mg[A-Za-z0-9]*|EXO[A-Za-z0-9]*|' +
             'Mailbox[A-Za-z0-9]*|DistributionGroup[A-Za-z0-9]*|UnifiedGroup[A-Za-z0-9]*|' +
             'Recipient[A-Za-z0-9]*|InboxRule|CASMailbox|MessageTrace[A-Za-z0-9]*|' +
             'QuarantineMessage|AcceptedDomain|OrganizationConfig|RemoteMailbox|User|' +
             'Place|CalendarProcessing|ExchangeCertificate|ExchangeServer|UnifiedAuditLog|' +
             'MailUser|MailMessage|PSSession|ConnectionInformation)$'

    $tokenRx = '(?<![\w\-])([A-Z][a-zA-Z]+-[A-Za-z0-9]+)'
    $resolved = @{}

    foreach ($full in $paths) {
        $leaf = Split-Path -Leaf $full
        if ($leaf -ieq 'Test-Toolkit.ps1') { continue }

        try { $raw = Get-Content -LiteralPath $full -Raw -ErrorAction Stop } catch { continue }
        $code = (Remove-PowerShellComment -Text $raw) -split "`n"

        # Here-string tracking. Remove-PowerShellComment strips comments, and a
        # here-string is a string, so every line of help text arrives looking
        # like code. That is where the first seven false positives came from --
        # "copy the sender or Message-ID", "the SMB-Access group" -- prose that
        # happens to be Verb-Noun shaped.
        #
        # Pass A still runs inside one, because $( ) in a double-quoted
        # here-string is genuinely evaluated. Only pass B is suppressed.
        $inHere  = $false
        $hereEnd = ''

        for ($i = 0; $i -lt $code.Count; $i++) {
            $line = $code[$i]
            if (-not $line) { continue }

            $wasInHere = $inHere
            if ($inHere) {
                if ($line.TrimEnd() -match ('^' + [regex]::Escape($hereEnd))) { $inHere = $false }
            } elseif ($line -match '@"\s*$') {
                $inHere = $true; $hereEnd = '"@'
            } elseif ($line -match "@'\s*$") {
                $inHere = $true; $hereEnd = "'@"
            }

            $cand = New-Object System.Collections.Generic.List[string]

            # Pass A -- inside $( ). Non-greedy to the first ')', so a nested
            # call is only partly seen; that costs coverage, never a false hit.
            foreach ($sub in [regex]::Matches($line, '\$\(([^()]*)\)')) {
                foreach ($t in [regex]::Matches($sub.Groups[1].Value, $tokenRx)) {
                    [void]$cand.Add($t.Groups[1].Value)
                }
            }

            # Pass B -- code with string literals removed. Skipped inside a
            # here-string, where every line is prose by definition.
            if (-not $wasInHere) {
                $bare = $line -replace "'[^']*'", "''" -replace '"[^"]*"', '""'
                foreach ($t in [regex]::Matches($bare, $tokenRx)) {
                    [void]$cand.Add($t.Groups[1].Value)
                }
            }

            foreach ($name in ($cand | Select-Object -Unique)) {
                $key = $name.ToLower()
                if ($defined.ContainsKey($key)) { continue }
                if ($name -match $allow) { continue }

                if (-not $resolved.ContainsKey($key)) {
                    $ok = $false
                    try { $ok = [bool](Get-Command -Name $name -ErrorAction SilentlyContinue) } catch { $ok = $false }
                    $resolved[$key] = $ok
                }
                if ($resolved[$key]) { continue }

                Add-Finding -Severity ERROR -File $full -Line ($i + 1) -Rule 'FUN003' `
                    -Message "'$name' is called here, is not defined anywhere in the toolkit, and does not resolve as a command. This throws at run time, not at parse time."
            }
        }
    }
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

        # The on-premises twin is checked the same way as Script, because it is
        # launched the same way -- and it is HARDER to notice when it is wrong.
        # A broken Script kills the button outright. A broken ScriptOnPrem is
        # invisible until an operator opens the chooser and picks Active
        # Directory, which may be days later and looks like a fault in the tool
        # rather than a typo in the config.
        #
        # Checked BEFORE the parameter contract below, deliberately: a path
        # that does not exist cannot have its param() block read, and reporting
        # "declares no parameters" about a missing file sends you to the wrong
        # place entirely.
        $twinRel = "$($t.ScriptOnPrem)".Trim()
        if ($twinRel) {
            $twinFull = Join-Path $Folder ($twinRel -replace '^\.\\', '')
            if (-not (Test-Path -LiteralPath $twinFull)) {
                Add-Finding -Severity ERROR -File $toolsPath -Line 1 -Rule 'CFG005' `
                    -Message "'$($t.Name)' names an on-premises twin '$twinRel', which does not exist. The button works until someone chooses Active Directory."
            } else {
                $twinDeclared = Get-TopLevelParameterName -File $twinFull
                if ($null -ne $twinDeclared) {
                    $known   = @($consoleParams | Where-Object { $twinDeclared -contains $_ })
                    $unknown = @($consoleParams | Where-Object { $twinDeclared -notcontains $_ })
                    if ($known.Count -gt 0 -and $unknown.Count -gt 0) {
                        Add-Finding -Severity ERROR -File $twinFull -Line 1 -Rule 'CFG004' `
                            -Message "Declares $($known -join ', ') but not $($unknown -join ', '). A script with a param() block REJECTS parameters it does not declare, so this fails on click unless the console filters."
                    }
                }
            }
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

        # The message is assembled rather than written out, because spelling the
        # sequence literally made this line match its own rule. Remove-PowerShellComment
        # strips comments but not string literals, so the text inside a finding
        # is still code as far as the scanner is concerned.
        #
        # It cost one permanent ERROR. That is worse than it sounds: -FailOn Error
        # is what gates shipping, so once the remaining files are cleaned the
        # toolkit could never report a clean run, and a check that is always red
        # is a check people learn to scroll past.
        if ($line -match '`u\{') {
            $bt = [char]0x60
            Add-Finding -Severity ERROR -File $File -Line $n -Rule 'VER001' -Message "${bt}u{...} is PowerShell 6+. In 5.1 the literal characters end up in the string."
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

#region T3b -- audit coverage ---------------------------------------------------
function Test-AuditCoverage {
    <#
        A tool that changes production and never says so.

        Common.ps1 has had Write-CompletedLog and Write-FailedLog for months,
        and the comment above them states the intent plainly: "a log that only
        records successes cannot be used to reconstruct a bad afternoon, which
        is the one job an audit trail has." The code says that. The tools did
        not do it -- Write-FailedLog was called by exactly one script out of
        sixty while six writers held sixty-six catch blocks between them.

        That is the same failure this whole file exists to catch: nothing was
        broken, nothing errored, and the gap was invisible because the only
        place it shows up is a log nobody reads until they need it. Whether an
        action was logged cannot be noticed by using the console. It can only
        be noticed by a check.

        WHY WARN AND INFO RATHER THAN ERROR. -FailOn Error is what gates
        shipping. A rule that lands red on scripts with a defensible reason not
        to log -- a preview-only tool, a report that happens to name a write
        cmdlet in a filter -- is a rule that gets switched off, and then it
        catches nothing at all. VER001 already taught that lesson once.

        STRING LITERALS AND HERE-STRINGS ARE REMOVED FIRST. Help text describing
        what a tool does is full of cmdlet names, and counting prose as code is
        how the "23 of 24 tools unguarded" miscount happened. Comments alone are
        not enough here: the ? button help lives in here-strings.
    #>
    param([string]$File)

    $raw = Get-Content -LiteralPath $File -Raw
    if (-not $raw) { return }
    $code = (Remove-PowerShellComment -Text $raw) -split "`n"

    $writeLines = New-Object System.Collections.ArrayList
    $catchCount = 0
    $inHere     = $false
    $hereEnd    = ''

    for ($i = 0; $i -lt $code.Count; $i++) {
        $line = $code[$i]

        if ($inHere) {
            if ($line -match "^\s*$([regex]::Escape($hereEnd))") { $inHere = $false }
            continue
        }
        if ($line -match '@("|'')\s*$') {
            $inHere  = $true
            $hereEnd = $Matches[1] + '@'
            continue
        }

        # Single-line string literals out. Done after here-strings so an
        # unbalanced quote inside one cannot swallow the rest of the file.
        $bare = $line -replace "'[^']*'", '' -replace '"[^"]*"', ''

        if ($bare -match '\bcatch\b') { $catchCount++ }

        if ($bare -match ('\b(Set|New|Remove|Add|Disable|Enable|Update|Move|Restore)-' +
                          '(Mailbox|ADUser|ADGroup|ADObject|ADGroupMember|ADPrincipalGroupMembership|' +
                          'DistributionGroup|DistributionGroupMember|UnifiedGroup|UnifiedGroupLinks|' +
                          'MailboxPermission|MailboxFolderPermission|RecipientPermission|ADPermission|' +
                          'InboxRule|MailboxAutoReplyConfiguration|CASMailbox|CalendarProcessing|' +
                          'MailContact|MailUser|MsolUser|MgUser|SPOSite|SPOUser|Place)\b')) {
            [void]$writeLines.Add(@{ Line = $i + 1; Cmdlet = $Matches[0] })
        }
    }

    if ($writeLines.Count -eq 0) { return }

    $flat      = $code -join "`n"
    $logsAny   = $flat -match '\bWrite-CompletedLog\b'
    $logsFail  = $flat -match '\bWrite-FailedLog\b'

    if (-not $logsAny) {
        $first = $writeLines[0]
        Add-Finding -Severity WARN -File $File -Line $first.Line -Rule 'LOG001' `
            -Message ("Calls $($writeLines.Count) write cmdlet(s) starting with $($first.Cmdlet) " +
                      "and never calls Write-CompletedLog. Changes production without leaving a record.")
        return
    }

    if (-not $logsFail -and $catchCount -gt 0) {
        Add-Finding -Severity INFO -File $File -Line 1 -Rule 'LOG002' `
            -Message ("Logs successes but never failures -- $catchCount catch block(s), no " +
                      "Write-FailedLog. A failed change is the row you will actually go looking for.")
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
        # Separate try from Test-Policy on purpose. A here-string tracker is the
        # kind of code that throws on one odd file, and losing every audit
        # finding in the folder because of it would reproduce exactly the silent
        # gap this check was added to find.
        try { Test-AuditCoverage -File $f.FullName } catch {
            Add-Finding -Severity WARN -File $f.FullName -Line 1 -Rule 'INT002' -Message "Audit-coverage check could not run: $($_.Exception.Message)"
        }
    }

    if ($Progress) { & $Progress 'Contract checks...' 98 }
    Test-Contract -Folder $Folder
    Test-TenantLeakage -Folder $Folder -Files $files
    Test-CommandResolution -Folder $Folder -Files $files
    Test-CommonContract -Folder $Folder -Files @($files | ForEach-Object { $_.FullName })

    return @($script:Findings.ToArray())
}

function Format-TestSummary {
    param($Findings, [string]$Folder, [int]$FileCount)

    $e = @($Findings | Where-Object { $_.Severity -eq 'ERROR' })
    $w = @($Findings | Where-Object { $_.Severity -eq 'WARN'  })
    $n = @($Findings | Where-Object { $_.Severity -eq 'INFO'  })

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Folder : $Folder")
    [void]$sb.AppendLine("Files  : $FileCount")
    [void]$sb.AppendLine("Result : $($e.Count) error(s), $($w.Count) warning(s), $($n.Count) note(s)")
    [void]$sb.AppendLine('')

    # INFO is rendered, and rendered on the clean path too.
    #
    # Add-Finding has always accepted INFO and the report has always dropped it:
    # the loop below listed ERROR and WARN only, and a clean run returned before
    # reaching it. So every INFO ever raised -- FUN002, and the TEN000 line added
    # specifically to prove the leakage rule had run -- was recorded and never
    # shown. A diagnostic you cannot see is not a diagnostic, and this is the
    # same "the check passed because it could not report" fault it was meant to
    # catch, one layer up in the tool doing the catching.
    #
    # The clean path matters most. "Clean. Safe to ship." is exactly when you
    # need to know WHICH checks ran to earn that word.
    if ($e.Count -eq 0 -and $w.Count -eq 0) {
        [void]$sb.AppendLine('Clean. Safe to ship.')
        [void]$sb.AppendLine('')
    }

    foreach ($group in @('ERROR', 'WARN', 'INFO')) {
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
    [void]$sb.AppendLine('  LOG001  a preview-only or -WhatIf tool may legitimately make no record.')
    [void]$sb.AppendLine('  LOG002  a catch that only rethrows does not need its own log row.')
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
