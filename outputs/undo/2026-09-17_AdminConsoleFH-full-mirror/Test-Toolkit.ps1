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

    $all = @(Get-ChildItem -LiteralPath $Folder -Filter '*.ps1' -File -Recurse -ErrorAction SilentlyContinue |
             Sort-Object FullName)

    # Projects\ is EXCLUDED, and that is a decision rather than an oversight.
    #
    # It holds migration work: one-off scripts written against a particular
    # cutover, kept because they record what was done. They fail the encoding
    # and logging rules by the dozen and every one of those findings is true
    # and none of them is going to be actioned -- which is the definition of
    # noise. Fifty rows nobody reads is how the twenty rows that matter get
    # skipped.
    #
    # The exclusion is ANNOUNCED at the top of the results, not silent. A
    # scanner that quietly skips a folder is how this tool once reported "no
    # errors" over eleven files it had never opened.
    # Which folders actually EXCLUDED something, not merely which were listed.
    # obsolete\ lives OUTSIDE the toolkit in a normal install, so naming it in a
    # report that skipped nothing there would describe a folder that is not
    # present -- and a notice you cannot check is one you stop reading.
    $script:ExcludedHit = New-Object System.Collections.ArrayList
    $kept = @($all)
    foreach ($name in $script:ExcludedFolders) {
        $skip   = [System.IO.Path]::Combine($Folder, $name) + [System.IO.Path]::DirectorySeparatorChar
        $before = $kept.Count
        $kept   = @($kept | Where-Object {
            -not $_.FullName.StartsWith($skip, [StringComparison]::OrdinalIgnoreCase)
        })
        # Only ANNOUNCED folders feed the notice -- name AND file count. Silent
        # ones are still skipped, they just are not talked about. Counting them
        # separately matters: the notice says "19 file(s) not checked", and that
        # number has to describe the folders it just named or it is a number
        # nobody can reconcile.
        $removed = $before - $kept.Count
        if ($removed -gt 0 -and $script:AnnouncedFolders -contains $name) {
            [void]$script:ExcludedHit.Add($name)
            $script:ExcludedFileCount += $removed
        }
    }

    return $kept
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
        # only, so 'E5 License B - Full' sat hardcoded in two scripts that
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
    # NO 'the check ran' NOTE. Deliberately silent on the normal path.
    #
    # This used to announce how many names TEN002 was matching, as proof it had
    # not silently checked nothing. TEN003 above already covers that: zero
    # names is an ERROR. What was left only caught PARTIAL degradation -- 5
    # names instead of 23 -- and spotting that requires remembering what the
    # number should be. A diagnostic you need a memorised baseline to read is
    # one you skim past for six months and then miss anyway.
    #
    # If the partial case ever bites, the honest version compares against the
    # PREVIOUS run's count from the report CSV and speaks only when it drops.
    # Silent when normal, loud when it changes, nothing to remember.


    # Anything shaped like an internal FQDN or a vanity SharePoint host, plus
    # whatever this deployment's own Tenant.json names.
    $shapes = @(
        '\b[a-z0-9-]+\.(?:franchises|corp|internal|ad|local)\.[a-z.]+\b'
        '\b[a-z0-9-]+-(?:my|admin)\.sharepoint\.com\b'
        '\b[a-z0-9-]+\.mail\.onmicrosoft\.com\b'
        '\b[a-z0-9-]+\.service-now\.com\b'
    )

    # Projects\ is exempt for the same reason EX- is: the path is a declaration
    # that the file belongs to one organisation's body of work and is pulled
    # before the toolkit ships elsewhere. The difference is that EX- was a
    # decision and this was not -- these files simply never produced a finding,
    # and an exemption nobody chose is indistinguishable from a rule that is
    # broken. Counting them makes the exemption something you can see and argue
    # with.
    #
    # SAID AS Projects\<Tenant>\ UNTIL 2026-08-13, WHICH IT NEVER WAS. The test
    # below is '^Projects[\\/]' -- one segment, no tenant level. The comment
    # described a folder layout that happened to exist rather than the rule
    # actually enforced, so when the tenant folder was removed the code was
    # unaffected and only the comment was wrong. A comment that describes the
    # data instead of the rule is a comment that expires without warning.
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

                # contoso is the DOCUMENTATION example domain and is used on
                # purpose -- in placeholder text, in help, in the tenant
                # template. It names nobody and leaks nothing. Flagging it
                # teaches people to stop writing examples, and a toolkit whose
                # own hint text trips its own linter is one nobody trusts.
                if ($m.Success -and $m.Value -match '(?i)contoso') { continue }

                if ($m.Success) {
                    Add-Finding -Severity ERROR -File $full -Line ($i + 1) -Rule 'TEN001' `
                        -Message "Tenant-specific value '$($m.Value)' in a script that ships to every tenant. Move it to Tenant.json, or rename the file EX- if it genuinely cannot be generalised."
                    break
                }
            }
            foreach ($n in $names) {
                if ($n.Length -lt 4) { continue }

                # WORD BOUNDARY, not substring. A TenantName that is a common word
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

    # NO PROJECTS NOTE HERE EITHER. SKP001 already reports that folder as not
    # checked, with a file count, at the top of every run. Two rules saying the
    # same thing in different words is how a report stops being read.
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
    # The second half of this list was added 2026-09-12. Every entry is a real
    # server cmdlet that exists only once a session is connected -- Compliance*
    # from Security and Compliance PowerShell (Search and Purge), the
    # connectors and filter policies from Exchange Online (Mail Flow
    # Connectors, Spam Policy Viewer), the migration and free/busy families
    # from the tools added the same day. All of them were reported as "does
    # not resolve" on a machine with nothing connected, which is exactly what
    # this list exists to prevent. QuarantineMessage gained its wildcard so
    # Get-QuarantineMessageHeader is covered with it.
    $allow = '(?i)-(AD[A-Za-z0-9]*|SPO[A-Za-z0-9]*|Mg[A-Za-z0-9]*|EXO[A-Za-z0-9]*|' +
             'Mailbox[A-Za-z0-9]*|DistributionGroup[A-Za-z0-9]*|UnifiedGroup[A-Za-z0-9]*|' +
             'Recipient[A-Za-z0-9]*|InboxRule|CASMailbox|MessageTrace[A-Za-z0-9]*|' +
             'QuarantineMessage[A-Za-z0-9]*|AcceptedDomain|OrganizationConfig|RemoteMailbox|User|' +
             'Place|CalendarProcessing|ExchangeCertificate|ExchangeServer|UnifiedAuditLog|' +
             'MailUser|MailMessage|PSSession|ConnectionInformation|' +
             'ComplianceSearch[A-Za-z0-9]*|InboundConnector|OutboundConnector|' +
             'MoveRequest[A-Za-z0-9]*|MigrationBatch|MigrationUser|MigrationEndpoint|' +
             'DkimSigningConfig|OrganizationRelationship|IntraOrganizationConnector|' +
             'IntraOrganizationConfiguration|SharingPolicy|HostedContentFilterPolicy|' +
             'HostedContentFilterRule|AntiPhishPolicy|AntiPhishRule|' +
             'HostedOutboundSpamFilterPolicy|HostedOutboundSpamFilterRule|' +
             'DynamicDistributionGroup[A-Za-z0-9]*|DnsName)$'

    # PROSE THAT LOOKS LIKE A CMDLET. 'Send-As' and 'Send-To' are permissions,
    # not commands, and they appear in captions, audit text and log lines --
    # 'Send' is an approved verb, so nothing about the shape of the token says
    # otherwise. The docstring above already named this as the failure that
    # gets a rule switched off; the two string passes catch most of it and
    # this catches the rest. Matching is on the whole token, case-insensitive.
    $prose = @('Send-As', 'Send-To', 'Send-On', 'On-Prem', 'On-Premises', 'Free-Busy',
               'Read-Only', 'Mail-Enabled', 'Non-Sync', 'Out-Of', 'E-Mail', 'Add-In',
               'Full-Access', 'Sign-In', 'Follow-Up', 'End-User', 'Self-Service')

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
                if ($prose -contains $name) { continue }
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
# The folders left out of every scan, and how many files that came to.
# Named once so the filter and the notice can never describe different things.
#
# Projects       migration work, kept as a record of a cutover
# obsolete       superseded versions, kept so a change can be undone
# DisabledTools  parked by Tool Status, kept rather than deleted
#
# All three hold code that is deliberately not maintained. Linting a file you
# have decided to stop maintaining is a treadmill -- it produces true findings
# nobody will action, and the backup discipline that makes obsolete\ useful
# would make the report useless within a week.
# TWO KINDS OF EXCLUSION, AND THE DIFFERENCE IS WHETHER YOU NEED REMINDING.
#
# ANNOUNCED -- work that is paused, not finished. Projects\ holds a body of
# work still to be brought onto this architecture, and the whole point of the
# notice is that you have not forgotten it. Skipping it silently is how a scan
# once reported "no errors" over eleven files it had never opened.
#
# SILENT -- code that is deliberately out of service. A tool in DisabledTools\
# has had its button removed on purpose; obsolete\ is kept only so it can be
# read later. Reporting either every run is noise about a decision already
# made, and noise is what stops the announced line being read.
$script:AnnouncedFolders  = @('Projects')
$script:SilentFolders     = @('obsolete', 'DisabledTools')
$script:ExcludedFolders   = @($script:AnnouncedFolders + $script:SilentFolders)
$script:ExcludedFileCount = 0
$script:ExcludedHit       = @()

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

$script:RuleCatalogue = @{
    # rule   = @{ Impact = ...; Plain = ... }
    #
    # PLAIN is written for someone who has never opened this toolkit: what
    # actually goes wrong, in one sentence, no jargon. The technical message
    # still exists -- it is on the Details tab. This is the "so what".
    #
    # IMPACT is about CONSEQUENCE, not about how hard it is to fix:
    #   Critical  it is broken now, or will break the moment it is used
    #   High      it will do the wrong thing, quietly
    #   Medium    it will fail visibly, or a safety net is missing
    #   Low       it is untidy and will bite someone eventually
    #   Cosmetic  it looks wrong; nothing misbehaves
    #   Note      not a fault at all -- something the run wants you to know
    #
    # Severity (ERROR/WARN/INFO) is what the CHECK thinks. Impact is what it
    # means for you. They are deliberately not the same axis: TEN001 is only a
    # WARN to the scanner and High to you, because a tenant name in shared code
    # is silently wrong in five other tenants.

    SYN000 = @{ Impact='Critical'; Plain='This file has a syntax error. It will not run at all.' }
    SYN001 = @{ Impact='Critical'; Plain='This file has a syntax error. It will not run at all.' }
    FUN003 = @{ Impact='Critical'; Plain='Calls something that does not exist anywhere. Crashes the moment that line is reached.' }
    CFG002 = @{ Impact='Critical'; Plain='Tools.json is not valid JSON, so the console cannot read its buttons.' }
    ENC001 = @{ Impact='Critical'; Plain='Missing the file marker PowerShell needs. On another machine this becomes a different file.' }
    ENC002 = @{ Impact='Critical'; Plain='File marker written twice. It puts an invisible character on line 1 and breaks the version check.' }
    ENC004 = @{ Impact='High';     Plain='Characters have already been lost saving this file. Some text is now wrong.' }
    CFG003 = @{ Impact='High';     Plain='A button points at a file that is not there. Clicking it does nothing useful.' }
    CFG006 = @{ Impact='Medium';   Plain='A parked-in-place entry names a script that is not there, so it parks nothing -- the tool it was meant to hide may still be on that tenant''s console.' }
    CFG004 = @{ Impact='High';     Plain='The console passes this tool something it refuses to accept, so it fails on click.' }
    FUN001 = @{ Impact='High';     Plain='Keeps its own copy of a shared function. Fixes made in Common.ps1 never reach it.' }
    VER001 = @{ Impact='High';     Plain='Uses newer PowerShell than 5.1. Here it does not error, it puts the wrong text in the string.' }
    REF001 = @{ Impact='High';     Plain='Sets a property on something before it is created. That setting is thrown away.' }
    TEN001 = @{ Impact='High';     Plain='A name specific to this company sits in a file that ships to all six tenants.' }
    TEN002 = @{ Impact='High';     Plain='A value from Tenant.json is hardcoded here, so it is wrong everywhere else.' }
    TEN003 = @{ Impact='High';     Plain='The tenant-leak check could not run this time, so nothing was actually checked.' }
    ENC003 = @{ Impact='Medium';   Plain='Line endings are the wrong kind. The file shows as one long line in some editors.' }
    LOG001 = @{ Impact='Medium';   Plain='Changes real things but never writes to the audit log. No record of what it did.' }
    HDL001 = @{ Impact='Medium';   Plain='A button talks to Exchange or AD with no error handling, so a failure shows an unreadable message.' }
    VER002 = @{ Impact='Medium';   Plain='Looks like newer PowerShell than 5.1, which would error here.' }
    PRC001 = @{ Impact='Medium';   Plain='Launches another program without quoting. A path with a space arrives broken in two.' }
    CFG001 = @{ Impact='Medium';   Plain='Tools.json was not found, so the button checks were skipped entirely.' }
    INT001 = @{ Impact='Medium';   Plain='A policy check could not run, so that part of the toolkit went unchecked.' }
    INT002 = @{ Impact='Medium';   Plain='The audit-log coverage check could not run, so that went unchecked.' }
    LOG002 = @{ Impact='Low';      Plain='Records what worked but not what failed -- and the failure is the row you go looking for.' }
    FUN002 = @{ Impact='Low';      Plain='Has its own copy of a shared function on purpose. Fine, but it will drift over time.' }
    EXO002 = @{ Impact='Cosmetic'; Plain='Exchange prints a "more results available" warning on every run. Harmless, just noisy.' }
    SKP001 = @{ Impact='Note';     Plain='Part of the folder was deliberately not checked. Listed so it is not forgotten.' }
}

function Get-RuleImpact {
    param([string]$Rule)
    $e = $script:RuleCatalogue[$Rule]
    if ($e) { return "$($e.Impact)" }
    # An unknown rule is treated as Medium rather than skipped. A finding with
    # no catalogue entry is a rule someone added without describing it, and
    # hiding it would make the grid quietly less complete than the Details tab.
    return 'Medium'
}

function Get-RulePlain {
    param([string]$Rule, [string]$Fallback)
    $e = $script:RuleCatalogue[$Rule]
    if ($e) { return "$($e.Plain)" }
    return $Fallback
}

# Sort order for the grid. Not alphabetical -- worst first, because the top of
# the list is the only part guaranteed to be read.
$script:ImpactRank = @{
    'Critical' = 0; 'High' = 1; 'Medium' = 2; 'Low' = 3; 'Cosmetic' = 4; 'Note' = 5
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
        # ENC005 -- plain non-ASCII -- was removed on 2026-08-13. It was a
        # convention check, not a correctness one: a smart quote in a comment
        # is untidy and the script still runs, so it produced findings nobody
        # was ever going to act on.
        #
        # ENC004 stays and is the one that matters. U+FFFD is not a style
        # choice, it is a character that has ALREADY been destroyed by a bad
        # encoding round-trip -- the file has lost information and the text is
        # now wrong.
        if ($line.IndexOf([char]0xFFFD) -ge 0) {
            Add-Finding -Severity ERROR -File $File -Line $n -Rule 'ENC004' -Message 'Replacement character (U+FFFD) -- this file has already lost characters to a bad encoding round-trip.'
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

    # ---- parked-in-place lists ---------------------------------------------
    #
    # EVERY TENANT'S LIST, NOT THE ACTIVE ONE. This script runs from a prompt
    # with no console open and therefore no active tenant, so asking
    # Get-ParkedListPath would check one file and call the estate clean.
    #
    # WHAT IS WORTH CHECKING is an entry naming a script that is not there.
    # Parking matches on the FILE NAME, so a renamed or deleted script leaves a
    # dead entry that parks NOTHING -- and the tool it was meant to hide is
    # quietly back on that tenant's console. Nothing else reports that: the
    # console cannot tell 'parked' from 'never mentioned', by design.
    #
    # A PARKED SCRIPT STILL LISTED IN Tools.json IS NOT A FINDING. That is how
    # parking works: one Tools.json for every tenant, filtered per tenant when
    # the console loads it.
    #
    # WHERE THE LIST LIVES: a 'Parked' block inside each tenant file
    # (Config\Tenants\*.json, or Config\Tenant.json single-tenant), shaped
    # { Scripts, Sections }. The older side files (*.parked.json,
    # Config\Parked.json) are still read, because a tenant that has not saved
    # since the change is still being driven by one.
    $parkFiles = @()
    $tenantsDir = Join-Path $Folder 'Config\Tenants'
    if (Test-Path -LiteralPath $tenantsDir) {
        $parkFiles += @(Get-ChildItem -LiteralPath $tenantsDir -Filter '*.json' -File -ErrorAction SilentlyContinue |
                        Where-Object { $_.Name -ne 'Tenants.json' })
    }
    foreach ($nm in @('Config\Tenant.json', 'Config\Parked.json')) {
        $sp = Join-Path $Folder $nm
        if (Test-Path -LiteralPath $sp) { $parkFiles += @(Get-Item -LiteralPath $sp) }
    }

    foreach ($pf in $parkFiles) {
        $pj = $null
        try {
            $pj = Get-Content -LiteralPath $pf.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        } catch {
            Add-Finding -Severity ERROR -File $pf.FullName -Line 1 -Rule 'CFG002' -Message "Could not be parsed: $($_.Exception.Message)"
            continue
        }
        if (-not $pj -or -not $pj.PSObject.Properties['Parked'] -or -not $pj.Parked) { continue }
        # Tenant file: Parked is { Scripts, Sections }. Side file: Parked IS the array.
        $entries = @()
        if ($pj.Parked -is [System.Array]) { $entries = @($pj.Parked) }
        elseif ($pj.Parked.PSObject.Properties['Scripts']) { $entries = @($pj.Parked.Scripts) }
        foreach ($pe in $entries) {
            if ($null -eq $pe) { continue }
            $ps = if ($pe.PSObject.Properties['Script']) { "$($pe.Script)" } else { "$pe" }
            $ps = $ps.Trim()
            if (-not $ps) { continue }
            # Relative to the toolkit root, not the leaf: project scripts live in
            # Projects\<name>\ and reuse file names across projects.
            $prel = ($ps -replace '/', '\') -replace '^(\.\\)+', ''
            if (-not (Test-Path -LiteralPath (Join-Path $Folder $prel))) {
                Add-Finding -Severity WARN -File $pf.FullName -Line 1 -Rule 'CFG006' `
                    -Message "Parked entry '$ps' matches no script in the toolkit folder -- it parks nothing."
            }
        }
    }

    # ---- per-user layout profiles -------------------------------------------
    #
    # SAME QUESTION AS THE PARKED LISTS, asked of a different file. A profile
    # can park a script and can place a tool into a section; both go stale the
    # same way, and both fail silently -- a dead parked entry hides nothing,
    # and a placement into a section that no longer exists puts the tool
    # wherever the console's fallback sends it.
    #
    # EVERY PROFILE IN THE FOLDER, not the current operator's. This runs from a
    # prompt and is checking the estate, so asking Get-UserProfilePath would
    # check one person and call the rest clean.
    $usersDir = Join-Path $Folder 'Config\Users'
    if (Test-Path -LiteralPath $usersDir) {
        # Section names come from the shared file plus whatever each profile
        # declares for itself; a placement is only wrong if neither knows it.
        $sharedSections = @()
        try {
            $sp = Join-Path $Folder 'Config\Sections.json'
            if (Test-Path -LiteralPath $sp) {
                $sj = Get-Content -LiteralPath $sp -Raw -Encoding UTF8 | ConvertFrom-Json
                $sharedSections = @(@($sj.Sections) | ForEach-Object { "$($_.Name)".Trim().ToLowerInvariant() })
            }
        } catch { }

        foreach ($uf in @(Get-ChildItem -LiteralPath $usersDir -Filter '*.json' -File -ErrorAction SilentlyContinue)) {
            $uj = $null
            try {
                $uj = Get-Content -LiteralPath $uf.FullName -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            } catch {
                Add-Finding -Severity ERROR -File $uf.FullName -Line 1 -Rule 'CFG002' -Message "Layout profile could not be parsed: $($_.Exception.Message)"
                continue
            }
            if (-not $uj) { continue }

            # Default plus every tenant block: each is a layout in its own right.
            $blocks = @()
            if ($uj.PSObject.Properties['Default'] -and $uj.Default) { $blocks += $uj.Default }
            if ($uj.PSObject.Properties['Tenants'] -and $uj.Tenants) {
                foreach ($tp in @($uj.Tenants.PSObject.Properties)) { if ($tp.Value) { $blocks += $tp.Value } }
            }

            foreach ($blk in $blocks) {
                $known = @($sharedSections)
                if ($blk.PSObject.Properties['Sections']) {
                    $known += @(@($blk.Sections) | ForEach-Object { "$($_.Name)".Trim().ToLowerInvariant() })
                }
                if ($blk.PSObject.Properties['Placement'] -and $blk.Placement) {
                    foreach ($pp in @($blk.Placement.PSObject.Properties)) {
                        $rel = "$($pp.Name)" -replace '/', '\'
                        if ($rel -and -not (Test-Path -LiteralPath (Join-Path $Folder $rel))) {
                            Add-Finding -Severity WARN -File $uf.FullName -Line 1 -Rule 'CFG006' `
                                -Message "Layout profile places '$rel', which is not in the toolkit folder -- that entry does nothing."
                        }
                        $sec = ''
                        try { $sec = "$($pp.Value.Section)".Trim().ToLowerInvariant() } catch { }
                        if ($sec -and $known -notcontains $sec) {
                            Add-Finding -Severity WARN -File $uf.FullName -Line 1 -Rule 'CFG006' `
                                -Message "Layout profile puts '$rel' in section '$($pp.Value.Section)', which no section list declares."
                        }
                    }
                }
                if ($blk.PSObject.Properties['Parked'] -and $blk.Parked) {
                    foreach ($pe in @($blk.Parked.Scripts)) {
                        if ($null -eq $pe) { continue }
                        $ps2 = $(if ($pe.PSObject.Properties['Script']) { "$($pe.Script)" } else { "$pe" }).Trim()
                        if (-not $ps2) { continue }
                        $prel2 = ($ps2 -replace '/', '\') -replace '^(\.\\)+', ''
                        if (-not (Test-Path -LiteralPath (Join-Path $Folder $prel2))) {
                            Add-Finding -Severity WARN -File $uf.FullName -Line 1 -Rule 'CFG006' `
                                -Message "Layout profile parks '$ps2', which is not in the toolkit folder -- it parks nothing."
                        }
                    }
                }
            }
        }
    }

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

        # The AD twin check (CFG004/CFG005) that used to sit here was removed
        # on 2026-08-13 with the ScriptOnPrem feature itself. See ARCHITECTURE
        # section 30. A Tools.json still carrying ScriptOnPrem is not an error
        # now -- the console ignores the field, and ToolManager drops it on the
        # next Save.

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

        # ---- FUN004: calls Common.ps1 without importing it -------------------
        #
        # THE OPPOSITE MISTAKE TO FUN001, AND A WORSE ONE. FUN001 is a local
        # copy shadowing a shared function -- the tool still runs, it just runs
        # the stale version. This is a call to a function that DOES NOT EXIST in
        # the process, so the tool dies at the moment it reaches the line, with
        # "'X' is not recognized as the name of a cmdlet" -- which reads like a
        # broken install rather than a missing import, and sends the operator
        # looking at PowerShell rather than at the script.
        #
        # It arrives by one route: a standalone script gains a feature that
        # needs shared code, the call is added, and the dot-source is not.
        # Add-ConsoleAdditions.ps1 did exactly that when it learned about
        # parked scripts -- three calls, no import, and it failed on the first
        # parked script anybody had (2026-09-17).
        #
        # COMMENTS AND STRINGS ARE STRIPPED FIRST. Every one of these files
        # names Common functions in its docstrings; matching those would report
        # every script in the toolkit and the rule would be switched off within
        # a day.
        # A VIEW FILE IS NOT STANDALONE AND MUST NOT IMPORT COMMON.
        #
        # Views\View_*.ps1 are DOT-SOURCED BY THE CONSOLE, into a scope where
        # Common.ps1 was loaded long before -- so every one of them calls
        # Set-ConsoleTheme and friends without an import, correctly, by design.
        # Without this exemption the rule reports all eight of them as errors,
        # which is the "twelve findings, of which two were real" failure this
        # file already learned once and wrote down twenty lines above.
        #
        # RECOGNISED BY THE HEADER, NOT THE FOLDER. A view is a file carrying a
        # VIEW-FUNCTION declaration; that is what makes the console load it, so
        # it is the honest test. A path check would miss one moved elsewhere
        # and would wrongly exempt an ordinary script dropped into Views\.
        $isView = [regex]::IsMatch($text, '(?m)^\s*#\s*VIEW-FUNCTION:\s*\S')

        if (-not $loadsCommon -and -not $isView) {
            $code = $text -replace '(?s)<#.*?#>', ''
            $code = ($code -split "`r?`n" | ForEach-Object { $_ -replace '#.*$', '' }) -join "`n"
            $code = $code -replace "'[^']*'", "''" -replace '"[^"]*"', '""'

            $called = @{}
            foreach ($m in [regex]::Matches($code, '(?<![-\w$.])([A-Z][a-zA-Z]+-[A-Za-z0-9]+)')) {
                $n = $m.Groups[1].Value
                if ($own -contains $n) { continue }
                if ($commonFns -notcontains $n) { continue }
                $called[$n] = $true
            }
            if ($called.Count -gt 0) {
                $list = (@($called.Keys) | Sort-Object) -join ', '
                Add-Finding -Severity ERROR -File $f -Line 1 -Rule 'FUN004' `
                    -Message "Calls $list from Common.ps1 but never dot-sources it. This fails at RUN TIME with 'is not recognized as the name of a cmdlet', on the first line that reaches one. Add: . (Join-Path `$PSScriptRoot 'Common.ps1')"
            }
        }

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

function Test-CommonParameterName {
    <#
        Every -Parameter used when calling a Common.ps1 function must be one
        that function declares.

        WHY THIS EXISTS. On 2026-08-25 ProtocolUsage.ps1 shipped with
        Start-UrlInBrowserProfile -Admin. The parameter is -Purpose; -Admin does
        not exist. Parse passed, the balance check passed, every rule in this
        file passed, and the operator got "a parameter cannot be found that
        matches parameter name admin" the first time they pressed the button.
        PowerShell binds parameters at RUN time, so a wrong name is invisible to
        everything that runs before then -- unless something reads the param()
        block and checks.

        WHAT IT CHECKS. For each Common.ps1 function, the set of names in its
        param() block. For each call to that function in a tool -- the function
        name followed by one or more -Name tokens on the same logical line --
        every -Name must be in that set. Common parameters (-ErrorAction,
        -Verbose and the rest) are always allowed. Splatted calls (@p) are
        skipped: the names are in a hashtable this rule does not follow.

        WHAT IT DOES NOT CHECK. Positional arguments, parameter TYPES, and
        mandatory parameters that were omitted. It answers exactly one
        question -- "is this a real parameter name" -- because that is the one
        that reached production.

        PREFIX ABBREVIATION IS HONOURED. PowerShell accepts -Purp for -Purpose
        when unambiguous, so a used name matches if it is a prefix of exactly
        one declared name. -Admin is not a prefix of anything, which is the
        point.
    #>
    param([string]$Folder, [string[]]$Files)

    $commonPath = Join-Path $Folder 'Common.ps1'
    if (-not (Test-Path -LiteralPath $commonPath)) { return }
    $commonText = Get-Content -LiteralPath $commonPath -Raw

    # Function name -> declared parameter names, READ FROM THE AST.
    #
    # THIS WAS A REGEX AND THE REGEX WAS WRONG, in a way that made the whole
    # rule worthless. It found 'param(' then walked for the matching ')'
    # starting ONE CHARACTER SHORT -- on the '(' itself -- so the depth
    # counter opened at 1, sailed past the real closing paren, and read to the
    # end of the FUNCTION BODY. Then it took names from '$x' followed by '='
    # or a comma, which meant every local variable in the body looked like a
    # parameter and every parameter written '[string]$Text)' -- closing paren
    # on the same line -- was missed.
    #
    # The result on 2026-09-12: Get-TargetSmtp reported "Declared: -t" when it
    # declares -Text, Get-ConfigFilePath reported "(none)" when it declares
    # -Name, and the run produced ~180 ERRORs of which ZERO were real. That is
    # exactly the ratio Test-CommonContract's own docstring warns about: a
    # rule nobody can trust is worse than no rule, because it still fills the
    # report.
    #
    # PowerShell already has a parser and this file already uses it (see
    # Get-TopLevelParameterName). It knows what a parameter is; a regex is
    # guessing. Nested functions are included -- a call site cannot tell the
    # difference, and neither should this.
    $declared = @{}
    $ctErrors = $null
    $ctAst = [System.Management.Automation.Language.Parser]::ParseFile($commonPath, [ref]$null, [ref]$ctErrors)
    if ($ctAst) {
        foreach ($fn in $ctAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
            $ps = @()
            # Both spellings: 'function f { param($a) }' and 'function f($a) {}'.
            if ($fn.Body -and $fn.Body.ParamBlock) { $ps = @($fn.Body.ParamBlock.Parameters) }
            elseif ($fn.Parameters) { $ps = @($fn.Parameters) }
            $declared[$fn.Name] = @($ps | ForEach-Object { $_.Name.VariablePath.UserPath } | Select-Object -Unique)
        }
    }

    $commonParams = @('Debug','ErrorAction','ErrorVariable','InformationAction','InformationVariable',
                      'OutBuffer','OutVariable','PipelineVariable','Verbose','WarningAction','WarningVariable',
                      'WhatIf','Confirm','ProgressAction')

    # Operators look like parameters to a regex and are not. -join after a
    # function call is the pipeline's business, not the function's.
    $operators = @('join','split','eq','ne','gt','lt','ge','le','like','notlike','match','notmatch',
                   'contains','notcontains','in','notin','and','or','not','xor','band','bor','bxor','bnot',
                   'replace','creplace','ireplace','f','is','isnot','as','shl','shr','ceq','cne','ieq','ine',
                   'clike','cnotlike','cmatch','cnotmatch','ccontains','cin','cnotin')

    # ONE REGEX PER LINE, NOT ONE PER FUNCTION PER LINE. This loop used to
    # test every line against every Common.ps1 function name in turn --
    # 120 files x ~700 lines x ~250 names, twenty million regex calls -- and
    # "Contract checks..." sat at 98% for minutes. Reported 2026-09-12. A
    # single alternation of all the names finds the (few) candidates on a
    # line in one pass; the per-function work below then runs only for those.
    # Same matches, same findings, a hundred times fewer regex evaluations.
    $anyFnRx = $null
    if ($declared.Count -gt 0) {
        $alt = ($declared.Keys | Sort-Object { -$_.Length } | ForEach-Object { [regex]::Escape($_) }) -join '|'
        $anyFnRx = [regex]"(?:^|[^A-Za-z0-9_-])($alt)(?:\s|$)"
    }
    foreach ($f in $Files) {
        if ((Split-Path -Leaf $f) -ieq 'Common.ps1') { continue }
        if (-not $anyFnRx) { break }
        $lines = Get-Content -LiteralPath $f

        # A FILE THAT DEFINES THE NAME ITSELF IS NOT CALLING COMMON'S.
        #
        # Several tools carry their own copy of a Common.ps1 function -- which
        # FUN001 reports, and rightly -- and the local copy is what the calls
        # in that file actually bind to. Checking those calls against COMMON's
        # parameter list is checking them against a function they never reach.
        #
        # On 2026-09-12 that produced two confident errors against working
        # code: AuditUser.ps1 calling -ProgressCallback (its own copy declares
        # it; Common's says -OnProgress) and UserDefinedDelegates.ps1 calling
        # -Identity (its own declares it; Common's says -Smtp). Both would run
        # fine and both were reported as run-time failures.
        #
        # So: skip any name this file defines. The divergence is still
        # reported -- by FUN001, which is the rule that is actually about it.
        $localFn = @{}
        foreach ($m in [regex]::Matches(($lines -join "`n"), '(?mi)^\s*function\s+([A-Za-z0-9_\-]+)')) {
            $localFn[$m.Groups[1].Value.ToLower()] = $true
        }
        $inBlock = $false
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $l = $lines[$i]
            # Skip block comments and line comments -- a docstring that says
            # "call X -Foo" is prose, not a call.
            if ($l -match '<#') { $inBlock = $true }
            if ($inBlock) { if ($l -match '#>') { $inBlock = $false }; continue }
            # Cheapest test first: a line with no dash-Name cannot misuse one.
            if ($l.IndexOf(' -') -lt 0) { continue }
            $code = ($l -replace "'[^']*'", "''") -replace '"[^"]*"', '""'
            $code = $code -replace '#.*$', ''
            $hits = @($anyFnRx.Matches($code) | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique)
            if ($hits.Count -eq 0) { continue }
            foreach ($fn in $hits) {
                if ($localFn[$fn.ToLower()]) { continue }
                if ($code -match '@[A-Za-z_]') { continue }   # splatted -- not followed
                $after = $code.Substring($code.IndexOf($fn) + $fn.Length)

                # THE CALL ENDS WHERE ITS ENCLOSING EXPRESSION ENDS.
                #
                # (Get-LogTarget $x) -Result 'Success' is a call to Get-LogTarget
                # inside an argument to something else, and -Result belongs to
                # that something else. Reading past the ')' that closes the
                # call blamed every nested Get-LogTarget for its caller's
                # parameters -- 19 findings on the first run, 19 wrong. Stop at
                # the first unmatched ')', or at a pipe or statement break.
                $depth = 0; $end = $after.Length
                for ($k = 0; $k -lt $after.Length; $k++) {
                    $ch = $after[$k]
                    if ($ch -eq '(') { $depth++ }
                    elseif ($ch -eq ')') { if ($depth -eq 0) { $end = $k; break }; $depth-- }
                    elseif (($ch -eq '|' -or $ch -eq ';') -and $depth -eq 0) { $end = $k; break }
                }
                $after = $after.Substring(0, $end)

                # And the mirror case: -Leaf inside (Split-Path -Leaf $x) is an
                # argument to Split-Path, not to the function being checked.
                # Anything inside a balanced pair belongs to whatever is inside
                # the pair. Blank it, keeping the length so nothing shifts.
                $sb = New-Object System.Text.StringBuilder($after)
                $depth = 0
                for ($k = 0; $k -lt $sb.Length; $k++) {
                    $ch = $sb[$k]
                    if ($ch -eq '(') { $depth++; continue }
                    if ($ch -eq ')') { $depth--; continue }
                    if ($depth -gt 0) { $sb[$k] = ' ' }
                }
                $after = $sb.ToString()

                foreach ($pm in [regex]::Matches($after, '\s-([A-Za-z][A-Za-z0-9]*)\b')) {
                    $used = $pm.Groups[1].Value
                    if ($commonParams -contains $used) { continue }
                    if ($operators -contains $used.ToLowerInvariant()) { continue }
                    $ok = @($declared[$fn] | Where-Object { $_ -like "$used*" })
                    if ($ok.Count -eq 1 -or ($declared[$fn] -contains $used)) { continue }
                    if ($ok.Count -gt 1 -and ($declared[$fn] -notcontains $used)) {
                        Add-Finding -Severity WARN -File $f -Line ($i + 1) -Rule 'PRM002' -Message "-$used is ambiguous for $fn (matches $($ok -join ', ')). PowerShell will refuse it at run time."
                        continue
                    }
                    Add-Finding -Severity ERROR -File $f -Line ($i + 1) -Rule 'PRM001' -Message "$fn has no parameter -$used. Declared: $(if ($declared[$fn].Count) { ($declared[$fn] | ForEach-Object { "-$_" }) -join ' ' } else { '(none)' }). This fails at run time, not at parse."
                }
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

    # Every function this FILE defines, plus every function Common.ps1
    # defines. HDL001 below needs it to tell a real Exchange cmdlet from a
    # local wrapper that is named like one. Common's list is read once per
    # run; this file's has to be per file and is cheap.
    $script:PolicyDefinedFn = @{}
    foreach ($m in [regex]::Matches($raw, '(?mi)^\s*function\s+([A-Za-z0-9_\-]+)')) {
        $script:PolicyDefinedFn[$m.Groups[1].Value.ToLower()] = $true
    }
    if ($null -eq $script:PolicyCommonFn) {
        $script:PolicyCommonFn = @{}
        try {
            $cp = Join-Path "$script:TestRoot" 'Common.ps1'
            if (Test-Path -LiteralPath $cp) {
                $ct = Get-Content -LiteralPath $cp -Raw -ErrorAction Stop
                foreach ($m in [regex]::Matches($ct, '(?mi)^\s*function\s+([A-Za-z0-9_\-]+)')) {
                    $script:PolicyCommonFn[$m.Groups[1].Value.ToLower()] = $true
                }
            }
        } catch { }
    }
    foreach ($k in @($script:PolicyCommonFn.Keys)) { $script:PolicyDefinedFn[$k] = $true }

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
        # PRC001 -- rewritten 2026-08-13. It had a 100% false-positive rate:
        # three findings, all three wrong.
        #
        # It asked "is there a double quote ON THIS LINE". Both launchers build
        # the argument on the PRECEDING line --
        #
        #     $argLine = '-NoExit ... -Command "' + ($parts -join '; ') + '"'
        #     Start-Process -FilePath $psExe -ArgumentList $argLine
        #
        # -- which is correctly quoted and was reported as unquoted every run.
        # The third finding was this rule matching its OWN message string, the
        # same self-match already documented for VER001 below.
        #
        # Now: only flag a LITERAL argument with no quote in it. When the value
        # is a variable, look back up to 10 lines for where it was built and
        # accept a quote there. A check that is wrong every time it fires is
        # worse than no check, because it teaches you to skim past the section
        # it lives in.
        # THE PATTERNS ARE ASSEMBLED, INCLUDING HERE IN THE CONDITION.
        #
        # First attempt assembled only the MESSAGE, and the condition line --
        # which spells out both literals to search for -- promptly matched
        # itself instead. Fixing half of a self-match just moves it.
        #
        # No line in this file may contain both spellings intact. That is the
        # actual rule, and it applies to conditions, messages and comments
        # alike; the comment above says "-Argument" + "List" for the same
        # reason.
        $spPat  = 'Start' + '-Process'
        $argPat = '-Argument' + 'List'
        if ($line -match $spPat -and $line -match $argPat) {
            $argPart = ''
            if ($line -match ($argPat + '\s+(\S+)')) { $argPart = $Matches[1] }
            $quoted = ($line -match '"')
            if (-not $quoted -and $argPart -match '^\$(\w+)') {
                $varName = $Matches[1]
                $back = $code[[Math]::Max(0, $i - 10)..$i] -join "`n"
                if ($back -match ('\$' + [regex]::Escape($varName) + '\s*=[\s\S]*?"')) { $quoted = $true }
            }
            if (-not $quoted) {
                # Assembled, so this line cannot match its own rule -- see the
                # VER001 note below for why that is not paranoia.
                $sp = 'Start-' + 'Process'
                Add-Finding -Severity WARN -File $File -Line $n -Rule 'PRC001' `
                    -Message "$sp passes -ArgumentList through unquoted. A value containing a space arrives as two arguments."
            }
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

        # STRINGS OUT BEFORE COUNTING A CALL -- section 28, applied to the rule
        # that was breaking it.
        #
        # $code has comments stripped and STRING LITERALS KEPT, which is right
        # for finding hardcoded values and wrong for counting code. GroupEditor
        # throws this:
        #
        #     throw "Set-DistributionGroup -ManagedBy replaces the whole owner
        #            list rather than removing one entry."
        #
        # -- a sentence explaining why the tool will not do something. The rule
        # read the cmdlet name inside it as a live call and reported an
        # unguarded handler for two weeks. The handler makes no Exchange call at
        # all.
        #
        # Naming a cmdlet in an error message is good practice and the check
        # must not punish it. Same replace pair LOG002 already uses.
        $bare = $body -replace "'[^']*'", '' -replace '"[^"]*"', ''

        # A NAME THE TOOLKIT DEFINES IS NOT AN EXCHANGE CMDLET, however much
        # it looks like one.
        #
        # The pattern is a verb plus a noun starting Mailbox, Recipient,
        # DistributionGroup and so on -- which is exactly the shape a tool
        # gives its OWN wrapper. MobileDevices has a local
        # Remove-MailboxDevice (three lines around Invoke-EXOCommand) and
        # Unified-Lookup has a local Get-MailboxSettingsPlan, which touches no
        # directory at all: it reads a row and returns a plan object. Both
        # were reported as unguarded Exchange calls on 2026-09-12, and in both
        # the error handling is where this rule's own JUDGEMENT clause says to
        # look -- in the function being delegated to.
        #
        # So the match has to be a name NOBODY here defines. Same principle as
        # PRM001 and FUN003: the toolkit's own functions are known quantities
        # and the rules have to know them. The names are NAMED in the finding
        # now as well, so the next judgement call takes seconds.
        $hdlNames = @([regex]::Matches($bare,
            '\b(?:Get|Set|New|Remove|Add|Enable|Disable)-(?:Mailbox|Recipient|DistributionGroup|UnifiedGroup|RemoteMailbox|InboxRule|MailboxPermission|RecipientPermission|ADUser|ADGroup)\w*\b') |
            ForEach-Object { $_.Value } | Select-Object -Unique)
        $hdlReal = @($hdlNames | Where-Object { -not $script:PolicyDefinedFn[$_.ToLower()] })

        if ($bare -notmatch '\btry\b' -and $hdlReal.Count -gt 0) {
            Add-Finding -Severity WARN -File $File -Line $line -Rule 'HDL001' -Message ("Click handler calls Exchange/AD directly with no try/catch -- $($hdlReal -join ', '). An exception here is caught by .NET, not PowerShell, and the operator sees assembly versions instead of a line number. JUDGEMENT: a handler that delegates to a function which catches its own errors is fine -- check where the call goes before adding a try here.")
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

#region T3c -- form geometry -----------------------------------------------------
function Test-FormGeometry {
    <#
        FINDS BUTTON ROWS THAT ARE OFF THE BOTTOM OF THEIR OWN WINDOW.

        Nearly every form here sets $form.Size and then places its controls in
        CLIENT coordinates. Those are not the same rectangle: Size includes the
        title bar and the border, about 39px vertically and 16px horizontally on
        the default style. So the design height silently loses 39px, and any
        control positioned within those last 39px is drawn outside the client
        area -- present, wired, and invisible.

        THERE IS NO SYMPTOM AT DESIGN TIME AND NO ERROR AT RUN TIME. The window
        simply looks like it was built without a Close button, and it cannot be
        recovered by resizing when the control is anchored to the bottom, since
        the anchor drags it further down as the form shrinks. Three tools
        shipped this way and each one had to be found by eye, in a screenshot,
        one at a time.

        WHAT IT CHECKS. For each form built with a literal .Size, the lowest
        control added DIRECTLY to that form. Past the client edge is an ERROR.
        Level with it -- 0 or 1px -- is a WARN, because the next control added
        crosses the edge.

        THE WARN THRESHOLD WAS 10px AND THAT WAS TOO LOUD. It produced six rows
        about margins of 4 to 9px, every one of them a window that looks fine
        and works fine. A check that reports things nobody is going to act on is
        how the rows that matter get skipped. Errors are unchanged: a control
        drawn off the form is always worth a row.

        WHAT IT CANNOT SEE, and this is worth knowing before trusting a clean
        result: controls placed with a running $y or any other computed value,
        controls inside a GroupBox or Panel, anything positioned at run time,
        and anything using Dock. A form that sizes itself from its content --
        ClientSize = ($x.Right + 12), ($x.Bottom + 12) -- is skipped, correctly:
        it cannot be wrong.

        THE FIX IS NEVER TO PAD THE NUMBER. It is to set ClientSize instead of
        Size, so the value in the source is the value the controls are laid out
        against and the bottom margin can be read off directly.
    #>
    param([Parameter(Mandatory)][string]$File)

    $text = [System.IO.File]::ReadAllText($File)
    $lines = $text -split "`r?`n"

    # Chrome for a normal sizable window. Approximate on purpose -- this is a
    # smell detector, not a layout engine, and the tolerance below absorbs the
    # difference between border styles.
    $chromeH = 39

    $formStarts = @()
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '\$(\w+)\s*=\s*New-Object\s+System\.Windows\.Forms\.Form\b') {
            $formStarts += [PSCustomObject]@{ Var = $Matches[1]; Line = $i }
        }
    }
    if ($formStarts.Count -eq 0) { return }

    for ($k = 0; $k -lt $formStarts.Count; $k++) {
        $fv    = $formStarts[$k].Var
        $start = $formStarts[$k].Line
        # A form's region ends at the NEXT form, so two dialogs that both use
        # $dlg are measured separately. Sharing a variable name across dialogs
        # is common here, and pooling their controls produced a 67px "overflow"
        # that did not exist.
        $end   = if ($k + 1 -lt $formStarts.Count) { $formStarts[$k + 1].Line } else { $lines.Count - 1 }

        $sizeH = 0; $sizeLine = 0
        for ($i = $start; $i -le $end; $i++) {
            if ($lines[$i] -match ('^\s*\$' + [regex]::Escape($fv) + '\.Size\s*=\s*New-Object\s+System\.Drawing\.Size\(\s*(\d+)\s*,\s*(\d+)\s*\)')) {
                $sizeH = [int]$Matches[2]; $sizeLine = $i + 1; break
            }
        }
        if ($sizeH -eq 0) { continue }   # ClientSize, or computed -- nothing to check

        $tops = @{}; $heights = @{}; $added = @{}
        for ($i = $start; $i -le $end; $i++) {
            $l = $lines[$i]
            if ($l -match '^\s*\$(\w+)\.Location\s*=\s*New-Object\s+System\.Drawing\.Point\(\s*(\d+)\s*,\s*(\d+)\s*\)') {
                if (-not $tops.ContainsKey($Matches[1])) { $tops[$Matches[1]] = [int]$Matches[3] }
            }
            if ($l -match '^\s*\$(\w+)\.Size\s*=\s*New-Object\s+System\.Drawing\.Size\(\s*(\d+)\s*,\s*(\d+)\s*\)') {
                if (-not $heights.ContainsKey($Matches[1])) { $heights[$Matches[1]] = [int]$Matches[3] }
            }
            if ($l -match ('\$' + [regex]::Escape($fv) + '\.Controls\.Add\(\$(\w+)\)')) {
                $added[$Matches[1]] = $true
            }
        }

        $clientH = $sizeH - $chromeH
        $worstBottom = 0; $worstName = ''
        foreach ($c in $added.Keys) {
            if (-not $tops.ContainsKey($c)) { continue }
            $h = if ($heights.ContainsKey($c)) { $heights[$c] } else { 28 }
            $b = $tops[$c] + $h
            if ($b -gt $worstBottom) { $worstBottom = $b; $worstName = $c }
        }
        if ($worstBottom -eq 0) { continue }

        $margin = $clientH - $worstBottom
        if ($margin -lt 0) {
            Add-Finding -Severity ERROR -File $File -Line $sizeLine -Rule 'LAY001' `
                -Message ("`$$fv sets .Size height $sizeH, leaving a client area of about $clientH, " +
                          "but `$$worstName reaches $worstBottom -- roughly $([Math]::Abs($margin))px BELOW the bottom of the window. " +
                          "That control is drawn off the form: visible in the source, invisible on screen, and unrecoverable by resizing if it is anchored to the bottom. " +
                          "Set ClientSize instead of Size (about $($worstBottom + 12) here), rather than padding the height.")
        }
        elseif ($margin -lt 2) {
            Add-Finding -Severity WARN -File $File -Line $sizeLine -Rule 'LAY002' `
                -Message ("`$$fv leaves ${margin}px below `$$worstName -- the control is level with the bottom edge of the client area. " +
                          "Size includes the title bar and border while the controls are placed in client coordinates, so the next control added here goes off the form. " +
                          "Set ClientSize (about $($worstBottom + 12) here).")
        }
    }
}
#endregion

#region T3d -- comma inside a method call ---------------------------------------
function Test-MethodArgumentComma {
    <#
        FINDS AN OPERATOR EXPRESSION CONTAINING A COMMA, PASSED TO A METHOD
        WITHOUT ITS OWN PARENTHESES.

            [void]$list.Add($x -replace '\,', ',')      <-- broken
            [void]$list.Add(($x -replace '\,', ','))    <-- correct

        Inside a method call a comma is an ARGUMENT SEPARATOR before it is
        anything else, so the first line passes TWO arguments and dies at run
        time with:

            Cannot find an overload for "Add" and the argument count: "2".

        WHY A CHECK AND NOT JUST CARE. It parses perfectly. Test-Parse is happy,
        the file is well-formed, and nothing says a word until somebody runs the
        code path -- which in Object-Lookup was the first search anybody tried.
        A fault that is invisible until run time and obvious afterwards is
        exactly what a static check is for.

        THE OPERATORS THAT MATTER are the binary ones whose right-hand side
        routinely contains a comma: -replace, -split, -join, and -f. Others take
        a comma so rarely that including them would cost more in false alarms
        than it saves.

        HOW IT AVOIDS FALSE POSITIVES. It only looks at method calls whose
        argument text contains NO nested parentheses -- so a correctly wrapped
        call, where the very next character is '(', never matches. That also
        means a call with any other parenthesised sub-expression is skipped
        rather than guessed at: this reports things it is sure of, and stays
        quiet otherwise.
    #>
    param([Parameter(Mandatory)][string]$File)

    $lines = [System.IO.File]::ReadAllText($File) -split "`r?`n"

    # BLOCK COMMENTS ARE SKIPPED, and that is not fussiness. The first version
    # stripped only '#' line comments, and its very first run reported the
    # WORKED EXAMPLE inside this function's own doc block -- a check whose
    # first finding is itself is a check nobody trusts twice. Doc blocks all
    # over this toolkit contain sample code, and every one of them would have
    # been a false alarm.
    $inBlock = $false

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $l = $lines[$i]

        if ($inBlock) {
            if ($l -match '#>') { $inBlock = $false; $l = ($l -split '#>', 2)[1] } else { continue }
        }
        if ($l -match '<#') {
            $before = ($l -split '<#', 2)[0]
            if ($l -match '#>') { $l = $before + ($l -split '#>', 2)[1] } else { $inBlock = $true; $l = $before }
        }

        $code = $l -replace '(?<!`)#.*$', ''
        if (-not $code.Trim()) { continue }

        foreach ($m in [regex]::Matches($code, '\.\w+\(([^()]*)\)')) {
            $args = $m.Groups[1].Value
            if ($args -notmatch '\s-(replace|split|join|f)\s') { continue }
            if ($args -notmatch ',') { continue }

            Add-Finding -Severity ERROR -File $File -Line ($i + 1) -Rule 'ARG001' `
                -Message ("A method call passes an operator expression containing a comma, without wrapping it: " +
                          "$($m.Value.Trim()). Inside a method call the comma separates ARGUMENTS, so this is passed as two " +
                          "and fails at run time with 'Cannot find an overload ... argument count: 2'. It parses cleanly, so " +
                          "nothing else catches it. Wrap the expression in its own parentheses.")
        }
    }
}
#endregion

#region T3a1 -- [void] piped to Out-Null ----------------------------------------
function Test-VoidPipedToOutNull {
    <#
        DISCARDING THE SAME VALUE TWICE, WHICH CRASHES AT COMPILE TIME.

            [void][System.Windows.Forms.MessageBox]::Show(...) | Out-Null

        The cast yields NOTHING. Piping nothing into Out-Null makes PowerShell
        try to build a pipeline over a System.Void result, and the compiler
        throws before a single statement runs:

            System.ArgumentException: Argument type cannot be System.Void
               at System.Linq.Expressions.Expression.Variable(Type, String)
               at ...Compiler.VisitPipeline(PipelineAst)

        WHY THIS RULE EARNS ITS PLACE. That message names no file, no line and
        no cmdlet -- it surfaces as an unhandled .NET dialog with a stack trace
        full of PowerShell internals. It shipped in ToolStatus's park handler on
        2026-08-21 and the operator's only clue was that a button crashed.

        AND NOTHING ELSE CATCHES IT. The file parses -- braces, parentheses and
        quotes all balance -- so a balance check passes it, and the failure is at
        COMPILE time of the enclosing scriptblock, which means it does not fire
        until the handler is invoked. A tool can be edited, reviewed and shipped
        with this in a button nobody happened to click.

        EITHER FORM ALONE IS CORRECT. [void]X discards; X | Out-Null discards.
        The fix is to delete one of them, and [void] is the house style here.

        MULTI-LINE CALLS ARE JOINED before testing, because every real instance
        spans several lines -- the cast opens the call and the '| Out-Null' sits
        on the closing line, so a line-at-a-time check sees neither.
    #>
    param([Parameter(Mandatory)][string]$File)

    $lines = @(Get-Content -LiteralPath $File -ErrorAction SilentlyContinue)
    $buf = @(); $start = 0; $depth = 0
    $inBlock = $false

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $l = $lines[$i]

        # BLOCK COMMENTS ARE SKIPPED, and this rule's FIRST RUN is why.
        #
        # The docstring above contains the broken form as its worked example, so
        # a line-comment-only strip flagged this very function -- the same way
        # ARG001 flagged its own docstring when it was written. A checker that
        # cannot tell its own explanation from a defect will report one every
        # time somebody documents the thing it looks for.
        if ($inBlock) {
            if ($l -match '#>') { $inBlock = $false; $l = ($l -split '#>', 2)[1] } else { continue }
        }
        if ($l -match '<#') {
            $before = ($l -split '<#', 2)[0]
            if ($l -match '#>') { $l = $before + ($l -split '#>', 2)[1] } else { $inBlock = $true; $l = $before }
        }

        $code = $l -replace '(?<!`)#.*$', ''

        if ($start -eq 0) {
            if ($code -notmatch '\[void\]') { continue }
            $start = $i + 1; $buf = @($code)
            $depth = ([regex]::Matches($code, '\(')).Count - ([regex]::Matches($code, '\)')).Count
        } else {
            $buf += $code
            $depth += ([regex]::Matches($code, '\(')).Count - ([regex]::Matches($code, '\)')).Count
        }

        if ($depth -le 0) {
            $joined = $buf -join ' '
            # The pipe must follow the CLOSING paren of the cast expression --
            # a pipe INSIDE the arguments, as in [void]$sb.AppendLine(($x | Out-String)),
            # is ordinary and correct.
            if ($joined -match '\[void\].*\)\s*\|\s*Out-Null') {
                Add-Finding -Severity ERROR -File $File -Line $start -Rule 'VOID001' `
                    -Message ("A [void] cast has its result piped to Out-Null. The cast already " +
                              "yields nothing, so the pipeline is built over System.Void and the " +
                              "SCRIPTBLOCK FAILS TO COMPILE -- 'Argument type cannot be System.Void', " +
                              "raised as an unhandled .NET dialog with no file or line. It does not " +
                              "fire until the enclosing handler runs, so it survives review. " +
                              "Delete one of the two: keep [void], drop the | Out-Null.")
            }
            $start = 0; $buf = @(); $depth = 0
        }
    }
}
#endregion

#region T3a2 -- the OwnProcess marker written in prose --------------------------
function Test-OwnProcessMarkerInProse {
    <#
        A HEADER COMMENT THAT TALKS ABOUT THE MARKER *IS* THE MARKER.

        Test-ToolDeclaresOwnProcess in AdminConsole.ps1 does this:

            $head = Get-Content -LiteralPath $ScriptPath -TotalCount 40
            [bool]($head -match 'TOOLKIT-REQUIRES:\s*OwnProcess')

        It reads forty lines and matches text. It cannot tell a declaration from
        a sentence about one, so a comment explaining that a tool does NOT need
        its own process declares that it does.

        THIS IS NOT HYPOTHETICAL. Unified-Lookup.ps1 shipped with line 4 reading
        "NO 'TOOLKIT-REQUIRES: OwnProcess', AND THAT IS THE WHOLE TRICK", was
        launched detached by the console because of it, opened without the
        Exchange Online session that lives in the console's process, and
        reported "ON-PREMISES ONLY". The connection model was rewritten twice
        before anyone read line 4.

        THE RULE: within the first 40 lines, the marker may appear ONCE and only
        as a declaration -- at the very start of a comment line, optionally
        followed by an explanation. Anything else (quoted, mid-sentence,
        preceded by NO/NOT/REMOVED) is prose, and prose here is a live switch.

        To write about it in a header, spell the colon out -- TOOLKIT-REQUIRES
        <colon> OwnProcess -- or just call it "the OwnProcess marker".
    #>
    param([Parameter(Mandatory)][string]$File)

    $lines = @(Get-Content -LiteralPath $File -TotalCount 40 -ErrorAction SilentlyContinue)
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $l = $lines[$i]
        if ($l -notmatch 'TOOLKIT-REQUIRES:\s*OwnProcess') { continue }

        # A real declaration: comment marker, optional space, then the token.
        # ^\s*#\s*TOOLKIT-REQUIRES:\s*OwnProcess -- nothing in front of it.
        if ($l -match '^\s*#\s*TOOLKIT-REQUIRES:\s*OwnProcess') { continue }

        Add-Finding -Severity ERROR -File $File -Line ($i + 1) -Rule 'OWN001' `
            -Message ("The OwnProcess marker appears in the first 40 lines but NOT as a declaration: " +
                      "$($l.Trim()). Test-ToolDeclaresOwnProcess matches this text and will launch the tool " +
                      "in a detached process -- with no Exchange Online session -- whatever the sentence " +
                      "around it says. Rewrite it as 'the OwnProcess marker', or spell the colon out.")
    }
}
#endregion

#region T3a3 -- audit Target is an address, not a name --------------------------
function Test-LogTargetIsName {
    <#
        AN AUDIT ROW THAT NAMES A PERSON INSTEAD OF AN OBJECT.

        The log had rows like

            Target : Armbruster, Glenn W.
            Detail : ForwardingSmtpAddress: 'x@y.com' -> '(cleared)'

        which record that SOMETHING was changed and lose which. One display name
        can cover three mailboxes; an address covers exactly one, is what every
        cmdlet takes as -Identity, and is what the next person greps for.

        Display names are also not stable -- renaming somebody rewrites every
        historical row's Target into a string nothing answers to.

        WHAT THIS FLAGS: -Target given something whose name says 'name'
        (.Name, $name, .DisplayName, .SamAccountName) rather than an address.

        WHAT IT DELIBERATELY DOES NOT FLAG: a literal string, because those are
        the non-object targets -- 'Exchange Online', 'Backup', 'Dashboard' --
        where there is no address to prefer and the literal IS the identifier.
        Nor Get-LogTarget, which exists precisely to express the rule.

        THE FIX IS ALWAYS THE SAME SHAPE:
            -Target (Get-LogTarget $x.PrimarySmtpAddress $x.Name)
    #>
    param([Parameter(Mandatory)][string]$File)

    $lines = @(Get-Content -LiteralPath $File -ErrorAction SilentlyContinue)
    $inBlock = $false

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $l = $lines[$i]

        if ($inBlock) {
            if ($l -match '#>') { $inBlock = $false; $l = ($l -split '#>', 2)[1] } else { continue }
        }
        if ($l -match '<#') {
            $before = ($l -split '<#', 2)[0]
            if ($l -match '#>') { $l = $before + ($l -split '#>', 2)[1] } else { $inBlock = $true; $l = $before }
        }
        $code = $l -replace '(?<!`)#.*$', ''
        if (-not $code.Trim()) { continue }

        if ($code -notmatch '-Target\s') { continue }
        if ($code -match 'Get-LogTarget')  { continue }   # already correct
        if ($code -match "-Target\s+'")    { continue }   # literal, see docstring

        $m = [regex]::Match($code, '-Target\s+(?<v>"[^"]*"|\$[^\s,)]+)')
        if (-not $m.Success) { continue }
        $v = $m.Groups['v'].Value

        if ($v -match '(?i)(\.Name\b|\.DisplayName\b|\.SamAccountName\b|\$name\b|\$displayName\b)') {
            Add-Finding -Severity WARN -File $File -Line ($i + 1) -Rule 'TGT001' `
                -Message ("Audit -Target is given a NAME rather than an address: $v. One display name can " +
                          "cover several mailboxes, so the row records that something changed and loses which -- " +
                          "and a rename rewrites every historical row. Use " +
                          "-Target (Get-LogTarget `$x.PrimarySmtpAddress `$x.Name). If the object genuinely has " +
                          "no address, Get-LogTarget still falls back to the name and states the intent.")
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

        if ($bare -match ('\b(Set|New|Remove|Add|Disable|Enable|Update|Move|Restore)-' +
                          '(Mailbox|ADUser|ADGroup|ADObject|ADGroupMember|ADPrincipalGroupMembership|' +
                          'DistributionGroup|DistributionGroupMember|UnifiedGroup|UnifiedGroupLinks|' +
                          'MailboxPermission|MailboxFolderPermission|RecipientPermission|ADPermission|' +
                          'InboxRule|MailboxAutoReplyConfiguration|CASMailbox|CalendarProcessing|' +
                          'MailContact|MailUser|MsolUser|MgUser|SPOSite|SPOUser|Place)\b')) {
            # -WhatIf recorded per call, not per file. One preview call in a
            # tool that also writes for real must not exempt the whole file.
            [void]$writeLines.Add(@{
                Line     = $i + 1
                Cmdlet   = $Matches[0]
                HasWhatIf = ($bare -match '-WhatIf\b')
            })
        }
    }

    # CATCHES THAT SWALLOW, counted separately from catches that rethrow.
    #
    # Done on the whole text rather than line by line, because deciding whether
    # a catch only rethrows means reading its BODY, and a body spans lines. The
    # braces are matched by hand -- a regex cannot balance them.
    #
    # "Only rethrows" means the body contains nothing but throw / Write-Error
    # -ErrorAction Stop and whitespace. Anything else -- a message to the user,
    # a status label, a return -- means the error stops there, and a failure
    # that stops somewhere with no log row is the thing LOG002 exists to find.
    $swallowCount = 0
    $flatForCatch = $code -join "`n"
    foreach ($m in [regex]::Matches($flatForCatch, '(?m)\bcatch\b[^{]*\{')) {
        $d = 0; $k = $m.Index + $m.Length - 1
        while ($k -lt $flatForCatch.Length) {
            if     ($flatForCatch[$k] -eq '{') { $d++ }
            elseif ($flatForCatch[$k] -eq '}') { $d--; if ($d -eq 0) { break } }
            $k++
        }
        $body = $flatForCatch.Substring($m.Index + $m.Length, [Math]::Max(0, $k - $m.Index - $m.Length))
        $stripped = ($body -replace '(?m)^\s*throw\b[^\n]*', '' -replace '\s', '')
        if ($stripped -ne '') { $swallowCount++ }
    }

    if ($writeLines.Count -eq 0) { return }

    $flat      = $code -join "`n"
    $logsAny   = $flat -match '\bWrite-CompletedLog\b'
    $logsFail  = $flat -match '\bWrite-FailedLog\b'

    # PREVIEW-ONLY TOOLS ARE EXEMPT, AND THE RULE WORKS IT OUT ITSELF.
    #
    # This was a footnote telling the reader to use judgement. A tool whose
    # every write carries -WhatIf changes nothing, so it has nothing to record
    # and LOG001 firing on it was simply wrong. Encoded here, the caveat is not
    # needed: the rule stops being wrong instead of being explained.
    #
    # TOOLKIT-PREVIEW-ONLY in a header is the explicit form, for a tool that
    # previews some other way -- building a plan it never applies, say.
    $previewOnly = ($flat -match '(?m)^\s*#\s*TOOLKIT-PREVIEW-ONLY\b')
    if (-not $previewOnly -and $writeLines.Count -gt 0) {
        $whatIfed = @($writeLines | Where-Object { $_.HasWhatIf }).Count
        $previewOnly = ($whatIfed -eq $writeLines.Count)
    }

    if (-not $logsAny) {
        if ($previewOnly) { return }
        $first = $writeLines[0]
        Add-Finding -Severity WARN -File $File -Line $first.Line -Rule 'LOG001' `
            -Message ("Calls $($writeLines.Count) write cmdlet(s) starting with $($first.Cmdlet) " +
                      "and never calls Write-CompletedLog. Changes production without leaving a record.")
        return
    }

    # $swallowCount, not $catchCount. A catch that only rethrows has not
    # handled anything -- the error still travels up to whoever called, and
    # writing a log row there would double-count a failure the caller is about
    # to record. That was the other footnote; counting only the catches that
    # actually SWALLOW makes the number honest and the note unnecessary.
    if (-not $logsFail -and $swallowCount -gt 0) {
        Add-Finding -Severity INFO -File $File -Line 1 -Rule 'LOG002' `
            -Message ("Logs successes but never failures -- $swallowCount catch block(s) that " +
                      "handle the error rather than rethrow it, and no Write-FailedLog. " +
                      "A failed change is the row you will actually go looking for.")
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

    # FIRST FINDING, ALWAYS. Added before any test runs so it sorts to the top
    # of the CSV and cannot be lost among a hundred rows -- a scan that skips
    # part of the toolkit has to say so where it will be read.
    if ($script:ExcludedFileCount -gt 0 -and @($script:ExcludedHit).Count -gt 0) {
        Add-Finding -Severity INFO -File $Folder -Line 1 -Rule 'SKP001' `
            -Message ("Not checked: $($script:ExcludedHit -join ', ') -- " +
                      "$($script:ExcludedFileCount) file(s). Work still to be brought onto this " +
                      "architecture; revisit when it is.")
    }

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
        # Its own try for the same reason as the one above: a regex walk over
        # arbitrary layout code is the kind of thing that throws on one odd
        # file, and one odd file must not cost the whole folder's findings.
        try { Test-FormGeometry -File $f.FullName } catch {
            Add-Finding -Severity WARN -File $f.FullName -Line 1 -Rule 'INT003' -Message "Form-geometry check could not run: $($_.Exception.Message)"
        }
        try { Test-VoidPipedToOutNull -File $f.FullName } catch {
            Add-Finding -Severity WARN -File $f.FullName -Line 0 -Rule 'VOID001' `
                -Message "Could not check for [void] piped to Out-Null: $($_.Exception.Message)"
        }
        try { Test-LogTargetIsName -File $f.FullName } catch {
            Add-Finding -Severity WARN -File $f.FullName -Line 0 -Rule 'TGT001' `
                -Message "Could not check audit -Target values: $($_.Exception.Message)"
        }
        try { Test-OwnProcessMarkerInProse -File $f.FullName } catch {
            Add-Finding -Severity WARN -File $f.FullName -Line 0 -Rule 'OWN001' `
                -Message "Could not check the OwnProcess marker: $($_.Exception.Message)"
        }
        try { Test-MethodArgumentComma -File $f.FullName } catch {
            Add-Finding -Severity WARN -File $f.FullName -Line 1 -Rule 'INT004' -Message "Method-argument check could not run: $($_.Exception.Message)"
        }
    }

    if ($Progress) { & $Progress 'Contract checks...' 98 }
    Test-Contract -Folder $Folder
    Test-TenantLeakage -Folder $Folder -Files $files
    Test-CommandResolution -Folder $Folder -Files $files
    Test-CommonContract -Folder $Folder -Files @($files | ForEach-Object { $_.FullName })
    try { Test-CommonParameterName -Folder $Folder -Files @($files | ForEach-Object { $_.FullName }) } catch {
        Add-Finding -Severity WARN -File $Folder -Line 0 -Rule 'PRM000' -Message "Parameter-name check could not run: $($_.Exception.Message)"
    }

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

    # NO LEGEND BLOCK. Removed entirely on 2026-08-13.
    #
    # It listed five judgement rules on every run, clean ones included, and was
    # the same fault as the tenant-name dump: text you learn to skip, which
    # then hides the one line that matters on the day it applies.
    #
    # All five are accounted for rather than merely hidden:
    #
    #   ENC005  rule deleted -- a style check nobody was going to act on
    #   LOG001  exempts preview-only tools, so the caveat is not needed
    #   LOG002  counts only catches that swallow, same
    #   EXO002  its own message already says what to do about it
    #   HDL001  caveat moved INTO the finding, where it is read at the moment
    #           it applies rather than as a footnote after a clean run
    #
    # A caveat belongs on the finding, not in a glossary. If a rule needs a
    # standing explanation to be usable, the rule is the thing to fix.
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

# TWO VIEWS OF ONE RUN.
#
# Issues   one row per finding: which file, what is wrong in plain words, and
#          how much it matters. Sorted worst first, because the top of a list
#          is the only part reliably read.
# Details  the full technical report, unchanged -- rule IDs, line numbers,
#          exact messages, and the version/integrity header.
#
# The grid is not a replacement for the text. It is the triage layer: the text
# answers "what exactly", the grid answers "do I care". Shipping only the text
# is what made a 12-row clean run look like a wall.
$tabs           = New-Object System.Windows.Forms.TabControl
$tabs.Location  = New-Object System.Drawing.Point(12, 148)
$tabs.Size      = New-Object System.Drawing.Size(960, 486)
$tabs.Anchor    = 'Top,Left,Right,Bottom'
$form.Controls.Add($tabs)

$tabIssues      = New-Object System.Windows.Forms.TabPage
$tabIssues.Text = 'Issues'
$tabs.TabPages.Add($tabIssues)

$tabDetails      = New-Object System.Windows.Forms.TabPage
$tabDetails.Text = 'Details'
$tabs.TabPages.Add($tabDetails)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock                  = 'Fill'
$grid.ReadOnly              = $true
$grid.AllowUserToAddRows    = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.RowHeadersVisible     = $false
$grid.SelectionMode         = 'FullRowSelect'
$grid.MultiSelect           = $false
$grid.AutoSizeColumnsMode   = 'Fill'
$grid.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
# Wrap, because a plain-language sentence is longer than a rule ID and
# truncating it defeats the entire point of writing one.
$grid.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True
$grid.AutoSizeRowsMode = [System.Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
[void]$grid.Columns.Add('Impact', 'Impact')
[void]$grid.Columns.Add('Script', 'Script')
[void]$grid.Columns.Add('Issue',  'What is wrong')
$grid.Columns['Impact'].FillWeight = 10
$grid.Columns['Script'].FillWeight = 22
$grid.Columns['Issue'].FillWeight  = 68
$tabIssues.Controls.Add($grid)

# EMPTY STATE. An empty grid is ambiguous -- it looks identical whether the
# tests found nothing or never ran. Saying so removes the doubt, and points at
# the tab that carries the evidence.
#
# Shown by TOGGLING VISIBILITY rather than layering, because in WinForms the
# control added EARLIER draws on top: a label added after the grid would sit
# behind it, and one added before would cover a grid full of rows. Only one of
# the two is ever visible, so z-order stops mattering.
$lblNoIssues           = New-Object System.Windows.Forms.Label
$lblNoIssues.Text      = 'No issues to display.  See details tab for more information'
$lblNoIssues.Dock      = 'Fill'
$lblNoIssues.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$lblNoIssues.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 10)
$lblNoIssues.Visible   = $false
$tabIssues.Controls.Add($lblNoIssues)

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Dock       = 'Fill'
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
# WordWrap ON. It was off with a horizontal scrollbar, and a finding whose
# explanation sits past the right edge is a finding nobody reads -- the same
# mistake that hid the write-scope message for three rounds on 2026-08-06.
$txtOut.WordWrap   = $true
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$tabDetails.Controls.Add($txtOut)

function Update-IssueGrid {
    <#
        Fill the Issues tab from the findings.

        Colour is on the Impact cell only, not the whole row. A red row reads as
        "everything here is on fire"; a red word in a column headed Impact reads
        as one fact about one finding, which is what it is.
    #>
    param($Findings)

    # One helper, called on every exit path below, so a future early return
    # cannot leave the two controls both visible or both hidden.
    $setEmpty = {
        param([bool]$IsEmpty)
        $lblNoIssues.Visible = $IsEmpty
        $grid.Visible        = -not $IsEmpty
        if ($IsEmpty) { $lblNoIssues.BringToFront() }
    }

    $grid.Rows.Clear()
    if (-not $Findings -or @($Findings).Count -eq 0) { & $setEmpty $true; return }

    # NOTES DO NOT BELONG ON A LIST HEADED "WHAT IS WRONG".
    #
    # SKP001 says Projects\ was skipped on purpose. TEN000 says which names the
    # leakage check looked for. Both are the run TELLING you what it did, not
    # reporting a fault -- and putting them here made the console appear in a
    # list of problems for something that is not its problem.
    #
    # They are still in Details, in full, which is where a record of what ran
    # belongs. This tab answers "what do I need to fix".
    $Findings = @($Findings | Where-Object { (Get-RuleImpact -Rule $_.Rule) -ne 'Note' })
    if (@($Findings).Count -eq 0) { & $setEmpty $true; return }
    & $setEmpty $false

    $rows = @($Findings | ForEach-Object {
        [PSCustomObject]@{
            Impact = (Get-RuleImpact -Rule $_.Rule)
            Script = "$($_.File)"
            Issue  = (Get-RulePlain -Rule $_.Rule -Fallback "$($_.Message)")
            Rank   = 9
        }
    })
    foreach ($r in $rows) {
        if ($script:ImpactRank.ContainsKey($r.Impact)) { $r.Rank = $script:ImpactRank[$r.Impact] }
    }

    foreach ($r in ($rows | Sort-Object Rank, Script)) {
        $i = $grid.Rows.Add()
        $grid.Rows[$i].Cells['Impact'].Value = $r.Impact
        $grid.Rows[$i].Cells['Script'].Value = $r.Script
        $grid.Rows[$i].Cells['Issue'].Value  = $r.Issue

        $c = switch ($r.Impact) {
            'Critical' { [System.Drawing.Color]::FromArgb(200,  40,  40) }
            'High'     { [System.Drawing.Color]::FromArgb(205,  95,   0) }
            'Medium'   { [System.Drawing.Color]::FromArgb(170, 140,   0) }
            'Low'      { [System.Drawing.Color]::FromArgb(100, 120, 160) }
            'Cosmetic' { [System.Drawing.Color]::FromArgb(130, 130, 130) }
            default    { [System.Drawing.Color]::FromArgb(110, 110, 110) }
        }
        $cell = $grid.Rows[$i].Cells['Impact']
        $cell.Style.ForeColor = $c
        $cell.Style.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
    }
}

$grid.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $tabs.SelectedTab = $tabDetails
})

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
        Update-IssueGrid -Findings $findings

        $errCount = @($findings | Where-Object { $_.Severity -eq 'ERROR' }).Count
        $lblStatus.Text = if ($errCount -gt 0) { "$errCount error(s) -- do not ship." } else { 'No errors.' }

        # ALWAYS THE ISSUES TAB. Same landing place every run, clean or not.
        #
        # This used to fall back to Details on a clean run, on the reasoning
        # that an empty grid looks like a scan that never happened. That was
        # true when it was written and stopped being true the moment the empty
        # state got its own line -- "No issues to display" says the run
        # happened and found nothing, which is the whole answer.
        #
        # Consistency is worth more than the cleverness. A tool that lands
        # somewhere different depending on the result makes you read the screen
        # before you can read the answer, and "clean" should be the calmest
        # thing the window can show -- one short sentence, not a wall of text
        # you have to scan to confirm there is nothing in it.
        $tabs.SelectedTab = $tabIssues

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
