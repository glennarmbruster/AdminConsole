# DEVIN REFERENCE CODE — eleven scripts and four config files

**THIS FILE IS SIXTEEN SEPARATE FILES CONCATENATED. DO NOT TREAT IT AS ONE.**

Each section below is a complete, standalone file from a working production
toolkit. Reproduce them as separate files under their own names -- the name is
in every heading. Concatenating them would produce something that cannot run.

They are combined here only because of a transfer constraint.

---

## THESE ARE REFERENCE IMPLEMENTATIONS, NOT A STARTING CODEBASE

The toolkit they come from served a different tenant with a different shape:
**two** on-premises AD domains where the target has **one**. Copying a tool
wholesale imports assumptions that stay invisible until they are wrong.

- **Match the conventions exactly** -- naming, parameter shapes, how a tool
  dot-sources `Common.ps1`, how it logs, how it lays out a form.
- **Do not copy a tool and adjust it.** Write each new tool from its spec in the
  build pack, consulting these for *how* rather than *what*.
- **Do not port multi-directory machinery.** Keep the config-driven shape;
  drop cross-directory pinning, per-directory controller selection, and
  side-routing. Dead branches on a single-domain tenant.
- **Do not port anything you cannot explain.** Leave it out and note the
  omission. Unexplained code that is carried forward becomes permanent.
- **Do not assume this code is bug-free.** It is working production code, not
  scripture. Several constraints in the build pack exist because a previous
  version of these very files had the bug.

**Read the comments.** They are dense and load-bearing. Where a comment explains
why something is done a surprising way, that explanation usually cost somebody a
day. Treat a comment that contradicts your instinct as information, not noise.

---

## Contents

| # | File | Lines | Reference for |
|---|---|---|---|
| 1 | `Common.ps1` | 7198 | THE SPINE |
| 2 | `Test-Toolkit.ps1` | 1950 | THE VERIFIER, already built |
| 3 | `AdminConsole.ps1` | 2232 | The launcher: connection state in the header, the PIM countdown, capability-driven greying, JSON-driven section/button layout, and the modal-loop deferral (constraint B2) |
| 4 | `ToolManager.ps1` | 948 | Config round-trip |
| 5 | `Connect-ExO.ps1` | 251 | Cloud connection |
| 6 | `Connect-ExSE.ps1` | 203 | On-premises connection |
| 7 | `MailboxHealth.ps1` | 373 | THE CANONICAL SMALL TOOL and the first one to build |
| 8 | `MessageTrace.ps1` | 688 | A clean mid-size tool: grid, filter-without-re-running, export, and an explicit statement of what the tool is NOT for |
| 9 | `Unified-Lookup.ps1` | 2756 | The most complex read tool |
| 10 | `Group-Manager-Editor.ps1` | 1246 | THE WRITE REFERENCE |
| 11 | `Protocols.ps1` | 404 | The tool that taught C15 (slice a large query; never retry an oversized one) and C13 (ask the cmdlet what it declares) |
| 12 | `Config/Tenant.TEMPLATE.json` | 133 | The ONLY file permitted to name the organisation |
| 13 | `Config/Settings.json` | 12 | Engineer preferences -- paths, browser profiles |
| 14 | `Config/Sections.json` | 21 | Console layout |
| 15 | `Config/Tools.json` | 373 | Button registry |
| 16 | `Config/Themes.TEMPLATE.json` | 579 | Palettes |

---
---


<br>

==============================================================================

# ◆ FILE 1 OF 16 — `Common.ps1`

> **7198 lines.** THE SPINE. Config readers, connection probes, capability gating, the module prerequisite subsystem, audit log, error display, theming, geometry, and the PIM helpers the console header depends on. Take functions from here verbatim -- but only the ones listed in the COMMON CONTRACT document, which is a FLOOR and not a complete list. 164 functions; porting all of them would consume the build.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```powershell
#Requires -Version 5.1
<#
    Common.ps1
    Shared helper library for the Admin Console launcher and every
    standalone tool script (MailboxHealth.ps1, MessageTrace.ps1,
    InboxRules.ps1, etc.).

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>

#region Tenant configuration ----------------------------------------------------
<#
    EVERY tenant-specific fact enters the toolkit here and nowhere else.

    What used to sit at the top of this file:

        $script:MailDomain            = 'a real domain'
        $script:SecondaryDCs          = @( eight real server names )
        $script:SecondaryDomain       = 'another real domain'
        $script:ExchangeOnPremServer  = 'a real Exchange server'

    That worked for exactly one organisation, and every other deployment became
    a copy -- which is how six copies drifted apart. The values now come from
    Config\Tenant.json, so one toolkit serves every tenant and the only thing
    that differs is a file nobody has to write code to change.

    TENANT.JSON vs SETTINGS.JSON. Tenant.json holds facts about the TENANT:
    domains, directories, servers, which capabilities exist. It is the same on
    every engineer's machine and ships with the toolkit. Settings.json holds
    what an ENGINEER prefers: report paths, browser profiles. It differs per
    machine and is deliberately excluded from the version manifest. Mixing the
    two is what made the old Config\ folder unhashable.

    CAPABILITIES ARE A CLAIM, NOT A PROBE. Capabilities.OnPremExchange says
    "this tenant HAS on-premises Exchange". Test-EmsAvailable says "I can reach
    it right now". Both are useful and they answer different questions -- and
    when they disagree that is worth saying out loud, because "you told me this
    tenant has on-premises Exchange and I cannot reach it" is a better message
    than either alone.
#>

$script:TenantConfigFileName = 'Tenant.json'

function Get-TenantConfigPath {
    return (Join-Path (Join-Path $PSScriptRoot $script:ConfigFolderName) $script:TenantConfigFileName)
}

function Get-TenantConfig {
    <#
        The tenant file, cached for the life of the process.

        NEVER THROWS. A toolkit that will not start because a config file is
        missing is worse than one that starts and says every capability is off:
        the first cannot tell you what is wrong, the second can.
    #>
    param([switch]$Refresh)

    if ($global:TenantConfig -and -not $Refresh) { return $global:TenantConfig }

    $empty = [PSCustomObject]@{
        TenantName        = ''
        PrimaryDomain     = ''
        OnMicrosoftDomain = ''
        Capabilities      = [PSCustomObject]@{
            ExchangeOnline = $false; SharePointOnline = $false
            OnPremAD       = $false; OnPremExchange   = $false
            Pim            = $false; PimTimer         = $false
        }
        Directories    = @()
        OnPremExchange = [PSCustomObject]@{ ServerUri = ''; PreferredDomainController = '' }
        Pim            = [PSCustomObject]@{ DurationHours = 4 }
        LicenseGroups  = @()
        FeatureRequest = [PSCustomObject]@{ Recipient = '' }
        SharePoint     = [PSCustomObject]@{ AdminUrl = ''; MySiteHost = '' }
        Theme          = ''
        Loaded         = $false
        Problem        = ''
    }

    $p = Get-TenantConfigPath
    if (-not (Test-Path -LiteralPath $p)) {
        $empty.Problem = "No $($script:TenantConfigFileName) in Config. Every capability reads as off."
        $global:TenantConfig = $empty
        return $global:TenantConfig
    }
    try {
        $j = Get-Content -LiteralPath $p -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        Add-Member -InputObject $j -NotePropertyName 'Loaded'  -NotePropertyValue $true  -Force
        Add-Member -InputObject $j -NotePropertyName 'Problem' -NotePropertyValue ''     -Force
        $global:TenantConfig = $j
    } catch {
        $empty.Problem = "$($script:TenantConfigFileName) could not be parsed: $($_.Exception.Message)"
        $global:TenantConfig = $empty
    }
    return $global:TenantConfig
}

function Test-TenantCapability {
    <#
        Does this tenant HAVE the capability? Not "is it connected".

        Unknown names return $false rather than throwing, so a toolkit newer
        than its config degrades to "off" instead of failing to launch.
    #>
    param([Parameter(Mandatory)]
          [ValidateSet('ExchangeOnline','SharePointOnline','OnPremAD','OnPremExchange','Pim','PimTimer')]
          [string]$Name)
    $c = (Get-TenantConfig).Capabilities
    if (-not $c) { return $false }
    if (-not $c.PSObject.Properties[$Name]) { return $false }
    return [bool]$c.$Name
}

function Get-TenantDirectory {
    <#
        The configured AD directories, or one of them by Id.

        A LIST, not a hardcoded pair. The old model had exactly two, named for
        one organisation's internal structure, and every function that touched
        AD carried a ValidateSet naming them. A tenant with one domain now has
        one entry; a tenant with five has five; a cloud-only tenant has none
        and every AD tool greys out on its own.
    #>
    param([string]$Id)
    $dirs = @((Get-TenantConfig).Directories)
    if (-not $Id) { return $dirs }
    return @($dirs | Where-Object { "$($_.Id)" -eq $Id }) | Select-Object -First 1
}

function Get-TenantDirectoryOu {
    <#
        Where new objects of a given kind go in a given directory.

        These used to be destined for a script: 'corp.example.com/Accounts/
        Resource Mailbox' branched on an "if <directory name>". That is one
        organisation's AD layout written into a tool, and the same mistake as
        the ValidateSet on directory names -- Test-Toolkit flags a domain in a
        .ps1 for exactly this reason.

        Config-driven instead, and the behaviour falls out rather than being
        branched. A directory that HAS the OU configured uses it; one that does
        not gets asked. A directory prompts because nobody filled it in, not
        because of what it is called -- so a tenant with five directories, or
        one, behaves sensibly without a code change.

        Returns '' when unset. The caller must prompt rather than invent a path.
    #>
    param(
        [string]$DirectoryId,
        [Parameter(Mandatory)][ValidateSet('SharedMailbox', 'DistributionGroup')]
        [string]$Kind
    )

    $dir = $null
    if ($DirectoryId) { $dir = Get-TenantDirectory -Id $DirectoryId }
    if (-not $dir)    { $dir = Get-DefaultTenantDirectory }
    if (-not $dir)    { return '' }

    $key = $(if ($Kind -eq 'SharedMailbox') { 'SharedMailboxOu' } else { 'DistributionGroupOu' })
    if (-not $dir.PSObject.Properties[$key]) { return '' }
    return "$($dir.$key)".Trim()
}

function Get-DefaultTenantDirectory {
    $dirs = @(Get-TenantDirectory)
    if ($dirs.Count -eq 0) { return $null }
    $d = @($dirs | Where-Object { $_.IsDefault }) | Select-Object -First 1
    if ($d) { return $d }
    return $dirs[0]
}

function Get-TenantDirectoryServer {
    <#
        A server to bind to for a given directory: the preferred DC if one is
        configured, otherwise the domain name, which lets AD locate a DC itself.

        Returning the DOMAIN rather than a specific server matters -- pinning a
        named controller is what broke cross-domain writes (ARCHITECTURE 14f).
        A configured list is for tenants that must pin; the default is not to.
    #>
    # -DirectoryId is an ALIAS, and it is here because the name is wrong.
    #
    # Every other function that takes a directory calls it -DirectoryId:
    # Get-TenantDomainController, Get-TenantDirectoryOu,
    # Resolve-TenantLicenseGroupName. This one says -Id, and a caller that
    # reaches for the familiar name binds NOTHING -- no error worth noticing,
    # the parameter is simply unset, and every downstream -Server $null call
    # silently falls back to the current domain. NotDeadYet searched the wrong
    # forest partition for exactly that reason.
    #
    # Renaming outright would break existing callers, so both spellings work.
    param([Alias('DirectoryId')][string]$Id)
    $d = if ($Id) { Get-TenantDirectory -Id $Id } else { Get-DefaultTenantDirectory }
    if (-not $d) { return $null }
    $dcs = @($d.DomainControllers)
    if ($dcs.Count -eq 1) { return "$($dcs[0])" }
    return "$($d.Domain)"
}

function ConvertTo-DomainDn {
    <#
        'contoso.com' -> 'DC=contoso,DC=com'.

        Used to decide which configured directory an object BELONGS to, from
        its distinguished name alone. No directory call, so it costs nothing
        and works on a DN read out of a CSV.
    #>
    param([string]$Domain)
    $d = "$Domain".Trim().Trim('.')
    if (-not $d) { return '' }
    return (($d -split '\.' | Where-Object { $_ } | ForEach-Object { "DC=$_" }) -join ',')
}

function Get-DirectoryDomainDn {
    <# The domain naming context of a configured directory, or '' if unknown. #>
    param([string]$DirectoryId)
    $d = if ($DirectoryId) { Get-TenantDirectory -Id $DirectoryId } else { Get-DefaultTenantDirectory }
    if (-not $d) { return '' }
    return (ConvertTo-DomainDn -Domain "$($d.Domain)")
}

function Resolve-DistinguishedNameDirectory {
    <#
        Which configured directory does this DN live in? Returns the directory
        object, or $null when it matches none of them.

        LONGEST SUFFIX WINS, AND THAT IS THE WHOLE POINT OF THE FUNCTION.

        A child domain's DN ends with its parent's. Given a parent domain and a
        child domain beneath it, EVERY child object also matches the parent
        suffix -- so a naive "does the DN end with the parent's DN" test says
        yes for both, and a report scoped to the parent silently swallows the
        entire child directory.

        That is not a hypothetical: the phase 1 group inventory is scoped to
        one directory precisely so the child's groups stay out of it, and a
        first-match test would have put them all in the file. Sorting the
        candidate suffixes longest-first makes the specific answer beat the
        general one.

        Comparison is case-insensitive because AD is, and both DNs are
        normalised to have no stray spaces around the commas -- a DN copied out
        of a console or a CSV frequently has them, and a string compare that
        trips over whitespace fails in a way nobody thinks to look for.
    #>
    param([string]$DistinguishedName)

    $dn = (("$DistinguishedName" -split ',') | ForEach-Object { $_.Trim() }) -join ','
    if (-not $dn) { return $null }

    $cands = @()
    foreach ($d in @(Get-TenantDirectory)) {
        $ddn = ConvertTo-DomainDn -Domain "$($d.Domain)"
        if ($ddn) { $cands += [PSCustomObject]@{ Directory = $d; Dn = $ddn; Len = $ddn.Length } }
    }
    if ($cands.Count -eq 0) { return $null }

    foreach ($c in @($cands | Sort-Object -Property Len -Descending)) {
        if ($dn.EndsWith($c.Dn, [StringComparison]::OrdinalIgnoreCase)) { return $c.Directory }
    }
    return $null
}

function Test-DnInDirectory {
    <#
        Does this DN belong to THIS directory and not to a child of it?

        The boolean form of Resolve-DistinguishedNameDirectory, for filtering a
        bulk result set down to one directory. Same longest-suffix rule, so a
        child domain's objects answer false.
    #>
    param(
        [string]$DistinguishedName,
        [Parameter(Mandatory)][string]$DirectoryId
    )
    $d = Resolve-DistinguishedNameDirectory -DistinguishedName $DistinguishedName
    if (-not $d) { return $false }
    return ("$($d.Id)" -ieq $DirectoryId)
}

function Get-PinnableDomainController {
    <#
        A value safe to hand to Set-ADServerSettings -PreferredServer, or ''
        meaning "do not pin -- go forest-wide and let AD locate one".

        THE SEAM THIS EXISTS TO COVER. Three parameters look interchangeable
        and are not:

          AD module -Server            accepts a DOMAIN or a server. AD locates
                                       a controller. This is the preferred form
                                       for Get-ADUser and friends
                                       (ARCHITECTURE 14f).
          Exchange -DomainController   accepts a SERVER ONLY.
          -PreferredServer             accepts a SERVER ONLY.

        CORRECTED 2026-08-12. This block used to group -Server and
        -DomainController together as "accept a DOMAIN", which is true of the
        AD module and false of Exchange. Exchange's -DomainController runs the
        SAME suitability check as -PreferredServer and rejects a domain with
        "Active Directory error 0x52 ... A local error occurred". NewUser was
        changed to trust that sentence and every enable failed on it.

        So the rule is by MODULE, not by parameter name: an AD cmdlet may be
        handed Get-TenantDirectoryServer's answer (which is a domain when a
        directory names several controllers); an Exchange cmdlet must be handed
        a controller from Get-TenantDomainController, or nothing at all.

        Get-TenantDirectoryServer deliberately returns the DOMAIN when a
        directory names several controllers, which is right for the first two
        and fatal for the third. NewUser crossed that seam and every mailbox
        enable in a multi-controller directory failed on it.

        Validation, not selection. A candidate is returned only when
        Tenant.json actually names it as a controller. Anything else -- a
        domain, a typo, a plausible-looking hostname -- comes back as '', and
        the caller goes forest-wide.

        Refusing to guess is the whole point. Bulk-DLMembership records the
        opposite failure on 2026-08-06: pinning to one directory's controller
        made a write in ANOTHER directory fail as "out of the current user's
        write scope", which
        reads like a permissions problem and is not. A wrong pin is worse than
        no pin.
    #>
    param(
        [string]$Candidate,
        [string]$DirectoryId
    )

    $cand = "$Candidate".Trim()
    if (-not $cand) { return '' }

    $known = @(if ($DirectoryId) {
        Get-TenantDomainController -DirectoryId $DirectoryId
    } else {
        Get-TenantDomainController
    })

    foreach ($dc in $known) {
        if ("$dc" -ieq $cand) { return "$dc" }
    }
    return ''
}

function Set-ADSessionScope {
    <#
        Sets the AD scope for this session, pinning a controller only if that
        actually works.

        TWO FAILURES THIS CLOSES.

        First, scope. Callers passed a candidate without saying which directory
        it belonged to, so it was validated against EVERY directory's
        controllers -- and one directory's controller would happily validate
        for an operation in another. That is the failure the docstring on
        Get-PinnableDomainController describes, arriving by a different door.

        Second, and worse: a rejected pin used to end the run. Set-ADServerSettings
        -PreferredServer runs a suitability check, and a value that is a domain
        rather than a server comes back as "Active Directory error 0x52 ...
        A local error occurred" -- which reads like the directory is broken. It
        is not. Forest-wide works perfectly well; the pin was only ever an
        optimisation. So a pin that is refused now falls back instead of
        throwing, and says so.

        Returns what it did, because "which controller am I talking to" is the
        first question when something later looks wrong.
    #>
    param(
        [string]$Candidate,
        [string]$DirectoryId
    )

    # FOREST-WIDE FIRST, PIN SECOND, AND NEVER A RECOVERY CALL.
    #
    # The first version pinned and, if that was refused, issued a second
    # Set-ADServerSettings to get back to forest-wide. Re-entering the cmdlet
    # immediately after a failed suitability check is how the console came to
    # sit there doing nothing instead of failing fast.
    #
    # Ordering it this way removes the recovery path entirely: the scope that
    # must be right is set once, unconditionally, and the pin is an optional
    # extra layered on top. If the pin is refused, the session is ALREADY where
    # it needs to be and there is nothing to undo.
    Set-ADServerSettings -ViewEntireForest $true -ErrorAction Stop

    $pin = Get-PinnableDomainController -Candidate $Candidate -DirectoryId $DirectoryId
    if (-not $pin) {
        return [PSCustomObject]@{ Pinned = ''; Fallback = $false; Detail = 'Forest-wide; no controller pinned.' }
    }

    try {
        Set-ADServerSettings -PreferredServer $pin -ErrorAction Stop
        return [PSCustomObject]@{ Pinned = $pin; Fallback = $false; Detail = "Pinned to $pin." }
    } catch {
        return [PSCustomObject]@{
            Pinned   = ''
            Fallback = $true
            Detail   = ("'$pin' was refused as a preferred server, so this session stays forest-wide " +
                        "and Active Directory will locate a controller. That is usually because the " +
                        "value is a DOMAIN rather than a server -- check Tenant Manager. " +
                        "($($_.Exception.Message))")
        }
    }
}

function Get-TenantDomainController {
    <#
        Every configured domain controller, or those for one directory.

        For the DC dropdowns. Those used to enumerate one hardcoded array, so a
        tenant with different servers got a list of somebody else's -- offered
        as valid choices, which is worse than an empty list.

        Empty is a correct answer: a tenant that pins no controllers gets no
        dropdown entries, and the tools fall back to letting AD locate one.
    #>
    param([string]$DirectoryId)
    $dirs = if ($DirectoryId) { @(Get-TenantDirectory -Id $DirectoryId) } else { @(Get-TenantDirectory) }
    $out = New-Object System.Collections.ArrayList
    foreach ($d in $dirs) {
        foreach ($dc in @($d.DomainControllers)) {
            if ($dc -and -not $out.Contains("$dc")) { [void]$out.Add("$dc") }
        }
    }
    return @($out.ToArray())
}

$global:ToolkitDiscoveredDC = @{}

function Resolve-DirectoryDomainController {
    <#
        A REAL domain controller for a directory, safe to hand to an Exchange
        -DomainController or -PreferredServer. '' if none can be found.

        WHY THIS EXISTS. Exchange rejects a domain name with "Active Directory
        error 0x52", so a directory that names no controller in Tenant.json
        left its tools with nothing to pass -- and a call with no
        -DomainController runs on whatever the SESSION is scoped to. Session
        scope is process-wide, outlives the tool that set it, and cannot be
        cleared: -PreferredServer is a mandatory Fqdn, and -ViewEntireForest
        $true is documented to clear RecipientViewRoot and nothing else. So a
        an operation in one directory would leave a pin behind and the next
        operation in another inherited it. That is the 2026-08-12 failure,
        where an account was verified against the wrong directory's controller.

        Naming a controller per directory in Tenant Manager fixes it, but it
        forces a choice nobody wants to make: commit to one box, and own it
        when that box is down. Discovery gets the same determinism without the
        commitment -- DCLocator picks a live controller in the right domain,
        every time, and the answer is a real server name rather than a domain.

        ORDER, AND WHY:
          1. Controllers named in Tenant.json win. An explicit choice is still
             an explicit choice, and some directories genuinely need a specific
             server.
          2. Otherwise discover one in that directory's domain.
          3. Otherwise '' -- and the CALLER says so. A silent fallback to
             ambient scope is exactly the failure this closes.

        -Writable, because the callers enable mailboxes. A read-only replica
        would accept the connection and refuse the write, which reads as a
        permissions problem.

        Cached in a GLOBAL, not $script:. The console launches tools with the
        call operator, so they share the process but not the script scope; a
        $script: cache would be $null in every child and each tool would pay
        discovery again. Pass -Refresh after a controller goes down.
    #>
    param(
        [Parameter(Mandatory)][string]$DirectoryId,
        [switch]$Refresh
    )

    $key = "$DirectoryId".Trim()
    if (-not $key) { return '' }

    if (-not $Refresh -and $global:ToolkitDiscoveredDC.ContainsKey($key)) {
        return "$($global:ToolkitDiscoveredDC[$key])"
    }

    # 1. Configured wins.
    $named = @(Get-TenantDomainController -DirectoryId $key)
    if ($named.Count -gt 0 -and "$($named[0])".Trim()) {
        $global:ToolkitDiscoveredDC[$key] = "$($named[0])"
        return "$($named[0])"
    }

    # 2. Discover one.
    $dir = Get-TenantDirectory -Id $key
    $domain = if ($dir) { "$($dir.Domain)".Trim() } else { '' }
    if (-not $domain) { return '' }
    if (-not (Get-Command -Name 'Get-ADDomainController' -ErrorAction SilentlyContinue)) { return '' }

    try {
        # -DomainName, NOT -Domain. Microsoft's own examples on this page use
        # -Domain; the syntax block says -DomainName. Following the examples
        # binds nothing and discovery silently answers for the LOCAL domain,
        # which for an operator in the default directory is right by accident
        # and for anyone else is
        # wrong without saying so.
        #
        # HostName is the FQDN specifically BECAUSE -Discover was used -- with
        # -Identity or -Filter it is the DNSHostName attribute and can be
        # unset. Taken as @(...)[0] because a discovery result has been seen to
        # carry the value as a collection.
        # $dcHost, never $host. $Host is an automatic variable holding the
        # PowerShell host object, and assigning to it throws "Cannot overwrite
        # variable Host because it is read-only or constant" -- inside a try,
        # which would have swallowed it and returned '' forever.
        $found  = Get-ADDomainController -Discover -DomainName $domain -Writable -ErrorAction Stop
        $dcHost = "$(@($found.HostName)[0])".Trim()
        if ($dcHost) {
            $global:ToolkitDiscoveredDC[$key] = $dcHost
            return $dcHost
        }
    } catch {
        # Discovery failing is not an error worth throwing. The caller has a
        # working, if less deterministic, path and will report that it took it.
    }

    return ''
}

function Get-TenantExchangeServerUri {
    $v = (Get-TenantConfig).OnPremExchange
    if ($v -and $v.ServerUri) { return "$($v.ServerUri)" }
    return ''
}

function Get-TenantSpoAdminUrl {
    <#
        The SharePoint admin centre URL from Tenant.json.

        WRITTEN SINCE THE BEGINNING AND READ BY NOTHING. Tenant.json has always
        carried SharePoint.AdminUrl, Tenant Manager has always offered a box
        for it, and Get-SpoAdminUrl looked only at Settings.json and then
        guessed. So an engineer could type the correct admin centre into the
        tool whose whole job is tenant facts, save it, and watch Connect SPO
        use a derived URL that had never existed -- with nothing anywhere
        saying the value was ignored.

        Found 2026-08-13 on a tenant whose SharePoint label does not match its
        mail domain, which is the exact case the field exists for.
    #>
    $c = Get-TenantConfig
    if (-not $c) { return '' }
    if (-not $c.PSObject.Properties['SharePoint']) { return '' }
    $sp = $c.SharePoint
    if (-not $sp -or -not $sp.PSObject.Properties['AdminUrl']) { return '' }
    return "$($sp.AdminUrl)".Trim()
}

function Get-MailDomain {
    <#
        The tenant's primary mail domain, with no leading '@'.

        This was called in 23 places across 12 files and defined in none of
        them. The tenant refactor replaced hardcoded domains with calls to it
        and never wrote the function.

        It survived every check because every call site is inside a STRING:
        "...@$(Get-MailDomain)". A subexpression in a double-quoted string is
        not resolved until the string is built, so the parser is happy, the
        encoding is fine, and the manifest hashes cleanly. Nothing can catch
        this except running the code or resolving names against definitions --
        which is exactly the check that caught the invented Get-TenantDisplayName
        in Dashboard, and it was not re-run after the refactor.

        The failure was also quiet in a second place. Get-SpoAdminUrl falls back
        to this to build its label, inside a try, so the SPO settings hint
        reported "nothing could be derived" -- which reads as a missing setting
        rather than a missing function.

        Returns '' rather than throwing, matching Get-TenantConfig. Callers that
        BUILD AN ADDRESS from this (Disable-User's SMTP server and its
        DO_NOT_REPLY sender) get a malformed value from a blank domain rather
        than an error, so PrimaryDomain is effectively required in Tenant.json.
    #>
    $c = $null
    try { $c = Get-TenantConfig } catch { }
    if ($c -and $c.PrimaryDomain) { return "$($c.PrimaryDomain)".Trim().TrimStart('@').ToLower() }
    return ''
}

function Get-TenantLicenseGroup {
    return @((Get-TenantConfig).LicenseGroups)
}

function Resolve-TenantLicenseGroupName {
    <#
        The licence group for a directory, from Tenant.json.

        A licence group name sat hardcoded in the leaver tools while
        Tenant.json carried a LicenseGroups list -- Name, Sku, Directory
        -- that nothing read. Two tenants have no group by that name, so the
        step could never do anything there.

        WORSE THAN NOT WORKING, IT REPORTED 'Skipped'. The membership test was
        "is the user in a group whose CN is this", and in a tenant without the
        group the honest answer to that question is no. So the run came back
        green with "Not a member of ...", which is indistinguishable from the
        genuine case where a user really has no licence. A red step gets fixed;
        a green one that did nothing gets shipped.

        Returns '' when nothing matches. Callers must treat that as a
        configuration failure, not as "nothing to do".
    #>
    param([string]$DirectoryId, [string]$Sku)

    $groups = @(Get-TenantLicenseGroup | Where-Object { $_ -and "$($_.Name)".Trim() })
    if ($groups.Count -eq 0) { return '' }

    if ($DirectoryId) {
        $byDir = @($groups | Where-Object { "$($_.Directory)" -ieq "$DirectoryId" })
        # Entries with no Directory apply everywhere, so they are the fallback
        # rather than a miss -- a single-directory tenant has no reason to fill
        # that column in.
        if ($byDir.Count -eq 0) {
            $byDir = @($groups | Where-Object { -not "$($_.Directory)".Trim() })
        }
        if ($byDir.Count -gt 0) { $groups = $byDir }
    }

    if ($Sku) {
        $bySku = @($groups | Where-Object { "$($_.Sku)" -ieq "$Sku" })
        if ($bySku.Count -gt 0) { $groups = $bySku }
    }

    return "$($groups[0].Name)".Trim()
}

function Get-TenantTerminationLicenseGroup {
    <#
        The licence groups a TERMINATION removes from. Not all of them.

        LicenseGroups serves two unrelated purposes and nothing in the file said
        which entry was for which. A two-directory deployment has four: E1
        and E5 per directory.
        The E1 groups exist so the Disabled/Licensed report can find people who
        are disabled but still licensed -- they are a reporting input, and a
        termination must not touch them. The E5 groups are the ones a leaver
        actually holds.

        With no way to tell them apart, every reader had to guess.
        Resolve-TenantLicenseGroupName guessed "the first one" and removed
        terminated staff from the E1 group -- one they were never in --
        then reported the step Done while the E5 licence stayed assigned. Green,
        and the invoice unchanged.

        RemoveOnTerm makes the file state the intent. Returns EMPTY when nothing
        is flagged, and callers must treat that as a configuration failure
        rather than falling back to a guess -- falling back is what caused this.
    #>
    param([string]$DirectoryId)

    $groups = @(Get-TenantLicenseGroup | Where-Object { $_ -and "$($_.Name)".Trim() })
    if ($groups.Count -eq 0) { return @() }

    if ($DirectoryId) {
        $byDir = @($groups | Where-Object { "$($_.Directory)" -ieq "$DirectoryId" })
        if ($byDir.Count -eq 0) {
            $byDir = @($groups | Where-Object { -not "$($_.Directory)".Trim() })
        }
        if ($byDir.Count -gt 0) { $groups = $byDir }
    }

    # Absent property is FALSE, not true. A tenant that upgrades the toolkit
    # without opening Tenant Manager gets "nothing flagged" and a step that
    # refuses -- which is noisy and safe. The other default would silently
    # resume removing whatever it found.
    $flagged = @($groups | Where-Object {
        $_.PSObject.Properties['RemoveOnTerm'] -and [bool]$_.RemoveOnTerm
    })
    return @($flagged | ForEach-Object { "$($_.Name)".Trim() } | Where-Object { $_ })
}

function Get-TenantLicenseGroupNameList {
    <#
        EVERY licence group for a directory, not the first one.

        Resolve-TenantLicenseGroupName returns a single name, which is only safe
        when a directory has exactly one licence group. A deployment with two
        per directory -- E1 and E5 -- got whichever happened to be first in
        Tenant.json, and the disable pipeline removed the user from the E1
        group while an E5 user kept their licence. It then
        reported Done, because it had successfully done the wrong thing.

        A caller that needs to REMOVE a licence does not need to know which SKU
        the user holds. It needs the list, so it can act on whichever one the
        user is actually in. Asking the directory beats inferring from config.
    #>
    param([string]$DirectoryId)

    $groups = @(Get-TenantLicenseGroup | Where-Object { $_ -and "$($_.Name)".Trim() })
    if ($groups.Count -eq 0) { return @() }

    if ($DirectoryId) {
        $byDir = @($groups | Where-Object { "$($_.Directory)" -ieq "$DirectoryId" })
        if ($byDir.Count -eq 0) {
            $byDir = @($groups | Where-Object { -not "$($_.Directory)".Trim() })
        }
        if ($byDir.Count -gt 0) { $groups = $byDir }
    }
    return @($groups | ForEach-Object { "$($_.Name)".Trim() } | Where-Object { $_ })
}

function Get-FeatureRequestRecipient {
    <#
        Where the Feature Request tool sends requests.

        In Tenant.json, NOT in the script. A mail address carries a domain, and
        a domain written into a .ps1 is the one thing Test-Toolkit refuses --
        the toolkit ships to six tenants and a hardcoded address would follow it
        to all of them.

        The operator does not see this address, because the tool puts it on BCC
        rather than in a visible header. That is a courtesy, not a secret: it is
        plainly readable in a config file. Nothing here should be relied on to
        stay hidden from somebody who goes looking.

        Returns '' when unset, and the caller must refuse to send rather than
        guess at a recipient.
    #>
    $fr = (Get-TenantConfig).FeatureRequest
    if ($fr -and $fr.PSObject.Properties['Recipient']) { return "$($fr.Recipient)".Trim() }
    return ''
}

function Remove-RecipientAddress {
    <#
        Strips the Feature Request / Bug Fix recipient out of a piece of text.

        THE SUCCESS PATH NEVER NAMED IT. The leak is the FAILURE path: a catch
        block hands the ErrorRecord to Write-FailedLog, which copies
        Exception.Message into the Detail column -- and an SMTP rejection quotes
        the address it could not deliver to:

            Mailbox unavailable. The server response was:
            5.1.1 <someone@example.com> User unknown

        So the one row in the audit log that names the recipient is the row
        written when something went wrong, which is also the row somebody is
        most likely to read closely. The same message goes on screen in the
        status label, which is worse -- it is shown to the operator directly.

        Case-insensitive, because -replace is by default and a mail server may
        echo the address back in any case it likes.

        WHAT THIS IS AND IS NOT. It removes an address from text. It does not
        make the address secret: it is in Config\Tenant.json, which travels with
        every deploy and which anyone can open. Routing feedback through the
        buttons is a CONVENTION supported by not putting the address in front of
        people -- if it has to be unguessable, the answer is a shared mailbox or
        a distribution list as the recipient, not more redaction here.
    #>
    param([string]$Text)

    $t = "$Text"
    if (-not $t) { return $t }

    $r = Get-FeatureRequestRecipient
    if (-not $r) { return $t }

    return ($t -replace [regex]::Escape($r), '(the request recipient)')
}

function Get-TenantPimSetting {
    param([Parameter(Mandatory)][string]$Name)
    $p = (Get-TenantConfig).Pim
    if ($p -and $p.PSObject.Properties[$Name]) { return $p.$Name }
    return $null
}
#endregion

function ConvertTo-OpathValue {
    <#
        Escapes a value for an Exchange OPATH -Filter string.

        THE SIBLING OF ConvertTo-LdapFilterValue BELOW, and they are not
        interchangeable. LDAP escapes backslash, parentheses, asterisk and NUL
        as hex; OPATH cares about one character, the single quote, and escapes
        it by doubling. Using either escaper on the other kind of filter
        produces a query that is wrong rather than one that fails.

        Every caller here is building a filter around something an operator
        typed -- an office name, a partial mailbox name -- so one apostrophe in
        an Irish surname or an O'Brien Staffing office is enough to break the
        query. It is a correctness fix, not a security one: these filters read,
        and the operator already has rights to read.
    #>
    param([string]$Value)
    return ("$Value" -replace "'", "''")
}

function ConvertTo-LdapFilterValue {
    <#
        Escapes a value for use inside an LDAP filter, per RFC 4515.

        THIS EXISTS BECAUSE THE INLINE VERSION THREW EVERY TIME IT RAN.

        It was written twice, identically, as a loop over pairs:

            foreach ($p in @(@('\','\5c'), ..., @([char]0,'\00'))) {
                $esc = $esc.Replace($p[0], $p[1])
            }

        [char]0 makes the first argument a CHAR, and String.Replace has a
        (char, char) overload as well as a (string, string) one. PowerShell
        binds to (char, char) and then cannot turn the four-character string
        '\00' into a single char, so every call died on the last iteration with
        "String must be exactly one character long" -- a message that says
        nothing about LDAP, addresses, or the function it came from.

        Every OTHER pair in that loop is string-to-string and binds correctly,
        which is why it looked right: four iterations work and the fifth throws.

        BACKSLASH IS ESCAPED FIRST, and the order is load-bearing. Escaping it
        after the others would find the backslashes this function had just
        introduced and turn \28 into \5c28 -- the same trap as the regex
        replacement in the DL tooling, arrived at from a different direction.
    #>
    param([string]$Value)

    $t = "$Value"
    $t = $t.Replace('\', '\5c')
    $t = $t.Replace('(', '\28').Replace(')', '\29').Replace('*', '\2a')
    # Cast to STRING explicitly. This is the argument that caused the fault
    # above, and leaving it as a char would reintroduce it exactly.
    return $t.Replace([string]([char]0), '\00')
}

function ConvertTo-MatchableName {
    <#
        Normalises the characters that look identical in a group name but are
        not.

        'E5 License - Full' typed in Word comes back with the hyphen turned
        into an EN DASH (U+2013), which renders almost the same as a real
        hyphen (U+002D). A name copied by eye then matches nothing, and the
        report says the group does not exist. Non-breaking spaces and smart
        quotes fail the same way and are just as invisible.

        This comment names the character rather than showing it. An earlier
        version pasted a real en dash in here to illustrate the point, which
        made a comment about invisible characters the only non-ASCII byte in
        the toolkit -- and the first thing an encoding round-trip would eat.
        See ARCHITECTURE 14i. The code below is the safe way to write one.
    #>
    param([string]$Text)
    if (-not $Text) { return '' }
    $t = "$Text"
    foreach ($d in @([char]0x2010, [char]0x2011, [char]0x2012, [char]0x2013,
                     [char]0x2014, [char]0x2015, [char]0x2212)) {
        $t = $t.Replace($d, '-')                       # dash family -> hyphen
    }
    $t = $t.Replace([char]0x00A0, ' ')                 # non-breaking space
    $t = $t.Replace([char]0x2018, "'").Replace([char]0x2019, "'")
    $t = $t.Replace([char]0x201C, '"').Replace([char]0x201D, '"')
    return ($t -replace '\s+', ' ').Trim().ToLower()
}

function Resolve-LicenseGroupObject {
    <#
        Finds a licence group by name in whichever forest holds it.

        Exact match first. If that fails, every group in the forest is compared
        on its NORMALISED name, so a Word-style dash or a stray non-breaking
        space cannot make a group invisible.

        A fallback match returns MatchedOn = 'normalised' and the real name, so
        callers can report the discrepancy rather than quietly papering over an
        inconsistency that is worth fixing at the source.
    #>
    param([string]$Name)

    $servers = @(
        # Every configured directory, not a hardcoded pair. A tenant with one
        # gets one; a cloud-only tenant gets none and the caller finds nothing,
        # which is the correct answer rather than an error.
        Get-TenantDirectory | ForEach-Object {
            [PSCustomObject]@{ Label = "$($_.Label)"; Server = (Get-TenantDirectoryServer -Id $_.Id) }
        }
    )

    foreach ($s in $servers) {
        if (-not $s.Server) { continue }
        try {
            $g = Get-ADGroup -Identity $Name -Server $s.Server -ErrorAction Stop
            if ($g) {
                return [PSCustomObject]@{
                    Group = $g; Server = $s.Server; Forest = $s.Label
                    MatchedOn = 'exact'; ActualName = "$($g.Name)"
                }
            }
        } catch { }
    }

    $target = ConvertTo-MatchableName $Name
    if (-not $target) { return $null }

    foreach ($s in $servers) {
        if (-not $s.Server) { continue }
        try {
            # Narrow the sweep with the first plain-text word so this is not a
            # full enumeration of every group in the forest.
            $stem = ($Name -split '[^\w]')[0]
            $filt = if ($stem) { "Name -like '$stem*'" } else { '*' }
            foreach ($g in (Get-ADGroup -Filter $filt -Server $s.Server -ErrorAction Stop)) {
                if ((ConvertTo-MatchableName $g.Name) -eq $target) {
                    return [PSCustomObject]@{
                        Group = $g; Server = $s.Server; Forest = $s.Label
                        MatchedOn = 'normalised'; ActualName = "$($g.Name)"
                    }
                }
            }
        } catch { }
    }
    return $null
}

function Get-LicensedUserMap {
    <#
        Returns @{ <lowercase DN> = @{ Groups=..; Skus=..; Environment=.. } }
        for every user in a licence group, plus the list of groups that could
        not be resolved.

        Built once and reused: several reports need "is this person licensed",
        and asking per user would be a call each.
    #>
    param([scriptblock]$Progress)

    $map      = @{}
    $notFound = @()
    $renamed  = @()
    $totals   = @{}

    $i = 0
    foreach ($lg in (Get-TenantLicenseGroup)) {
        $i++
        if ($Progress) { & $Progress "Reading licence group '$($lg.Name)'..." }
        $found = Resolve-LicenseGroupObject -Name $lg.Name
        if (-not $found) { $notFound += $lg.Name; continue }

        # Surfaced here rather than left to each caller: the discrepancy is
        # discovered in this function, so throwing it away would mean every
        # caller either re-derives it or silently loses it.
        if ($found.MatchedOn -eq 'normalised') {
            $renamed += "$($lg.Name)  ->  actual AD name: $($found.ActualName)"
        }

        try {
            $members = @(Get-ADGroupMember -Identity $found.Group.DistinguishedName `
                            -Server $found.Server -Recursive -ErrorAction Stop |
                         Where-Object { $_.objectClass -eq 'user' })
        } catch { $notFound += "$($lg.Name) (unreadable)"; continue }

        $totals[$lg.Name] = $members.Count
        foreach ($m in $members) {
            $k = "$($m.DistinguishedName)".ToLower()
            if (-not $map.ContainsKey($k)) {
                $map[$k] = [PSCustomObject]@{
                    Groups      = $lg.Name
                    Skus        = $lg.Sku
                    Environment = $lg.Environment
                    Server      = $found.Server
                }
            } else {
                $e = $map[$k]
                $e.Groups = "$($e.Groups); $($lg.Name)"
                if ($e.Skus -notlike "*$($lg.Sku)*") { $e.Skus = "$($e.Skus); $($lg.Sku)" }
            }
        }
    }
    return [PSCustomObject]@{ Map = $map; NotFound = $notFound; Renamed = $renamed; Totals = $totals }
}

function Get-SimpleMailboxType {
    <#
        RecipientTypeDetails collapsed to Shared, Resource, or User.

        Exchange reports at least eight values here and an operator triaging a
        ticket cares about three: is this a person's mailbox, a shared one, or a
        room or piece of equipment. Those three imply different answers to
        "should it be licensed", "should it have a password", and "why can
        nobody sign in".

        Room, Equipment and Scheduling are one answer -- RESOURCE. They are
        booked rather than used, and every rule that matters treats them alike.
        The caller keeps the raw value too, so the distinction is never lost,
        only set aside.

        An unrecognised value returns the RAW STRING rather than falling into
        'User'. Microsoft adds types, and a new one silently reported as a user
        mailbox is a wrong answer stated confidently; the raw value is at worst
        an unfamiliar one the operator can look up.
    #>
    param([string]$RecipientTypeDetails)

    $t = "$RecipientTypeDetails".Trim()
    if (-not $t) { return 'Unknown' }

    switch -Regex ($t) {
        '^SharedMailbox$'                                              { return 'Shared' }
        '^(Room|Equipment|Scheduling)Mailbox$'                         { return 'Resource' }
        '^(User|Remote(User)?|Legacy|Linked|MailUser)Mailbox$'         { return 'User' }
        '^RemoteSharedMailbox$'                                        { return 'Shared' }
        '^Remote(Room|Equipment)Mailbox$'                              { return 'Resource' }
        default                                                        { return $t }
    }
}

function Get-UserLicenseState {
    <#
        Is ONE person licensed, and out of which group.

        Get-LicensedUserMap answers this for everybody, by enumerating every
        member of every licence group. That is the right shape for a report and
        the wrong shape for a single lookup -- Mailbox Health would spend
        several seconds reading thousands of members to answer a question about
        one mailbox.

        NOT GRAPH. Where licensing is group-based, the licence a person holds
        is a fact about AD group membership and is readable with the
        permissions the operator already has. This deliberately does not report
        the SKU as Microsoft sees it -- it reports the group that drives the
        assignment, which is the thing an operator can actually act on.

        NESTING IS CHECKED, but only when the direct answer is no.

        Get-LicensedUserMap reads members -Recursive while the disable pipeline
        tests $user.MemberOf, which is direct-only. Those two disagree for a
        user who is licensed through a nested group, and this function would
        have had to pick a side. It checks direct membership first because that
        is what a group-based licensing model actually does and it costs
        nothing -- memberOf is
        already on the object -- and only falls back to an LDAP_MATCHING_RULE_IN_CHAIN
        query when that finds nothing. So a nested licence is reported rather
        than missed, and the common case still costs zero extra round trips.

        Returns Known = $false when the question could not be answered, which
        callers MUST render differently from Licensed = $false. "This person has
        no licence" and "I could not reach a domain controller" look identical
        on screen otherwise, and only one of them is a reason to raise a ticket.
    #>
    param(
        [string]$Smtp,
        # New User works in logon names and already knows the directory and the
        # controller, because it just wrote to them. Making it round-trip
        # through an SMTP address it does not have -- for an account whose
        # mailbox was created seconds ago -- would be a lookup that can fail for
        # reasons having nothing to do with the licence.
        [string]$SamAccountName,
        [string]$DirectoryId,
        [string]$Server,
        [hashtable]$Cache
    )

    if (-not "$Smtp".Trim() -and -not "$SamAccountName".Trim()) {
        throw 'Get-UserLicenseState needs -Smtp or -SamAccountName.'
    }

    $who = if ("$SamAccountName".Trim()) { "sam:$SamAccountName" } else { "smtp:$Smtp" }
    $key = "lic::$($who.ToLowerInvariant())"
    if ($Cache -and $Cache.ContainsKey($key)) { return $Cache[$key] }

    $out = [PSCustomObject]@{
        Smtp       = $Smtp
        Known      = $false
        Licensed   = $false
        Groups     = @()
        Skus       = @()
        MatchedVia = ''
        Reason     = ''
    }

    $finish = {
        if ($Cache) { $Cache[$key] = $out }
        return $out
    }

    # A cloud-only tenant has no on-premises groups to read, so the honest
    # answer is "unknown here", not "unlicensed".
    if (-not (Test-TenantCapability -Name 'OnPremAD')) {
        $out.Reason = 'This tenant has no on-premises directory; licence state is not readable without Graph.'
        return (& $finish)
    }

    $ad = Test-AdAvailable
    if (-not $ad.Available) {
        $out.Reason = "Active Directory is not available: $($ad.Detail)"
        return (& $finish)
    }

    $bySam = [bool]("$SamAccountName".Trim())
    $raw   = if ($bySam) { "$SamAccountName".Trim() } else { "$Smtp".Trim() }
    $esc   = ConvertTo-LdapFilterValue -Value $raw
    $filter = if ($bySam) {
        "(sAMAccountName=$esc)"
    } else {
        "(|(mail=$esc)(userPrincipalName=$esc)(proxyAddresses=smtp:$esc))"
    }

    # A caller that already knows where the account lives says so, and this
    # asks one directory instead of walking all of them.
    $candidates = @()
    if ("$Server".Trim()) {
        $candidates = @([PSCustomObject]@{ Id = "$DirectoryId"; Server = "$Server".Trim() })
    } elseif ("$DirectoryId".Trim()) {
        $candidates = @([PSCustomObject]@{ Id = "$DirectoryId"; Server = (Get-TenantDirectoryServer -Id $DirectoryId) })
    } else {
        $candidates = @(Get-TenantDirectory | ForEach-Object {
            [PSCustomObject]@{ Id = "$($_.Id)"; Server = (Get-TenantDirectoryServer -Id $_.Id) }
        })
    }

    $user = $null; $server = ''; $dirId = ''
    foreach ($d in $candidates) {
        if (-not $d.Server) { continue }
        try {
            $u = @(Get-ADUser -LDAPFilter $filter -Server $d.Server -Properties MemberOf -ErrorAction Stop)
        } catch { continue }
        if (@($u).Count -gt 0) { $user = $u[0]; $server = $d.Server; $dirId = "$($d.Id)"; break }
    }

    if (-not $user) {
        $out.Reason = 'No on-premises account matches.'
        return (& $finish)
    }

    # Past this point the directory answered, so any result -- including "not
    # licensed" -- is a real finding rather than a failure to look.
    $out.Known = $true

    $names = @(Get-TenantLicenseGroupNameList -DirectoryId $dirId)
    if ($names.Count -eq 0) {
        $out.Known  = $false
        $out.Reason = "No licence groups are configured for directory '$dirId' in Tenant Manager."
        return (& $finish)
    }

    $hits = @()

    # Direct membership, straight off the object already in hand.
    foreach ($n in $names) {
        if (@($user.MemberOf) | Where-Object { $_ -like "CN=$n,*" }) { $hits += $n }
    }
    if ($hits.Count -gt 0) { $out.MatchedVia = 'direct' }

    # Only now is the more expensive question worth asking.
    if ($hits.Count -eq 0) {
        foreach ($n in $names) {
            # The GROUP is cached under its own key, separately from the
            # per-account result above. Resolving a group name to a DN is the
            # expensive half of this branch and the answer is identical for
            # every account in a batch, so caching only the account would leave
            # the cost exactly where it is.
            $gKey = "licgrp::$($n.ToLowerInvariant())"
            if ($Cache -and $Cache.ContainsKey($gKey)) {
                $grp = $Cache[$gKey]
            } else {
                $grp = Resolve-LicenseGroupObject -Name $n
                if ($Cache) { $Cache[$gKey] = $grp }
            }
            if (-not $grp) { continue }
            # Same escaping as the address above, through the same function --
            # a DN goes into an LDAP filter and needs identical treatment. This
            # was a second hand-rolled loop until the first one turned out to be
            # broken; two implementations of one rule is how that happened.
            $gDn = ConvertTo-LdapFilterValue -Value "$($grp.Group.DistinguishedName)"
            try {
                $chain = @(Get-ADUser -LDAPFilter "(memberOf:1.2.840.113556.1.4.1941:=$gDn)" `
                              -SearchBase "$($user.DistinguishedName)" -SearchScope Base `
                              -Server $server -ErrorAction Stop)
                if (@($chain).Count -gt 0) { $hits += $n }
            } catch { }
        }
        if ($hits.Count -gt 0) { $out.MatchedVia = 'nested' }
    }

    $out.Licensed = ($hits.Count -gt 0)
    $out.Groups   = @($hits)
    $out.Skus     = @(
        Get-TenantLicenseGroup |
            Where-Object { $hits -contains "$($_.Name)".Trim() } |
            ForEach-Object { "$($_.Sku)".Trim() } |
            Where-Object { $_ } | Select-Object -Unique
    )
    return (& $finish)
}

function Format-UserLicenseState {
    <#
        The one-line rendering, so Mailbox Health and New User cannot drift into
        describing the same state with different words.
    #>
    param([Parameter(Mandatory)]$State)

    if (-not $State.Known)    { return "Unknown -- $($State.Reason)" }
    if (-not $State.Licensed) { return 'No -- not a member of any configured licence group.' }

    $txt = "Yes -- $($State.Groups -join ', ')"
    if ($State.MatchedVia -eq 'nested') { $txt = "$txt  (via a nested group)" }
    return $txt
}
#endregion

#region Parsing (moved unchanged -- original lines 550-559) ------------------
function Resolve-ADAccountBySmtp {
    <#
        Finds the on-prem AD account behind a cloud SMTP address, across both
        forests, and reports whether it is enabled.

        Shared because two reports need the same answer and it must not drift:
        "is the person behind this cloud object still an employee" is decided by
        on-prem AD in this environment, the same way licence state is (see
        Get-TenantLicenseGroup).

        Searched by mail, userPrincipalName AND proxyAddresses, because a cloud
        object's primary SMTP is frequently a proxy address on-prem rather than
        the mail attribute -- matching only on mail produces false "not found"
        results that read as "this person left" when they did not.

        Pass the same -Cache hashtable across a whole run. Group owners repeat
        heavily, and without caching a few hundred groups becomes thousands of
        identical directory lookups.

        Returns Found = $false with Searched populated when nothing matched, so
        the caller can distinguish "no such account" from "could not ask".
    #>
    param(
        [Parameter(Mandatory)][string]$Smtp,
        [hashtable]$Cache
    )

    $key = $Smtp.ToLowerInvariant()
    if ($Cache -and $Cache.ContainsKey($key)) { return $Cache[$key] }

    # LDAP special characters, per RFC 4515. Rare in an address, fatal if present.
    #
    # THIS USED TO BE AN INLINE LOOP AND IT THREW ON EVERY CALL -- see
    # ConvertTo-LdapFilterValue for why. It was not noticed because the throw
    # happens before the AD query, so the function never reported "no such
    # account"; it failed with a MethodException about converting "\00" to a
    # char, and the callers that wrap this in a try reported the address as
    # unresolvable. "Could not ask" was being displayed as "does not exist",
    # which is the exact distinction the rest of this function is built around.
    $esc = ConvertTo-LdapFilterValue -Value $Smtp
    $filter = "(|(mail=$esc)(userPrincipalName=$esc)(proxyAddresses=smtp:$esc))"

    $servers = @(Get-TenantDirectory | ForEach-Object { Get-TenantDirectoryServer -Id $_.Id })
    $result = [PSCustomObject]@{
        Smtp = $Smtp; Found = $false; Enabled = $null; SamAccountName = ''
        DisplayName = ''; Server = ''; Environment = ''; Error = ''
    }
    $errors = @()

    foreach ($srv in $servers) {
        if (-not $srv) { continue }
        try {
            $u = @(Get-ADUser -LDAPFilter $filter -Server $srv `
                        -Properties Enabled, mail, DisplayName -ErrorAction Stop)
        } catch { $errors += "$srv : $($_.Exception.Message)"; continue }

        if (@($u).Count -gt 0) {
            $hit = $u[0]
            $result.Found          = $true
            $result.Enabled        = [bool]$hit.Enabled
            $result.SamAccountName = "$($hit.SamAccountName)"
            $result.DisplayName    = "$($hit.DisplayName)"
            $result.Server         = $srv
            $result.Environment    = "$((@(Get-TenantDirectory) | Where-Object { (Get-TenantDirectoryServer -Id $_.Id) -eq $srv } | Select-Object -First 1).Id)"
            break
        }
    }

    # Every directory refused: that is "could not ask", not "does not exist".
    if (-not $result.Found -and $errors.Count -eq @($servers).Count) {
        $result.Error = ($errors | Select-Object -First 1)
    }

    if ($Cache) { $Cache[$key] = $result }
    return $result
}

#region Configuration files ---------------------------------------------------
<#
    THE CODE IS MINE, THE CONFIGURATION IS YOURS.

    Every .json the console reads lives in Config\, and nothing in Config\ is
    ever shipped over. That makes the update instruction one rule with no
    exceptions to remember:

        replace the files, leave Config\ alone.

    A per-file exception list ("replace everything except these four, and don't
    miss Settings.json") is the kind of instruction that gets followed correctly
    nine times and then not. A folder boundary cannot be half-followed.

    Config\ holds:
        Tools.json      which tools appear, where, in what order
        Sections.json   the console's sections and row counts
        Themes.json     colour palettes
        Settings.json   mail domain, DC, folders, Edge profiles
        Images\         logos

    NOT in Config\: reports\, Logs\ and captures\. Those are OUTPUT, a different
    category again, and they are already relocatable through Settings.json.
    Folding them in here would blur the boundary this folder exists to make
    obvious.
#>
$script:ConfigFolderName = 'Config'
$script:ConfigFileNames  = @('Tools.json', 'Sections.json', 'Themes.json', 'Settings.json')

function Get-ConfigFolder {
    <#
        The Config folder, created on demand. ONE definition used by both reads
        and writes -- if the resolver could disagree with itself, ToolManager
        would write somewhere the console does not read and the two files would
        silently diverge.
    #>
    $dir = Join-Path $PSScriptRoot $script:ConfigFolderName
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }
    return $dir
}

function Get-ConfigFilePath {
    <# Full path to a named config file inside Config\. #>
    param([Parameter(Mandatory)][string]$Name)
    return (Join-Path (Get-ConfigFolder) $Name)
}

function Get-ConfigImageFolder {
    <# Where logos live. Config\Images -- a logo is your content, not mine. #>
    $dir = Join-Path (Get-ConfigFolder) 'Images'
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }
    return $dir
}

function Invoke-ConfigMigration {
    <#
        Moves configuration left in the toolkit root into Config\, once.

        Runs on launch so an existing install fixes itself when the new files
        are copied over the top. Deliberately a MOVE, not a copy: leaving the old
        file behind would create exactly the two-locations-that-disagree problem
        this folder exists to prevent.

        A file already present in Config\ wins -- the root copy is then stale by
        definition, and is renamed rather than deleted so nothing is destroyed by
        a tool the user did not ask to reorganise anything.

        Returns what it did, so the console can say so rather than moving the
        user's files silently.
    #>
    $moved = @()
    $dir   = Get-ConfigFolder

    foreach ($n in $script:ConfigFileNames) {
        $old = Join-Path $PSScriptRoot $n
        $new = Join-Path $dir $n
        if (-not (Test-Path -LiteralPath $old)) { continue }
        try {
            if (Test-Path -LiteralPath $new) {
                $bak = "$old.superseded"
                Move-Item -LiteralPath $old -Destination $bak -Force -ErrorAction Stop
                $moved += "$n was already in Config\ -- the root copy is now $n.superseded"
            } else {
                Move-Item -LiteralPath $old -Destination $new -Force -ErrorAction Stop
                $moved += "$n moved into Config\"
            }
        } catch {
            $moved += "$n could not be moved: $($_.Exception.Message)"
        }
    }

    # Logos too, from the old Themes\Images.
    $oldImg = Join-Path $PSScriptRoot 'Themes\Images'
    if (Test-Path -LiteralPath $oldImg) {
        $newImg = Get-ConfigImageFolder
        foreach ($f in @(Get-ChildItem -LiteralPath $oldImg -File -ErrorAction SilentlyContinue)) {
            $dest = Join-Path $newImg $f.Name
            if (Test-Path -LiteralPath $dest) { continue }
            try {
                Move-Item -LiteralPath $f.FullName -Destination $dest -Force -ErrorAction Stop
                $moved += "$($f.Name) moved into Config\Images\"
            } catch { }
        }
    }
    return $moved
}
#endregion

#region Settings.json ---------------------------------------------------------
<#
    Locations and per-machine values, kept OUT of this file.

    Two reasons this is a separate file rather than constants here:

      * This file gets replaced when the toolkit is updated, and a setting
        stored in it disappears on every update -- silently, with a symptom that
        looks nothing like the cause.
      * Hardcoded absolute paths assume a drive letter exists. 'E:\Scripts\...'
        works on the machine it was written on and fails on a server built
        differently, at the moment it tries to write a log.

    The tenant constants above stay hardcoded on purpose -- this is the
    tenant's own facts and that is the point of it. Settings.json only carries
    what varies by MACHINE.
#>
$script:SettingsFileName = 'Settings.json'

function Get-SettingsFilePath {
    return (Get-ConfigFilePath -Name $script:SettingsFileName)
}

function Get-ToolkitSettings {
    <#
        Reads Settings.json, or returns empty values if it is absent -- a
        missing file is the normal first-run state, not an error. Cached in a
        global rather than $script: so tools the console launches with '&' see
        the same answer instead of each re-reading it.
    #>
    param([switch]$Refresh)

    if ($global:ToolkitSettings -and -not $Refresh) { return $global:ToolkitSettings }

    $settings = [PSCustomObject]@{
        # SpoAdminUrl and ExchangeServerUri are NOT here, and their absence is
        # the design. Both describe the TENANT, so both live in Tenant.json and
        # are edited in Tenant Manager. Removed 2026-08-13; see Get-SpoAdminUrl
        # for what having them in two files actually cost.
        ReportsPath       = ''
        LogPath           = ''
        RetryQueuePath    = ''
        OperatorEmail     = ''
        BrowserProfiles   = $null
    }
    $path = Get-SettingsFilePath

    if (Test-Path -LiteralPath $path) {
        try {
            $j = (Get-Content -LiteralPath $path -Raw -ErrorAction Stop) | ConvertFrom-Json -ErrorAction Stop
            foreach ($k in @('ReportsPath', 'LogPath', 'RetryQueuePath', 'OperatorEmail')) {
                if ($j.PSObject.Properties[$k]) { $settings.$k = "$($j.$k)".Trim() }
            }

            # DisableRequestPath was this setting's name when the folder held a
            # disable list as well as the queue. Read it when the new key is
            # absent so an existing Settings.json keeps working -- a rename that
            # silently resets an operator's configured path to the default is a
            # rename that loses their queue.
            if (-not $settings.RetryQueuePath -and $j.PSObject.Properties['DisableRequestPath']) {
                $settings.RetryQueuePath = "$($j.DisableRequestPath)".Trim()
            }
            if ($j.PSObject.Properties['BrowserProfiles']) { $settings.BrowserProfiles = $j.BrowserProfiles }
        } catch {
            Write-Warning "Settings.json could not be read and is being ignored: $($_.Exception.Message)"
        }
    }
    $global:ToolkitSettings = $settings
    return $settings
}

function Save-ToolkitSettings {
    param(
        [string]$ReportsPath = '',
        [string]$LogPath = '',
        [string]$RetryQueuePath = '',
        [string]$OperatorEmail = '',
        [hashtable]$BrowserProfiles
    )
    $profiles = [ordered]@{}
    if ($BrowserProfiles) {
        foreach ($k in ($BrowserProfiles.Keys | Sort-Object)) { $profiles[$k] = "$($BrowserProfiles[$k])".Trim() }
    }
    $obj = [ordered]@{
        ReportsPath       = "$ReportsPath".Trim()
        LogPath           = "$LogPath".Trim()
        RetryQueuePath    = "$RetryQueuePath".Trim()
        OperatorEmail     = "$OperatorEmail".Trim()
        BrowserProfiles   = $profiles
    }
    $json = ([PSCustomObject]$obj | ConvertTo-Json -Depth 3)
    $enc  = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText((Get-SettingsFilePath), ($json -replace "`r?`n", "`r`n"), $enc)
    $global:ToolkitSettings = $null
    return (Get-ToolkitSettings -Refresh)
}

function Resolve-ToolkitFolder {
    <#
        A configured folder, or one beside the toolkit. A relative configured
        value is relative to the TOOLKIT, not to whatever directory the shell
        happened to be sitting in when the console launched.

        An unreachable share falls back beside the toolkit with a warning rather
        than taking the tool down -- failing to write a log should not stop the
        work the log was recording.
    #>
    param([string]$Configured, [Parameter(Mandatory)][string]$DefaultLeaf)

    $path = if ($Configured) {
                if ([System.IO.Path]::IsPathRooted($Configured)) { $Configured }
                else { Join-Path $PSScriptRoot $Configured }
            } else { Join-Path $PSScriptRoot $DefaultLeaf }

    if (-not (Test-Path -LiteralPath $path)) {
        try { New-Item -Path $path -ItemType Directory -Force | Out-Null }
        catch {
            Write-Warning "Could not create '$path' ($($_.Exception.Message)). Falling back to the toolkit folder."
            $path = Join-Path $PSScriptRoot $DefaultLeaf
            if (-not (Test-Path -LiteralPath $path)) { New-Item -Path $path -ItemType Directory -Force | Out-Null }
        }
    }
    return $path
}

function Get-ReportPath {
    <#
        Where report CSVs are written. ONE definition -- this was previously
        written out in ten separate report scripts, the same
        one-value-many-copies shape that let a duplicated domain controller
        drift out of step.
    #>
    return (Resolve-ToolkitFolder -Configured (Get-ToolkitSettings).ReportsPath -DefaultLeaf 'reports')
}

function Get-LogPath {
    <# Where CompletedTasks.csv is written. #>
    return (Resolve-ToolkitFolder -Configured (Get-ToolkitSettings).LogPath -DefaultLeaf 'Logs')
}

function Get-RetryQueuePath {
    <#
        Folder holding retry queues -- work a tool could not finish and intends
        to pick up again. Disable-User writes disable-retry.csv here today; the
        folder is named for the job rather than for one tool so others can queue
        alongside it without a second setting.

        CREATION IS DELIBERATELY ASYMMETRIC.

        The default, RetryQueue beside the toolkit, is ours and is created on
        demand like Reports and Logs. A configured path is not ours and is never
        created: if the operator typed it and it does not exist, the honest
        answer is that the path is wrong. Creating it would turn "your path is
        wrong" into "the queue is empty", which reads as good news.

        That distinction is the whole rule, and it was previously stated and
        then broken -- the old helper promised never to create the folder while
        Export-RetryQueue created it unconditionally, so a mistyped path
        produced an empty queue in the wrong place and looked fine.
    #>
    param([switch]$Ensure)

    $c = (Get-ToolkitSettings).RetryQueuePath
    if ($c) {
        if ([System.IO.Path]::IsPathRooted($c)) { return $c }
        return (Join-Path $PSScriptRoot $c)
    }

    $default = Join-Path $PSScriptRoot 'RetryQueue'
    if ($Ensure -and -not (Test-Path -LiteralPath $default)) {
        New-Item -Path $default -ItemType Directory -Force | Out-Null
    }
    return $default
}

function Test-RetryQueueConfigured {
    <#
        True when the folder a caller is about to write to actually exists.

        Separated from Get-RetryQueuePath so a caller can tell "not there yet,
        and I may create it" from "the operator pointed this somewhere that does
        not exist" without re-deriving the path or duplicating the rule above.
    #>
    $configured = "$((Get-ToolkitSettings).RetryQueuePath)".Trim()
    $path = Get-RetryQueuePath
    return [PSCustomObject]@{
        Path         = $path
        IsDefault    = -not $configured
        Exists       = (Test-Path -LiteralPath $path)
        MayBeCreated = -not $configured
    }
}
function Get-OperatorEmail {
    <#
        Where a tool sends the operator's own copy of anything it mails out.

        This exists because the address was hardcoded to one person, so whoever
        ran the tool was never the one who received the copy -- which turned the
        copy from a check on your own work into mail about somebody else's.

        Settings.json first, then the running account's own mail attribute in
        AD. Returns '' when neither answers, and callers must treat that as
        "send no copy" rather than falling back to a default: a copy going
        quietly to the wrong person is worse than no copy at all.
    #>
    $configured = "$((Get-ToolkitSettings).OperatorEmail)".Trim()
    if ($configured) { return $configured }

    if ($global:ToolkitOperatorEmailProbed) { return "$($global:ToolkitOperatorEmail)" }
    $global:ToolkitOperatorEmailProbed = $true
    $global:ToolkitOperatorEmail = ''
    try {
        $sam = $env:USERNAME
        if ($sam) {
            $u = Get-ADUser -Identity $sam -Properties mail -ErrorAction Stop
            if ($u -and $u.mail) { $global:ToolkitOperatorEmail = "$($u.mail)" }
        }
    } catch { }
    return "$($global:ToolkitOperatorEmail)"
}

#endregion

function Get-TargetSmtp {
    param([string]$Text)
    if (-not $Text) { return $null }
    $t = $Text.Trim()
    if ($t -eq '') { return $null }
    if ($t -like '*@*') { return $t }
    return "$t@$(Get-MailDomain)"
}
#endregion

#region Error formatting (moved unchanged -- original lines 561-587) --------
function Show-HandlerError {
    <#
        Surfaces an error from a WinForms event handler as something readable.

        An exception that escapes a Click handler is caught by .NET, not by
        PowerShell, and you get a CLR crash dialog: a hundred lines of
        System.Management.Automation internals, the loaded assembly list, and
        NOT the one thing you need -- which script and which line. This puts the
        PowerShell view back.

        Also writes a Failed row to the audit log, because an error the operator
        dismissed is otherwise the one kind of failure that leaves no trace.
    #>
    param(
        [Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord,
        [string]$Context = 'Action'
    )
    $inv  = $ErrorRecord.InvocationInfo
    $line = if ($inv) { "$($inv.ScriptName):$($inv.ScriptLineNumber)" } else { '(unknown location)' }
    $stmt = if ($inv -and $inv.Line) { "$($inv.Line)".Trim() } else { '' }

    $msg = "$Context failed." + [Environment]::NewLine + [Environment]::NewLine +
           "$($ErrorRecord.Exception.Message)" + [Environment]::NewLine + [Environment]::NewLine +
           "Where : $line"
    if ($stmt) { $msg += [Environment]::NewLine + "Line  : $stmt" }
    if ($ErrorRecord.FullyQualifiedErrorId) {
        $msg += [Environment]::NewLine + "Id    : $($ErrorRecord.FullyQualifiedErrorId)"
    }

    [void][System.Windows.Forms.MessageBox]::Show($msg, 'Error', 'OK', 'Error')
    try {
        Write-CompletedLog -Action 'Error' -Target $Context -Result 'Failed' `
            -Detail "$($ErrorRecord.Exception.Message) [$line]"
    } catch { }
}

function Format-ErrorDetail {
    param(
        [Parameter(Mandatory)] $ErrorRecord,
        [string]$Context = ''
    )
    $ex   = $ErrorRecord.Exception
    $cmd  = if ($ErrorRecord.InvocationInfo -and $ErrorRecord.InvocationInfo.MyCommand) {
                $ErrorRecord.InvocationInfo.MyCommand.Name
            } else { '(cmdlet unknown)' }
    $line = if ($ErrorRecord.InvocationInfo) { $ErrorRecord.InvocationInfo.ScriptLineNumber } else { '?' }
    $sb = New-Object System.Text.StringBuilder
    $prefix = if ($Context) { "ERROR [$Context]" } else { 'ERROR' }
    [void]$sb.AppendLine("$prefix -- $cmd (script line $line)")
    [void]$sb.AppendLine("    Message  : $($ex.Message)")
    [void]$sb.AppendLine("    Type     : $($ex.GetType().FullName)")
    if ($ErrorRecord.CategoryInfo) {
        [void]$sb.AppendLine("    Category : $($ErrorRecord.CategoryInfo.Category) ($($ErrorRecord.CategoryInfo.Reason))")
    }
    if ($ex.InnerException) {
        [void]$sb.AppendLine("    Inner    : $($ex.InnerException.Message)")
    }
    return $sb.ToString().TrimEnd()
}
#endregion

#region EXO cmdlet resolution (moved unchanged -- original lines 588-625) ---
function Get-EXOCommand {
    <#
        Resolve a cmdlet to the Exchange Online copy.

        Originally matched on module NAME -- tmpEXO_* then
        ExchangeOnlineManagement. That worked, but it hard-codes an
        implementation detail of implicit remoting: the generated module name.
        It is not contractual and has changed shape between module versions.

        Now delegates to Get-ExchangeModuleMap, which identifies the two worlds
        by what they EXPORT (on-premises exports Get-RemoteMailbox; the cloud
        does not). That is a property of Exchange rather than of the module
        loader.

        Name and signature kept so existing call sites are untouched. The
        name-based match survives as a fallback for an exotic load that the
        capability test does not recognise.
    #>
    param([Parameter(Mandatory)][string]$Name)

    $map = Get-ExchangeModuleMap
    if (-not $map.Cloud) { $map = Get-ExchangeModuleMap -Refresh }
    if ($map.Cloud) {
        $cmd = Get-Command -Name "$($map.Cloud.Name)\$Name" -ErrorAction SilentlyContinue
        if ($cmd) { return $cmd }
    }

    $all = @(Get-Command -Name $Name -All -ErrorAction SilentlyContinue)
    if ($all.Count -eq 0) {
        throw "No command named '$Name' exists in this session. Is Exchange Online connected (Connect-ExchangeOnline)?"
    }
    $cmd = $all | Where-Object { $_.ModuleName -like 'tmpEXO_*' -or $_.Source -like 'tmpEXO_*' } | Select-Object -First 1
    if (-not $cmd) {
        $cmd = $all | Where-Object { $_.ModuleName -eq 'ExchangeOnlineManagement' -or $_.Source -eq 'ExchangeOnlineManagement' } | Select-Object -First 1
    }
    if (-not $cmd) {
        $sources = ($all | ForEach-Object { "$($_.Name) [ModuleName=$($_.ModuleName); Source=$($_.Source)]" }) -join '  |  '
        throw "Found '$Name' but none of the loaded copies look like Exchange Online: $sources -- run 'Get-Command $Name -All | Select Name,ModuleName,Source' to confirm which one this should be; an on-prem Exchange module/snap-in may be shadowing it."
    }
    return $cmd
}
function Get-ConsoleResizeCommand {
    <#
        The statement that shrinks a launched PowerShell console window.

        Returned as TEXT because it has to run in the child process, not this
        one -- it is prepended to that process's -Command payload.

        Lives here rather than in each launcher. Two launchers open consoles
        (the console's own-process tools and the dashboard's detached tools) and
        two copies of a window size is exactly the shape that let a duplicated
        domain controller drift out of step.

        ORDER MATTERS inside it: WindowSize can never exceed BufferSize, so the
        window shrinks FIRST and the buffer after. The other way round throws.
        Buffer height stays large so scrollback survives.

        try/catch because a host with no real console has no RawUI to set, and
        failing to resize is not a reason to fail to launch.
    #>
    param(
        [int]$Width      = 100,
        [int]$Height     = 25,
        [int]$Scrollback = 3000
    )
    return ('try { $ui = $Host.UI.RawUI; ' +
            "`$w = `$ui.WindowSize; `$w.Width = $Width; `$w.Height = $Height; `$ui.WindowSize = `$w; " +
            "`$b = `$ui.BufferSize; `$b.Width = $Width; `$b.Height = $Scrollback; `$ui.BufferSize = `$b } catch { }")
}

function Resolve-AdObjectType {
    <#
        What kind of recipient is this, derived from Active Directory alone.

        NO msExchRecipientTypeDetails BITMASK TABLE, ON PURPOSE. That number
        would answer this in one line, and the table mapping its values to names
        is not in current Microsoft documentation -- the Exchange recipient
        types page (learn.microsoft.com/exchange/recipients/recipients) lists
        every type by name and gives no numbers at all. Hard-coding a table
        copied from memory or a blog is exactly the kind of thing that is right
        for years and then silently wrong for one recipient type, which is the
        worst failure mode a lookup tool has: a confident wrong label.

        So every branch below reads an attribute whose meaning is documented and
        stable:

          msExchResourceSearchProperties   present on resource mailboxes and
                                           says Room or Equipment outright
          homeMDB                          the mailbox database. Present means
                                           the mailbox is HERE, on-premises.
          msExchRemoteRecipientType        present means the mailbox is in the
                                           cloud and this is the on-premises
                                           stub -- a remote mailbox
          userAccountControl bit 0x2       account disabled. A mailbox on a
                                           disabled account is the definition
                                           of a shared mailbox (see the
                                           Microsoft page above: "The Active
                                           Directory user associated with a
                                           shared mailbox must be a disabled
                                           account")
          groupType bit 0x80000000         security-enabled group
          proxyAddresses                   mail-enabled, whatever it is

        WHAT THIS CANNOT SEE: a room LIST reads as an ordinary distribution
        group here, because what makes it a room list is an Exchange flag with
        no plain AD equivalent. When either Exchange is connected its own
        RecipientTypeDetails is used instead and the distinction comes back.
        This function is the floor, not the ceiling.
    #>
    param([Parameter(Mandatory)]$AdObject)

    $classes  = @($AdObject.objectClass) | ForEach-Object { "$_".ToLowerInvariant() }
    $proxies  = @($AdObject.proxyAddresses) | Where-Object { "$_".Trim() }
    $mailed   = ($proxies.Count -gt 0)

    if ($classes -contains 'contact') {
        if ($mailed) { return 'MailContact' }
        return 'Contact (not mail-enabled)'
    }

    if ($classes -contains 'group') {
        # -band on the raw value. groupType is a signed 32-bit integer in AD and
        # the security bit is the SIGN bit, so the value arrives NEGATIVE for
        # every security group -- comparing it to 2147483648 would never match.
        # Testing the bit works regardless of sign.
        $gt  = 0
        try { $gt = [int]$AdObject.groupType } catch { }
        $sec = [bool]($gt -band 0x80000000)
        if ($sec)    { if ($mailed) { return 'MailUniversalSecurityGroup' }     else { return 'Security group (not mail-enabled)' } }
        if ($mailed) { return 'MailUniversalDistributionGroup' }
        return 'Distribution group (not mail-enabled)'
    }

    # Everything else is a user-class object.
    $res    = (@($AdObject.msExchResourceSearchProperties) -join ' ')
    $remote = "$($AdObject.msExchRemoteRecipientType)".Trim()
    $onHere = [bool]("$($AdObject.homeMDB)".Trim())
    $uac    = 0
    try { $uac = [int]$AdObject.userAccountControl } catch { }
    $disabled = [bool]($uac -band 0x2)

    if ($res -match '(?i)room')      { if ($onHere) { return 'RoomMailbox' }      else { return 'RemoteRoomMailbox' } }
    if ($res -match '(?i)equipment') { if ($onHere) { return 'EquipmentMailbox' } else { return 'RemoteEquipmentMailbox' } }

    if ($onHere) {
        if ($disabled) { return 'SharedMailbox' }
        return 'UserMailbox'
    }
    if ($remote) {
        if ($disabled) { return 'RemoteSharedMailbox' }
        return 'RemoteUserMailbox'
    }
    if ($mailed) { return 'MailUser' }
    if ($disabled) { return 'User (no mailbox, disabled)' }
    return 'User (no mailbox)'
}

function Start-ToolInOnPremProcess {
    <#
        Open a tool in its OWN PowerShell process, connected to on-premises
        Exchange first.

        THE THIRD COPY WAS THE ONE THAT MOVED IT HERE. This logic already
        existed twice -- Start-ToolScriptInNewProcess in AdminConsole.ps1 and
        Start-OnPremToolProcess in Group-Manager.ps1, whose own docstring says
        "the right end state is one copy in Common.ps1 with both callers using
        it". Unified-Lookup needed a third, and three copies of a launch command
        is exactly the shape that let a duplicated domain controller drift out
        of step. So it lives here now.

        THE TWO EXISTING COPIES ARE DELIBERATELY NOT REWIRED YET. Both work,
        both are in daily use, and changing how every tool in the console
        launches is not a change to make in the same edit as a new feature. The
        next person to touch either one should delete it and call this instead.

        WHY A NEW PROCESS AT ALL: the target tools call Exchange cmdlets
        UNQUALIFIED, so on-premises and Exchange Online cannot share a window --
        whichever module was imported last owns the name. The caller's process
        may well have Exchange Online loaded; this one starts clean and connects
        on-premises only.

        -NoExit so the window survives the tool closing. The output is the
        record of what happened, and a window that vanishes takes it with it.
    #>
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [string]$SmtpAddress,
        # Extra switches for the child, e.g. -OnPrem. Passed as separate tokens
        # rather than folded into a string, so the child binds them as switches.
        [string[]]$Extra = @()
    )

    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        throw "$([System.IO.Path]::GetFileName($ScriptPath)) is not in the toolkit folder."
    }

    $psExe = $null
    try { $psExe = (Get-Process -Id $PID).Path } catch { }
    if (-not $psExe) { $psExe = 'powershell.exe' }

    # Every path single-quoted with internal quotes doubled, so the whole
    # -Command payload can be wrapped in double quotes without anything inside
    # fighting it. Start-Process does NOT quote for you.
    $esc = { param([string]$s) "'" + ("$s" -replace "'", "''") + "'" }

    $root  = Split-Path -Parent $ScriptPath
    $parts = New-Object System.Collections.ArrayList
    [void]$parts.Add((Get-ConsoleResizeCommand))

    $connect = Join-Path $root 'Connect-ExSE.ps1'
    if (Test-Path -LiteralPath $connect) { [void]$parts.Add("& $(& $esc $connect)") }

    $call = "& $(& $esc $ScriptPath)"
    if ("$SmtpAddress".Trim()) { $call += " -SmtpAddress $(& $esc $SmtpAddress)" }
    foreach ($x in $Extra) { if ("$x".Trim()) { $call += " $x" } }
    [void]$parts.Add($call)

    $argLine = '-NoExit -ExecutionPolicy Bypass -Command "' + ($parts -join '; ') + '"'
    Start-Process -FilePath $psExe -ArgumentList $argLine | Out-Null
}

function Get-SccModule {
    <#
        Which loaded module is the Security and Compliance session.

        Identified by what it EXPORTS, the same principle Get-ExchangeModuleMap
        uses and for the same reason: the module name generated by implicit
        remoting is an implementation detail that has changed shape between
        versions.

        The marker is a PAIR, not one cmdlet. On-premises Exchange also exports
        Get-ComplianceSearch -- that collision is the entire reason this
        function exists -- but on-premises additionally exports
        Get-RemoteMailbox, which Security and Compliance never does. So the
        test is: exports the compliance cmdlet, does NOT export the
        on-premises one.
    #>
    param([switch]$Refresh)

    if ($Refresh) { $script:SccModule = $null }
    if ($script:SccModule -and (Get-Module -Name $script:SccModule.Name)) {
        return $script:SccModule
    }
    $script:SccModule = $null
    foreach ($m in @(Get-Module)) {
        $ex = @($m.ExportedCommands.Keys)
        if (($ex -contains 'New-ComplianceSearchAction') -and ($ex -notcontains 'Get-RemoteMailbox')) {
            $script:SccModule = $m
            break
        }
    }
    return $script:SccModule
}

function Get-SccCommand {
    <#
        Resolve a cmdlet to the Security and Compliance copy.

        The counterpart to Get-EXOCommand and Get-OnPremCommand. Without it,
        Get-ComplianceSearch in a console that has also connected on-premises
        Exchange means whichever module was imported last -- and the on-premises
        one answers about a completely different service.
    #>
    param([Parameter(Mandatory)][string]$Name)

    $m = Get-SccModule
    if (-not $m) { $m = Get-SccModule -Refresh }
    if ($m) {
        $cmd = Get-Command -Name "$($m.Name)\$Name" -ErrorAction SilentlyContinue
        if ($cmd) { return $cmd }
    }

    $all = @(Get-Command -Name $Name -All -ErrorAction SilentlyContinue)
    if ($all.Count -eq 0) {
        throw "No command named '$Name' exists in this session. Connect Security and Compliance first (Connect S&C)."
    }
    $cmd = $all | Where-Object {
        $_.Module -and (@($_.Module.ExportedCommands.Keys) -notcontains 'Get-RemoteMailbox')
    } | Select-Object -First 1
    if (-not $cmd) {
        throw "'$Name' resolves only to an on-premises Exchange session in this process. That is a different cmdlet against a different service -- connect Security and Compliance first."
    }
    return $cmd
}

function Invoke-SccCommand {
    param(
        [Parameter(Mandatory)][string]$Name,
        [hashtable]$Parameters = @{}
    )
    $cmd = Get-SccCommand -Name $Name
    & $cmd @Parameters
}

function Test-SccConnectionState {
    <#
        Is Security and Compliance usable in this process?

        Resolving the cmdlet is not enough -- an on-premises session or an
        expired remote session both leave a resolvable name behind. The cheap
        call is what separates a live session from its wreckage.
    #>
    param([switch]$Refresh)

    $out = [PSCustomObject]@{
        Connected = $false
        Detail    = 'Security and Compliance is not connected in this process.'
    }
    if ($Refresh) { [void](Get-SccModule -Refresh) }

    $cmd = $null
    try { $cmd = Get-SccCommand -Name 'Get-ComplianceSearch' }
    catch { $out.Detail = "$($_.Exception.Message)"; return $out }

    try {
        # ResultSize 1 because this asks "is the session alive", not "what is
        # in it". Exchange answers a capped query with "There are more results
        # available", which is true and irrelevant -- so the warning is
        # silenced rather than the limit raised.
        & $cmd -ResultSize 1 -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
        $out.Connected = $true
        $out.Detail    = 'Security and Compliance session is live.'
    } catch {
        $out.Detail = "A compliance module is loaded but not usable: $($_.Exception.Message)"
    }
    return $out
}

$script:CmdletParamCommon = @(
    'Verbose', 'Debug', 'ErrorAction', 'WarningAction', 'InformationAction',
    'ProgressAction', 'ErrorVariable', 'WarningVariable', 'InformationVariable',
    'OutVariable', 'OutBuffer', 'PipelineVariable', 'WhatIf', 'Confirm'
)

function Get-EXOWritableParameter {
    <#
        Which properties a Set-* cmdlet will actually accept, asked of the
        cmdlet rather than of a list in this file.

        Transcribing sixty parameter names out of Microsoft Learn into a script
        is a copying exercise with no check on it: a typo produces a control
        that looks editable and throws on save, and a parameter Microsoft adds
        or removes leaves the list quietly wrong. Get-Command already knows,
        knows for the module version actually loaded, and cannot be out of date.

        Returns a hashtable of ParameterName -> [type]. Common parameters are
        removed because they configure PowerShell, not the object, and Identity
        is removed because it selects the target rather than describing it.

        The caller is expected to intersect this with what the matching Get-*
        returned: a parameter with no property to show has nothing to edit, and
        a property with no parameter is read-only. That intersection is the
        entire safety model for these editors.
    #>
    param([Parameter(Mandatory)][string]$Name)

    $out = @{}
    $cmd = $null
    try { $cmd = Get-EXOCommand -Name $Name } catch { return $out }
    if (-not $cmd) { return $out }

    foreach ($k in $cmd.Parameters.Keys) {
        if ($script:CmdletParamCommon -contains $k) { continue }
        if ($k -eq 'Identity') { continue }
        $out[$k] = $cmd.Parameters[$k].ParameterType
    }
    return $out
}

function ConvertTo-CmdletParameterValue {
    <#
        A string from a textbox, converted to what the cmdlet's parameter wants.

        Everything typed into a grid is a string. Passing "5" where an Int32 is
        expected happens to work; passing "true" where a Boolean is expected
        does NOT do what it looks like -- PowerShell casts any non-empty string
        to $true, so "false" would silently ENABLE whatever it was meant to turn
        off. That is the single most dangerous coercion in this whole exercise
        and the reason booleans are handled explicitly rather than left to the
        engine.

        An empty string means "clear it", which the string-valued Set-Place
        parameters document as $null.

        Throws on anything it cannot convert. A rejected edit with a readable
        message beats a converted one that means something else.
    #>
    param([string]$Value, [type]$Type)

    # Nullable<T> unwrapped via GetUnderlyingType rather than by comparing
    # against a generic type literal. The literal form needs a backtick, and a
    # backtick inside a here-string or a badly quoted edit is the kind of
    # character that survives review and breaks at parse time.
    $t = $Type
    if ($t) {
        $under = [System.Nullable]::GetUnderlyingType($t)
        if ($under) { $t = $under }
    }
    $name = if ($t) { $t.Name } else { 'String' }
    $v = "$Value".Trim()

    if ($name -eq 'String[]' -or ($t -and $t.IsArray)) {
        if (-not $v) { return @() }
        return @($v -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }

    if (-not $v) { return $null }

    switch ($name) {
        'Boolean' {
            # Explicit list. NOT [bool]$v -- see above.
            if ($v -match '^(?i)(true|\$true|yes|y|1|enabled)$')  { return $true }
            if ($v -match '^(?i)(false|\$false|no|n|0|disabled)$') { return $false }
            throw "'$Value' is not a yes/no value. Use True or False."
        }
        'Int32' {
            $n = 0
            if (-not [int]::TryParse($v, [ref]$n)) { throw "'$Value' is not a whole number." }
            return $n
        }
        default { return $v }
    }
}

function Invoke-EXOCommand {
    param(
        [Parameter(Mandatory)][string]$Name,
        [hashtable]$Parameters = @{}
    )
    $cmd = Get-EXOCommand -Name $Name
    & $cmd @Parameters
}

function Get-OnPremCommand {
    <#
        The counterpart nobody wrote, which is why the hybrid scripts qualified
        their cloud calls and left their on-premises ones bare.

        An asymmetric guard is a guard with a hole in it: the cloud half was
        pinned down and the on-premises half went wherever import order sent
        it.
    #>
    param([Parameter(Mandatory)][string]$Name)

    $map = Get-ExchangeModuleMap
    if (-not $map.OnPrem) { $map = Get-ExchangeModuleMap -Refresh }
    if (-not $map.OnPrem) {
        throw "'$Name' was requested against on-premises Exchange, which is not connected in this process. Refusing to run it against Exchange Online instead."
    }
    $cmd = Get-Command -Name "$($map.OnPrem.Name)\$Name" -ErrorAction SilentlyContinue
    if (-not $cmd) {
        throw "'$Name' is not exported by the on-premises Exchange module ($($map.OnPrem.Name))."
    }
    return $cmd
}
function Invoke-OnPremCommand {
    param(
        [Parameter(Mandatory)][string]$Name,
        [hashtable]$Parameters = @{}
    )
    $cmd = Get-OnPremCommand -Name $Name
    & $cmd @Parameters
}
#endregion

#region Object-type resolution (moved unchanged -- original lines 628-668) --
function Resolve-Target {
    param(
        [string]$Smtp,
        [string]$ADServer
    )
    $result = [ordered]@{
        Type              = 'NotFound'
        Smtp              = $Smtp
        ADGroupDN         = $null
        ResolutionErrors  = @()
    }

    # THE ALIAS RETRY, AND THE REASON IT IS NEEDED.
    #
    # Get-TargetSmtp completes a bare word by appending the tenant's primary
    # mail domain, so typing 'qateam' becomes 'qateam@primarydomain'. For nearly
    # everything that is right and saves typing.
    #
    # IT IS WRONG FOR ANY OBJECT WHOSE ONLY ADDRESS IS @tenant.onmicrosoft.com
    # -- Microsoft 365 Groups very often are, because nobody gave them a vanity
    # address. The completed address then genuinely does not exist, and the
    # operator sees three stacked "couldn't be found" exceptions for an object
    # that is sitting right there. Typing the full address worked; typing the
    # name did not; and nothing on screen explained the difference.
    #
    # EXCHANGE ALREADY SOLVES THIS. -Identity accepts an alias, a name, a GUID
    # or a distinguished name, not only an SMTP address -- so the bare word the
    # operator typed is itself a valid identity. This retries with it.
    #
    # ONLY WHEN THE COMPLETED ADDRESS FAILED, so nothing changes for the case
    # that already worked, and the extra call is paid only on the way to an
    # error we would otherwise have reported.
    $bare = ''
    if ("$Smtp" -match '^([^@]+)@') { $bare = $Matches[1] }

    $probe = {
        param([string]$Id)
        foreach ($c in @('Get-Mailbox', 'Get-UnifiedGroup', 'Get-DistributionGroup')) {
            try {
                $o = Invoke-EXOCommand -Name $c -Parameters @{ Identity = $Id; ErrorAction = 'Stop' }
                return [PSCustomObject]@{ Cmdlet = $c; Object = $o }
            } catch { }
        }
        return $null
    }

    try {
        Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } | Out-Null
        $result.Type = 'Mailbox'
        return [PSCustomObject]$result
    } catch {
        $result.ResolutionErrors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox (EXO) resolution attempt')
    }
    try {
        Invoke-EXOCommand -Name 'Get-UnifiedGroup' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } | Out-Null
        $result.Type = 'UnifiedGroup'
        return [PSCustomObject]$result
    } catch {
        $result.ResolutionErrors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-UnifiedGroup (EXO) resolution attempt')
    }

    # A CLOUD-NATIVE DISTRIBUTION GROUP WAS NOT A CATEGORY THIS FUNCTION KNEW.
    #
    # It asked Exchange Online for a mailbox, then for a Microsoft 365 group,
    # and then went straight to Active Directory -- so the ONE cmdlet that finds
    # an ordinary distribution list in the cloud was never called. A group like
    # 'AP-3255@...' therefore fell through to the on-premises lookup, which
    # cannot find it either because it has no on-premises object, and Group
    # Editor reported 'did not resolve to a group (Type: NotFound)'.
    #
    # It worked for every SYNCED distribution group, which is why it survived:
    # those do have an AD object, so the third attempt found them and returned
    # OnpremGroup. Every group that had never been on-premises was invisible.
    #
    # ORDER MATTERS -- this sits BEFORE the AD fallback deliberately. Exchange
    # Online is the authority for a converted group even though its old AD
    # object still exists, so asking the cloud first keeps the answer correct
    # through the migration rather than only after the AD objects are cleaned up.
    # CLOUD-MANAGED ONLY. A SYNCED GROUP DELIBERATELY FALLS THROUGH.
    #
    # Exchange Online answers Get-DistributionGroup for BOTH kinds, so
    # returning on the first hit would have classified every synced group as
    # 'DistributionGroup' -- and AD-GroupEditor, which expects 'OnpremGroup' so
    # it knows to use Get-ADGroup and Set-ADGroup, would have stopped resolving
    # the groups it exists to edit. Fixing the cloud case would have broken the
    # on-premises one.
    #
    # IsDirSynced is the same property Group-Manager routes on, read from the same
    # cmdlet, so the two cannot disagree about where a group lives.
    try {
        $dg = Invoke-EXOCommand -Name 'Get-DistributionGroup' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        if (-not $dg.IsDirSynced) {
            $result.Type = 'DistributionGroup'
            return [PSCustomObject]$result
        }
        # Synced: Active Directory is the authority, so keep going and let the
        # on-premises lookup below return OnpremGroup with a real DN.
    } catch {
        $result.ResolutionErrors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-DistributionGroup (EXO) resolution attempt')
    }

    # RETRY WITH THE BARE NAME. See the note at the top of this function.
    #
    # Everything above asked Exchange Online for a completed SMTP address. If
    # that address was invented by Get-TargetSmtp -- the operator typed a name,
    # not an address -- then the object may exist under a different domain, and
    # Exchange will find it by alias.
    #
    # The three errors already collected are DISCARDED on success. Reporting
    # "qateam@primarydomain couldn't be found" three times alongside a result
    # for the object we did find would describe a failed guess rather than an
    # answer.
    if ($bare -and $bare -ne "$Smtp") {
        $hit = & $probe $bare
        if ($hit) {
            $result.ResolutionErrors = @()
            # THE RESOLVED ADDRESS REPLACES THE GUESS. Every caller uses
            # $result.Smtp as the identity for its real work, so returning the
            # invented address would fix the lookup and then fail the edit.
            if ("$($hit.Object.PrimarySmtpAddress)".Trim()) {
                $result.Smtp = "$($hit.Object.PrimarySmtpAddress)"
            }
            switch ($hit.Cmdlet) {
                'Get-Mailbox'      { $result.Type = 'Mailbox';      return [PSCustomObject]$result }
                'Get-UnifiedGroup' { $result.Type = 'UnifiedGroup'; return [PSCustomObject]$result }
                'Get-DistributionGroup' {
                    # Same synced-versus-cloud rule as above, for the same
                    # reason: a synced group must fall through to the AD lookup
                    # so its editor gets a real distinguished name.
                    if (-not $hit.Object.IsDirSynced) {
                        $result.Type = 'DistributionGroup'
                        return [PSCustomObject]$result
                    }
                }
            }
        }
    }

    # The on-premises fallback only runs where there IS an on-premises
    # directory. Without this, a cloud-only tenant looking up an address that
    # is not a mailbox got a Get-ADGroup failure in the output box -- an error
    # about a directory the tenant does not have, for a lookup that had already
    # correctly concluded "not in Exchange Online".
    $tryAd = $true
    try { $tryAd = Test-TenantCapability -Name 'OnPremAD' } catch { $tryAd = $true }
    if (-not $tryAd -or -not (Get-Command -Name 'Get-ADGroup' -ErrorAction SilentlyContinue)) {
        return [PSCustomObject]$result
    }

    # NO SERVER MEANS DO NOT ASK, not ask with nothing.
    #
    # -Server is a mandatory-valued parameter: passing an empty string throws
    # ParameterBindingValidationException, "The argument is null or empty".
    # That is what appeared in Group Editor's error box underneath the two real
    # answers -- a binding fault dressed up as a resolution failure, which reads
    # as though the directory rejected the group rather than never being asked.
    #
    # This became reachable the moment -CloudOk allowed a null ADServer through.
    # It was always latent: any caller that did not supply one hit it.
    if (-not "$ADServer".Trim()) {
        $result.ResolutionErrors += ('The on-premises lookup was not attempted: no domain controller was ' +
            'resolved for this address, which is expected for an object that exists only in the cloud.')
        return [PSCustomObject]$result
    }

    try {
        $grp = @(Get-ADGroup -Filter "mail -eq '$Smtp'" -Server $ADServer `
                    -Properties mail, ManagedBy, GroupCategory, GroupScope, msExchHideFromAddressLists `
                    -ErrorAction Stop)
        if ($grp.Count -eq 1) {
            $result.Type      = 'OnpremGroup'
            $result.ADGroupDN = $grp[0].DistinguishedName
            return [PSCustomObject]$result
        } elseif ($grp.Count -gt 1) {
            $result.ResolutionErrors += "ERROR [Get-ADGroup (on-prem) resolution attempt] -- $($grp.Count) AD groups matched mail '$Smtp'. Expected 0 or 1. Resolve the duplicate mail attribute before using this tool."
        }
    } catch {
        $result.ResolutionErrors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-ADGroup (on-prem) resolution attempt')
    }
    return [PSCustomObject]$result
}
#endregion

#region Inbox rule read (moved unchanged -- original lines 798-806) ---------
function Get-InboxRulesForMailbox {
    param([string]$Smtp)
    try {
        $rules = @(Invoke-EXOCommand -Name 'Get-InboxRule' -Parameters @{ Mailbox = $Smtp; ErrorAction = 'Stop' })
        return [PSCustomObject]@{ Rules = $rules; Error = $null }
    } catch {
        return [PSCustomObject]@{ Rules = @(); Error = (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-InboxRule') }
    }
}
#endregion

#region Completed-tasks logging ----------------------------------------------
#
# Schema v2. The v1 file was Task,DateTime -- one free-text sentence and a
# timestamp to the minute. That records that something happened; it does not let
# you ask what happened to a given mailbox, who did it, or whether it actually
# worked. Everything below exists to answer one of those three questions.
#
# Three deliberate choices:
#
#   Timestamp is ISO 8601 WITH OFFSET. The old format had no seconds, so two
#   actions in the same minute could not be ordered, and no zone, so logs from
#   tenants in different countries could not be merged without guessing. ISO
#   8601 also sorts correctly as plain text, which matters because this is a CSV
#   somebody will open in Notepad.
#
#   Result is its own column. Previously a partial success was recorded as
#   English inside the sentence ("-- 4 applied"), which is invisible to a filter.
#   Failures are logged too: a log that only records successes cannot be used to
#   reconstruct a bad afternoon, which is the one job an audit trail has.
#
#   RunId groups the rows of a single operation. ResourceDuplicator writes one
#   row for the create and another for the settings, and without a shared id
#   there is nothing tying them together -- or tying either to the failure row
#   when the third step dies.
#
# Schema v3 adds Tenant.
#
# Environment was carrying a job it cannot do. It holds a directory Id or a
# cloud token -- a DIRECTORY within one tenant, not the tenant itself. This
# console ships to several of them, every copy writes the same CSV format, and the moment two of
# those files are put side by side to answer "who changed what", nothing in a
# row says which tenant it happened in. Operator and Machine do not settle it
# either: the same engineer on the same jump box administers all six.
#
# It is a separate column rather than something folded into Detail because the
# entire point of the v2 rewrite was that a filter beats a grep.
$script:LogSchemaVersion = 3
$script:LogColumns = @(
    'Timestamp', 'Operator', 'Machine', 'Tenant', 'Script',
    'Action', 'Target', 'Result', 'Detail', 'RunId', 'Environment'
)
$script:LogHeaderLine = ($script:LogColumns -join ',')

# Controlled vocabulary. Free text is kept in Detail; Action exists so that
# "how many mailboxes did we disable last month" is a filter rather than a grep.
$script:LogActionMap = @{
    'add' = 'Add'; 'added' = 'Add'; 'bulk-added' = 'Add'
    'remove' = 'Remove'; 'removed' = 'Remove'; 'bulk-removed' = 'Remove'
    'delete' = 'Delete'; 'deleted' = 'Delete'
    'create' = 'Create'; 'created' = 'Create'
    'enable' = 'Enable'; 'enabled' = 'Enable'
    'disable' = 'Disable'; 'disabled' = 'Disable'
    'grant' = 'Grant'; 'granted' = 'Grant'
    'revoke' = 'Revoke'; 'revoked' = 'Revoke'
    'set' = 'Configure'; 'configured' = 'Configure'; 'edited' = 'Configure'
    'restore' = 'Restore'; 'restored' = 'Restore'
    'released' = 'Release'
    'ran' = 'Run'
}

function Get-LogOperator {
    <#
        DOMAIN\user -- the identity that actually held the token, not a display
        name. On a jump box where several people share a console this is the only
        thing that attributes the change to a person.
    #>
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent().Name
        if ($id) { return $id }
    } catch { }
    if ($env:USERDOMAIN) { return "$($env:USERDOMAIN)\$($env:USERNAME)" }
    return "$($env:USERNAME)"
}

function Get-LogTenantLabel {
    <#
        Which tenant this row belongs to.

        Read from Tenant.json rather than from the connected session, and that
        is deliberate: this has to work on the failure rows too. A row logged
        because Connect-ExO threw has no session to ask, and those are exactly
        the rows worth keeping.

        Name first, primary domain second. Name is what an operator calls the
        tenant and what the console header shows; the domain is unambiguous but
        wordier. Either identifies it -- an empty string does not, so both are
        tried before giving up.
    #>
    try {
        $c = Get-TenantConfig
        if ($c) {
            if ($c.PSObject.Properties['Name'] -and "$($c.Name)".Trim()) {
                return "$($c.Name)".Trim()
            }
            if ($c.PSObject.Properties['PrimaryDomain'] -and "$($c.PrimaryDomain)".Trim()) {
                return "$($c.PrimaryDomain)".Trim().TrimStart('@').ToLower()
            }
        }
    } catch { }
    return ''
}

function Get-LogScriptName {
    <#
        The tool that made the call, not this file. Common.ps1 is dot-sourced, so
        it appears in the call stack for every logged action and has to be walked
        past.
    #>
    try {
        foreach ($frame in (Get-PSCallStack)) {
            if (-not $frame.ScriptName) { continue }
            $leaf = Split-Path -Leaf $frame.ScriptName
            if ($leaf -eq 'Common.ps1') { continue }
            return $leaf
        }
    } catch { }
    return 'unknown'
}

function Get-ToolRunId {
    <#
        One id per tool, for the life of the process.

        Held in a GLOBAL keyed by script name rather than a $script: variable,
        because tools are launched with & inside the console's own process: they
        share globals but each gets its own script scope. A $script: variable
        here would be $null in every tool, and a single global without the key
        would give every tool the same id -- which is worse than none, because it
        would falsely group unrelated actions.
    #>
    param([string]$ScriptName)
    if (-not $global:ToolkitRunIds) { $global:ToolkitRunIds = @{} }
    $key = $ScriptName
    if (-not $key) { $key = 'unknown' }
    if (-not $global:ToolkitRunIds.ContainsKey($key)) {
        $global:ToolkitRunIds[$key] = [guid]::NewGuid().ToString('N').Substring(0, 12)
    }
    return $global:ToolkitRunIds[$key]
}

function ConvertTo-CsvField {
    <#
        RFC 4180 quoting. The v1 writer used a bare -f format string, so a task
        containing a double quote produced a row that no CSV reader could parse
        -- and the corruption was invisible until someone opened the file months
        later. Newlines are flattened for the same reason.
    #>
    param([string]$Value)
    if ($null -eq $Value) { $Value = '' }
    $v = $Value -replace '\r?\n', ' '
    $v = $v -replace '"', '""'
    return '"' + $v + '"'
}

function Split-LegacyTask {
    <#
        Best-effort Action and Target from a v1-style sentence.

        Deliberately conservative: it returns blank rather than a guess. The full
        sentence always survives in Detail, so a blank column costs nothing,
        whereas a wrong one is a filter that quietly returns the wrong answer --
        which is worse than no filter at all.
    #>
    param([string]$Task)
    $out = @{ Action = ''; Target = ''; Environment = '' }
    if (-not $Task) { return $out }

    $first = ($Task.Trim() -split '\s+')[0]
    if ($first) {
        $k = $first.ToLower()
        if ($script:LogActionMap.ContainsKey($k)) { $out.Action = $script:LogActionMap[$k] }
    }

    # An SMTP address is the only token that can be identified as the target
    # without understanding the sentence.
    $m = [regex]::Match($Task, '[\w.+''-]+@[\w-]+\.[\w.-]+')
    if ($m.Success) { $out.Target = $m.Value }

    # The tools already write "(<directory>)" or "(EXO)" into the sentence.
    # That is an existing convention at a known position, so reading it back is
    # recovery rather than guesswork.
    #
    # THE NAMES COME FROM Tenant.json, NOT FROM HERE. This used to spell out
    # one deployment's two directories, which meant the recovery worked only at
    # that deployment and silently returned blank everywhere else -- and blank
    # is indistinguishable from "the sentence did not say", so nobody would
    # have noticed. The cloud tokens are fixed because they name a service
    # rather than a directory.
    $envTokens = @(@(Get-TenantDirectory) | ForEach-Object { "$($_.Id)" } |
                   Where-Object { $_ }) + @('EXO', 'Cloud', 'On-Prem')
    $alt = ($envTokens | Sort-Object -Unique | ForEach-Object { [regex]::Escape($_) }) -join '|'
    if ($alt) {
        $e = [regex]::Match($Task, '\((' + $alt + ')[^);]*')
        if ($e.Success) { $out.Environment = $e.Groups[1].Value.Trim() }
    }

    return $out
}

function Get-CompletedLogPath {
    <#
        Returns the file to append to, or $null when the caller must be warned.

        A v1 file is NOT rewritten or appended to. Export-Csv -Append throws on a
        header mismatch in 5.1, and silently widening someone's existing audit
        file is not a decision a logging function should make on its own -- so v2
        rows go to a clearly named sibling until the old file is moved aside.
    #>
    $folder = Get-LogPath
    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -Path $folder -ItemType Directory -Force | Out-Null
    }

    # Candidates in order. The fallback name is KEYED TO THE SCHEMA VERSION
    # rather than hardcoded.
    #
    # It used to be a literal 'CompletedTasks-v2.csv', which worked exactly once.
    # Adding the Tenant column made v2 stale, and a deployment that had already
    # been pushed onto the v2 fallback would have had v3 rows appended under a v2
    # header -- every column after Machine shifted by one, in a file whose name
    # still claimed to be v2. Silent, and only discovered by someone reading the
    # log months later to answer a question that matters.
    $names = @('CompletedTasks.csv', "CompletedTasks-v$($script:LogSchemaVersion).csv")

    foreach ($name in $names) {
        $path = Join-Path $folder $name
        if (-not (Test-Path -LiteralPath $path)) {
            $script:LogHeaderLine | Out-File -FilePath $path -Encoding utf8
            if ($name -ne $names[0]) {
                Write-Warning ("$($names[0]) uses an older column set. Writing to " +
                    "$name instead. Archive the old file and the original name " +
                    "will be used again.")
            }
            return $path
        }
        try { $head = (Get-Content -LiteralPath $path -TotalCount 1 -ErrorAction Stop) } catch { $head = '' }
        if ("$head".Trim() -eq $script:LogHeaderLine) { return $path }
    }

    # Both names exist and neither matches -- only reachable if a versioned file
    # was hand-edited, since its name already encodes the schema. Date-stamp a
    # sibling rather than return $null: losing the row is the one outcome an
    # audit log cannot have, and a confusing filename is recoverable where a
    # missing entry is not.
    $stamped = Join-Path $folder ("CompletedTasks-v$($script:LogSchemaVersion)-" +
                                  (Get-Date -Format 'yyyyMMdd') + '.csv')
    if (-not (Test-Path -LiteralPath $stamped)) {
        $script:LogHeaderLine | Out-File -FilePath $stamped -Encoding utf8
        Write-Warning ("Existing log files do not match the current columns. " +
            "Writing to $(Split-Path -Leaf $stamped).")
    }
    return $stamped
}

function Get-LogTarget {
    <#
        THE VALUE THAT GOES IN THE AUDIT LOG'S Target COLUMN.

        THE RULE: an address, not a name. Pass the candidates in preference
        order and the first non-empty one wins -- in practice
        Get-LogTarget $row.PrimarySmtpAddress $row.Name.

        WHY THIS EXISTS. A row reading

            Target : Armbruster, Glenn W.
            Detail : ForwardingSmtpAddress: 'x@y.com' -> '(cleared)'

        cannot be acted on. A person with three mailboxes has one display name
        across all of them, so the log records that SOMETHING was changed and
        loses which. An address is unique, is what every cmdlet takes as
        -Identity, and is what the next person will search the log for.

        A DISPLAY NAME IS ALSO NOT STABLE. Renaming somebody rewrites their
        display name and leaves every historical audit row pointing at a string
        nothing answers to any more.

        THE FALLBACK IS DELIBERATE, NOT LAZY. Some things genuinely have no
        address -- a plain security group, a room list at the moment before
        Exchange assigns one, a failed create. A name is a worse identifier than
        an address and a far better one than blank, so the chain degrades rather
        than refusing.

        '(unidentified)' IS RETURNED RATHER THAN AN EMPTY STRING when every
        candidate is blank. An empty Target column reads as a logging bug and
        sends the reader looking for one; the literal word says the tool knew it
        had nothing and wrote that down.
    #>
    param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Candidate)

    foreach ($c in @($Candidate)) {
        if ("$c".Trim()) { return "$c".Trim() }
    }
    return '(unidentified)'
}

function Write-CompletedLog {
    <#
        .SYNOPSIS
        Appends one row to the completed-tasks audit log.

        .DESCRIPTION
        -Task is the v1 parameter and still works: it lands in Detail, and Action
        and Target are derived from it where that can be done safely. New callers
        should pass -Action, -Target and -Result explicitly, which is the only way
        Result can be anything other than Success.

        -Target IS AN ADDRESS, NOT A DISPLAY NAME, whenever the object has one.
        Use Get-LogTarget to express that -- it takes candidates in preference
        order and its docstring explains why the rule exists. Test-Toolkit rule
        TGT001 flags call sites that pass a name where an address was available.
    #>
    param(
        [string]$Task,
        [string]$Action,
        [string]$Target,
        [ValidateSet('Success', 'Partial', 'Failed')][string]$Result = 'Success',
        [string]$Detail,
        [string]$Environment,
        [string]$RunId,

        # Almost never passed. It exists for the tools that act ACROSS tenants
        # -- a migration script reading one and writing another -- where the
        # config file this console was launched from is not the tenant the row
        # describes.
        [string]$Tenant
    )

    if (-not $Detail) { $Detail = $Task }
    if (-not $Detail -or $Detail.Trim() -eq '') {
        if (-not $Action -and -not $Target) { return }
    }

    if (-not $Action -or -not $Target -or -not $Environment) {
        $parsed = Split-LegacyTask -Task $Detail
        if (-not $Action)      { $Action      = $parsed.Action }
        if (-not $Target)      { $Target      = $parsed.Target }
        if (-not $Environment) { $Environment = $parsed.Environment }
    }

    $scriptName = Get-LogScriptName
    if (-not $RunId) { $RunId = Get-ToolRunId -ScriptName $scriptName }
    if (-not $Environment) { $Environment = Get-LogEnvironmentLabel }

    # Round-tripping through DateTimeOffset keeps the machine's real UTC offset,
    # including whichever side of a DST change today falls on.
    $ts = [DateTimeOffset]::Now.ToString('yyyy-MM-ddTHH:mm:sszzz')

    if (-not $Tenant) { $Tenant = Get-LogTenantLabel }

    $values = @(
        $ts, (Get-LogOperator), $env:COMPUTERNAME, $Tenant, $scriptName,
        $Action, $Target, $Result, "$Detail".Trim(), $RunId, $Environment
    )
    $line = (($values | ForEach-Object { ConvertTo-CsvField -Value $_ }) -join ',')

    $file = Get-CompletedLogPath
    if (-not $file) { return }
    try { Add-Content -Path $file -Value $line -ErrorAction Stop }
    catch { Write-Warning "Could not write audit log: $($_.Exception.Message)" }
}

function Write-FailedLog {
    <#
        Convenience wrapper so an error path is one line. Without this the catch
        blocks stay empty, which is exactly how a log ends up recording only
        successes.
    #>
    param(
        [string]$Action,
        [string]$Target,
        [string]$Detail,
        [System.Management.Automation.ErrorRecord]$ErrorRecord,
        # Forwarded when the caller knows them. Added 2026-08-13: the failure
        # row was landing with a blank Tenant/Environment and no RunId while the
        # success row beside it had both, so the rows you most need to correlate
        # were the ones you could not. Optional, so no existing caller changes.
        [string]$Environment,
        [string]$RunId
    )
    $d = $Detail
    if ($ErrorRecord) {
        $msg = $ErrorRecord.Exception.Message
        if ($d) { $d = "$d -- $msg" } else { $d = $msg }
    }
    $p = @{ Action = $Action; Target = $Target; Result = 'Failed'; Detail = $d }
    if ("$Environment".Trim()) { $p['Environment'] = $Environment }
    if ("$RunId".Trim())       { $p['RunId']       = $RunId }
    Write-CompletedLog @p
}
function Get-LogEnvironmentLabel {
    <#
        A fallback only. Environment cannot be derived reliably at log time --
        the operator's own forest is not necessarily the target's, and this
        console reaches both. Callers that know should pass -Environment; the
        legacy parser recovers it from the "(<directory>)" convention the
        tools already follow. This is what is left when neither applies.
    #>
    if ($global:ToolkitEnvironment) { return "$($global:ToolkitEnvironment)" }
    if ($env:USERDNSDOMAIN) { return $env:USERDNSDOMAIN }
    return ''
}
#endregion

#region Directory resolution ---------------------------------------------------
function Get-ADServerString {
    <#
        The server to bind to for a directory.

        Was: a ValidateSet naming two directories -- one organisation's
        internal structure written into a type constraint, in a function 35
        scripts call. Any tenant without those names could not use it, and one
        with three directories could not be expressed at all.

        An explicit -DomainController still wins, because a caller that has
        already resolved a server should not have it second-guessed.
    #>
    param(
        [string]$Environment,
        [string]$DomainController
    )
    if ($DomainController) { return $DomainController }
    if ($Environment)      { return (Get-TenantDirectoryServer -Id $Environment) }
    return (Get-TenantDirectoryServer)
}

function Find-ADObjectDirectory {
    <#
        Which configured directory holds this object.

        ONE loop over Get-TenantDirectory, replacing two hand-written blocks
        that named the directories inline. That duplication is why adding a
        third would have meant editing the function rather than the config --
        and why the two blocks had already drifted, one passing a pinned DC and
        the other not.

        Returns the FIRST match. Directory order in Tenant.json is therefore
        search order, which is worth knowing if an account exists in two.
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('Mail','Sam')][string]$By,
        [Parameter(Mandatory)][string]$Value
    )
    $errors = @()
    $val    = "$Value".Trim()

    # Two gates before touching AD at all.
    #
    # A cloud-only tenant has no on-premises directory to search, and a machine
    # without RSAT has no Get-ADObject to search it with. Neither is a fault,
    # and neither should surface as an exception from a lookup the operator did
    # not know was happening -- 'Out of Office' failing with an AD type-loading
    # error tells them nothing about either.
    $hasAd = $true
    try { $hasAd = Test-TenantCapability -Name 'OnPremAD' } catch { $hasAd = $true }
    if (-not $hasAd) {
        return [PSCustomObject]@{
            Environment = $null; ADServer = $null; DCParam = $null
            Object      = $null
            Errors      = @()   # not an error: this tenant has no on-premises directory
        }
    }
    if (-not (Get-Command -Name 'Get-ADObject' -ErrorAction SilentlyContinue)) {
        return [PSCustomObject]@{
            Environment = $null; ADServer = $null; DCParam = $null
            Object      = $null
            Errors      = @('Tenant.json says this tenant has on-premises AD, but the ActiveDirectory module is not available on this machine. Install RSAT, or switch OnPremAD off in TenantManager.')
        }
    }

    foreach ($d in @(Get-TenantDirectory)) {
        $srv = Get-TenantDirectoryServer -Id $d.Id
        if (-not $srv) { continue }
        try {
            $hit = if ($By -eq 'Mail') {
                @(Get-ADObject -Filter "mail -eq '$val'" -Server $srv -ErrorAction Stop)
            } else {
                @(Get-ADUser -Identity $val -Server $srv -ErrorAction Stop)
            }
            if (@($hit).Count -gt 0) {
                return [PSCustomObject]@{
                    Environment = "$($d.Id)"
                    ADServer    = $srv
                    DCParam     = $(if (@($d.DomainControllers).Count -eq 1) { $srv } else { '' })
                    Object      = @($hit)[0]
                    Errors      = $errors
                }
            }
        }
        catch {
            # MATCHED BY NAME, NOT BY TYPE.
            #
            # This was 'catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]'.
            # A typed catch makes PowerShell RESOLVE that type when an exception
            # occurs, and the type only exists once the ActiveDirectory module
            # is loaded. On a machine without RSAT the catch itself threw
            # "Unable to find type" -- and that replaced the real error, so
            # every AD failure reported a cause that was not the cause.
            #
            # A string comparison needs nothing loaded and cannot fail that way.
            $typeName = ''
            try { $typeName = $_.Exception.GetType().FullName } catch { }
            if ($typeName -match 'ADIdentityNotFoundException') {
                # Not in this directory. Expected while searching, not an error.
                continue
            }
            $errors += (Format-ErrorDetail -ErrorRecord $_ -Context "$($d.Label) lookup ($By)")
        }
    }
    return [PSCustomObject]@{
        Environment = $null; ADServer = $null; DCParam = $null
        Object      = $null; Errors   = $errors
    }
}

function Find-TargetEnvironment {
    # Kept as a name because 35 scripts call it. Now a thin wrapper.
    param([string]$Smtp)
    $r = Find-ADObjectDirectory -By 'Mail' -Value $Smtp
    return [PSCustomObject]@{ Environment = $r.Environment; ADServer = $r.ADServer; Errors = $r.Errors }
}

function Find-UserEnvironmentBySam {
    param([Parameter(Mandatory)][string]$SamAccountName)
    $r = Find-ADObjectDirectory -By 'Sam' -Value $SamAccountName
    return [PSCustomObject]@{
        Environment = $r.Environment; ADServer = $r.ADServer; DCParam = $r.DCParam
        User = $r.Object; Errors = $r.Errors
    }
}
#endregion

#region NEW -- consistent window placement -----------------------------------
function Test-ConsoleLaunch {
    <#
        Was this tool started by the console, or typed at a prompt?

        WHAT THIS IS FOR, PLAINLY: it keeps honest people honest. A service desk
        engineer who opens the folder, double-clicks a script and sees "This
        standalone script is unsupported" stops. That is the entire goal, and
        it is worth having -- a tool run on its own has no looked-up object, no
        directory detection and no connection guard, so it fails somewhere
        confusing instead of at the front door.

        IT IS NOT A SECURITY CONTROL AND MUST NEVER BE DESCRIBED AS ONE. The
        script is plain text in a folder the reader can write to. Anyone who
        wants past this can read the check, set the variable, or delete four
        lines. It is a locked front door: it turns back the people who were
        never going to break a window.

        WHICH IS WHY THE MESSAGE SAYS "unsupported" AND NOT "unauthorized".
        Unauthorized is a claim about identity, and this is not checking
        identity -- it is reading an environment variable. That word would
        invite a ticket asking for access that does not exist, and if anyone
        ever found the variable, the claim would have been a lie. Everything
        else in this toolkit says exactly what it means.

        The message also deliberately does NOT mention the console, Tool
        Manager, or how to add a script. Telling someone the door is locked is
        enough; drawing them a map of the window is not.

        HOW IT WORKS. AdminConsole sets ToolkitConsoleSession to a per-run GUID
        before launching anything. Child processes inherit environment
        variables, so this covers BOTH launch paths -- in-process with the call
        operator, and detached through Start-Process for tools that need their
        own window. A prompt opened separately has no such variable.

        THE ENGINEER'S BYPASS is ToolkitAllowDirect. Set it once in your
        PowerShell profile and every tool runs directly again, which is
        necessary because the person maintaining this runs scripts by hand all
        day. It is deliberately undocumented outside this comment.

        Returns $true to proceed. The caller does:

            if (-not (Test-ConsoleLaunch)) { return }

        RETURN, not exit -- so a tool refusing to start cannot take the
        operator's PowerShell window down with it.
    #>
    param(
        # For the handful of scripts that legitimately run from a prompt:
        # Unblock (needed before the console works at all), Publish-Toolkit and
        # Deploy (master only), Test-Toolkit (reads everything, touches
        # nothing). Those simply never call this.
        [string]$ToolName = 'This'
    )

    if ("$env:ToolkitAllowDirect".Trim())    { return $true }
    if ("$env:ToolkitConsoleSession".Trim()) { return $true }

    $msg = 'This standalone script is unsupported.'
    try {
        # Load WinForms if the caller has not. Six tools dot-source Common
        # BEFORE their Add-Type lines, and for those the box would otherwise
        # fall through to Write-Host -- into a window that closes immediately,
        # which is the same as saying nothing at all. The guard has to be seen.
        if (-not ('System.Windows.Forms.MessageBox' -as [type])) {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        }
        [void][System.Windows.Forms.MessageBox]::Show(
            $msg, 'Unsupported',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
    } catch {
        # No WinForms loaded, or no desktop. Say it plainly instead -- a guard
        # that silently does nothing because it could not draw a box would let
        # through exactly the case it exists to stop.
        Write-Host ''
        Write-Host $msg -ForegroundColor Yellow
        Write-Host ''
    }
    return $false
}

function Set-FormResizable {
    <#
        Let the operator resize a window, without letting them break it.

        THREE SETTINGS, AND THE THIRD IS THE ONE THAT MATTERS.

          FormBorderStyle Sizable   a grab handle on the edges
          MaximizeBox     $true     a maximise button, since it now means something
          MinimumSize     = Size    the floor

        MinimumSize is the whole point. Every window in this toolkit is laid out
        with absolute coordinates at a designed size; shrink it below that and
        controls do not reflow, they get covered by the window edge. So the
        designed size becomes the FLOOR and the operator can only grow it --
        which is the only direction anybody actually wants. Nobody has ever
        wanted a smaller Message Trace.

        CALL THIS LAST, after the size is final. It reads $Form.Size to set the
        floor, so calling it before a later ClientSize assignment would pin the
        minimum to a size the window no longer has. Next to Set-FormScreenCorner
        is the right place.

        RESIZING IS ONLY USEFUL IF THE CONTROLS FOLLOW. A Sizable form whose
        controls have no Anchor just reveals grey space when dragged, which
        looks more broken than a fixed window does. 49 of the 69 forms here are
        Sizable by default and have no anchors -- that is the state this
        function exists to stop being the default outcome. Anchoring stays
        per-form because only the author knows which control should absorb the
        space; the convention is Top,Left,Right for a full-width control and
        Top,Left,Right,Bottom for the one output box that should soak up the
        rest.
    #>
    param(
        [Parameter(Mandatory)]$Form,
        # For a window that genuinely can go smaller -- none today, but a
        # future one with a docked layout could.
        [switch]$AllowShrink
    )
    $Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $Form.MaximizeBox     = $true
    $Form.MinimizeBox     = $true
    if (-not $AllowShrink) { $Form.MinimumSize = $Form.Size }
}

function Add-ToolTargetBar {
    <#
        Puts the "who are we looking at" box ON the tool's own form, and shifts
        everything already there down to make room.

        WHY THIS EXISTS. Resolve-ToolTarget pops a modal dialog when no address
        was supplied, so opening a tool meant answering a question in one window
        and reading the answer in another -- and looking up a SECOND person meant
        closing the tool and starting again. That is the wrong shape for work
        that is naturally "check these four people".

        CALL IT AFTER EVERY CONTROL IS ADDED AND BEFORE Add-ToolHelpButton.
        It shifts existing controls, so anything added afterwards would be left
        behind at the old offset; and the help button positions itself from
        ClientSize at run time, so it must be placed after the form has grown.
        Getting this order wrong is not subtle -- the '?' lands in the middle of
        the window.

        ANCHORS ARE TURNED OFF FOR THE DURATION, and that is load-bearing.

        A bottom-anchored control remembers its distance from the bottom edge,
        and that distance is re-measured the moment its bounds change. Shifting
        it down 40px re-measures it as 40px CLOSER to the bottom; growing the
        form by 40 then moves it down another 40 to restore that distance. It
        travels 80px, and the button row ends up below the client area --
        present, wired, invisible, and unreachable by resizing, because the
        anchor drags it further down as the form shrinks.

        SuspendLayout does not prevent this and an earlier version of this
        comment claimed it did. Suspending batches the LAYOUT PASS; it does not
        stop a control from re-measuring its own anchor when its Top is
        assigned. The only reliable fix is to drop each control to Top,Left
        while the geometry moves and restore the real anchor afterwards, which
        re-measures once, from the finished positions.

        IT RETURNS THE CONTROLS. IT DOES NOT TAKE A CALLBACK, and that is a
        deliberate reversal.

        The first version took an -OnLookup scriptblock. That would have worked
        for a one-line body and failed for the real ones: a tool's refresh
        routine is a function nested INSIDE its form function, and a scriptblock
        handed across into Common.ps1 carries its own session state rather than
        that local scope. The existing handlers in these tools -- 
        $btnRefresh.Add_Click({ Update-HealthFromCurrent }) -- work precisely
        because they are wired in the scope where that function is defined.

        So the caller wires the click, in its own scope, on equal footing with
        every other button it already owns. This function only does the part
        that is genuinely shared: the geometry.
    #>
    param(
        [Parameter(Mandatory)]$Form,
        [string]$Label        = 'Mailbox or address:',
        [string]$InitialValue = '',
        [int]$LabelWidth      = 150
    )

    $barH = 40

    # Captured BEFORE anything moves. $Form.Controls is live, so the list is
    # taken first and the bar's own controls are added after the loop.
    $topLeft = [System.Windows.Forms.AnchorStyles]::Top -bor
               [System.Windows.Forms.AnchorStyles]::Left
    $saved = @()
    foreach ($c in $Form.Controls) {
        $saved += [PSCustomObject]@{ Control = $c; Anchor = $c.Anchor }
        $c.Anchor = $topLeft
    }

    $Form.SuspendLayout()
    try {
        foreach ($e in $saved) { $e.Control.Top = $e.Control.Top + $barH }
        $Form.ClientSize = New-Object System.Drawing.Size(
            $Form.ClientSize.Width, ($Form.ClientSize.Height + $barH))
    } finally {
        $Form.ResumeLayout()
    }

    # Restored last. Setting Anchor measures the distance to the edge from the
    # control's CURRENT bounds, so each one is re-anchored against the layout
    # it actually ended up in.
    foreach ($e in $saved) { $e.Control.Anchor = $e.Anchor }

    $lbl          = New-Object System.Windows.Forms.Label
    $lbl.Text     = $Label
    $lbl.Location = New-Object System.Drawing.Point(12, 13)
    $lbl.AutoSize = $true
    $Form.Controls.Add($lbl)

    $txt          = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point((12 + $LabelWidth), 10)
    $txt.Font     = New-Object System.Drawing.Font('Consolas', 9)
    $txt.Anchor   = 'Top,Left,Right'
    $txt.Text     = "$InitialValue"
    $Form.Controls.Add($txt)

    $btn          = New-Object System.Windows.Forms.Button
    $btn.Text     = 'Look Up'
    $btn.Size     = New-Object System.Drawing.Size(100, 24)
    $btn.Anchor   = 'Top,Right'
    $Form.Controls.Add($btn)

    # WIDTHS ARE COMPUTED FROM THE FORM, not hardcoded.
    #
    # These tools range from 560px to 960px wide, so a fixed box is either
    # cramped or leaves a gap. The 52px on the right is the help button's
    # reserved corner -- Add-ToolHelpButton sits at ClientSize-40, and a Look Up
    # button that reaches further overlaps it. That collision has already
    # happened once, in Group-Manager.
    $rightEdge  = $Form.ClientSize.Width - 52
    $btn.Left   = $rightEdge - $btn.Width
    $btn.Top    = 10
    $txt.Width  = [Math]::Max(120, ($btn.Left - 10) - $txt.Left)

    $bar = [PSCustomObject]@{
        Label   = $lbl
        TextBox = $txt
        Button  = $btn
    }

    # Enter looks up. Typing an address and pressing Enter is what everyone
    # does, and without this it does nothing at all.
    $Form.AcceptButton = $btn

    return $bar
}

function Set-FormScreenCorner {
    param(
        [Parameter(Mandatory)]$Form,
        [ValidateSet('TopLeft', 'BottomLeft')][string]$Corner = 'TopLeft'
    )
    try {
        $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual

        # SHRINK BEFORE PINNING. A form taller than the working area cannot be
        # placed with both its edges on screen, and the clamp below chooses the
        # TOP edge -- so whatever falls off the bottom, which is always the
        # button row, silently disappears. There is no scrollbar and no clue:
        # the window just looks like it was built without a Close button.
        #
        # Clamping the SIZE first keeps that row reachable. Controls anchored to
        # the bottom follow the new edge; a form whose controls are not anchored
        # gets AutoScroll, so the row is one scroll away rather than gone.
        # Either beats off-screen.
        if ($Form.Height -gt $wa.Height -or $Form.Width -gt $wa.Width) {
            $Form.AutoScroll = $true
            $Form.Size = New-Object System.Drawing.Size(
                [Math]::Min($Form.Width,  $wa.Width),
                [Math]::Min($Form.Height, $wa.Height))
        }

        $x = $wa.Left
        if ($Corner -eq 'TopLeft') {
            $y = $wa.Top
        } else {
            $y = $wa.Bottom - $Form.Height
            if ($y -lt $wa.Top) { $y = $wa.Top }
        }
        $Form.Location = New-Object System.Drawing.Point($x, $y)
    } catch {
        $Form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    }
}
#endregion

#region Theming --------------------------------------------------------------
# Deliberately NOT $script:-scoped state for the path.
#
# Set-ConsoleTheme is called from tool scripts that AdminConsole launches with
# the call operator. Those run in their OWN script scope, so a $script:
# variable assigned when Common.ps1 was dot-sourced into the console is $null
# by the time the function runs inside the child -- which fed $null to
# Test-Path -LiteralPath and threw. $PSScriptRoot inside a function always
# resolves to the file that DEFINED the function, so it is scope-proof.
$script:ThemeCache = $null

function Get-ThemeFilePath {
    return (Get-ConfigFilePath -Name 'Themes.json')
}

function ConvertTo-ThemeColor {
    param([string]$Hex)
    if ([string]::IsNullOrWhiteSpace($Hex)) { return $null }
    try { return [System.Drawing.ColorTranslator]::FromHtml($Hex) } catch { return $null }
}

function Get-ConsoleTheme {
    <#
        The resolved theme is cached in a GLOBAL, not a $script: variable.

        Tool scripts are launched by the console with the call operator, so
        they run in their own script scope and see none of the console's
        $script: state -- including the tenant-domain cache this resolution
        depends on. When the tenant lookup came up empty in a child, $Name fell
        back to ActiveTheme, which is normally the disabled 'Default', and the
        tool silently opened unthemed. A global is inherited by children, so
        the console's answer travels with them.

        It also means Get-AcceptedDomain is called once per session rather than
        once per tool launch.
    #>
    param([string]$Name, [switch]$Refresh)

    if ($Refresh) {
        $script:ThemeCache = $null
        $global:AdminConsoleTheme = $null
    }
    if (-not $Name -and $global:AdminConsoleTheme) { return $global:AdminConsoleTheme }
    if (-not $script:ThemeCache) {
        $themeFile = Get-ThemeFilePath
        if (-not $themeFile) { return $null }
        if (-not (Test-Path -LiteralPath $themeFile)) { return $null }
        try {
            $script:ThemeCache = Get-Content -LiteralPath $themeFile -Raw |
                                 ConvertFrom-Json
        } catch { return $null }
    }
    $cfg = $script:ThemeCache

    if (-not $Name) {
        $Name = $cfg.ActiveTheme
        if ($cfg.FollowTenant) {
            $d = Get-ConnectedTenantDomain
            if ($d -and $cfg.TenantThemes.PSObject.Properties.Name -contains $d) {
                $Name = $cfg.TenantThemes.$d
            }
        }
    }
    if (-not $Name) { return $null }
    if ($cfg.Themes.PSObject.Properties.Name -notcontains $Name) { return $null }

    $t = $cfg.Themes.$Name
    if (-not $t.Enabled) { return $null }
    Add-Member -InputObject $t -NotePropertyName 'Name' -NotePropertyValue $Name -Force
    $global:AdminConsoleTheme = $t
    return $t
}

function Get-ShadedColor {
    <#
        Nudges a colour lighter or darker by a fixed amount, staying in range.

        Direction is chosen from the base colour's own brightness: a dark theme
        gets a LIGHTER stripe, a light theme gets a darker one. A fixed
        direction would produce an invisible stripe on half the themes, which is
        the same as having no stripe at all.
    #>
    param(
        [Parameter(Mandatory)][System.Drawing.Color]$Base,
        [int]$Amount = 12
    )
    $lift = (($Base.R + $Base.G + $Base.B) / 3) -lt 128
    $d = if ($lift) { $Amount } else { -$Amount }
    $clamp = { param($v) [Math]::Max(0, [Math]::Min(255, [int]$v)) }
    return [System.Drawing.Color]::FromArgb(
        (& $clamp ($Base.R + $d)), (& $clamp ($Base.G + $d)), (& $clamp ($Base.B + $d)))
}

function Get-BlendedColor {
    <#
        Mixes two colours. Amount is how far to move from -From toward -To,
        0.0 to 1.0. Used to dim read-only text toward its own background, which
        reads as deliberately inactive at any brightness -- unlike a hardcoded
        grey, which disappears on a dark theme and shouts on a light one.
    #>
    param(
        [Parameter(Mandatory)][System.Drawing.Color]$From,
        [Parameter(Mandatory)][System.Drawing.Color]$To,
        [double]$Amount = 0.5
    )
    if ($Amount -lt 0) { $Amount = 0 } elseif ($Amount -gt 1) { $Amount = 1 }
    $mix = { param($a, $b) [int][Math]::Round(($a * (1 - $Amount)) + ($b * $Amount)) }
    return [System.Drawing.Color]::FromArgb(
        (& $mix $From.R $To.R), (& $mix $From.G $To.G), (& $mix $From.B $To.B))
}

function Set-FormDoubleBuffered {
    param([Parameter(Mandatory)]$Form)
    try {
        [System.Windows.Forms.Control].GetProperty('DoubleBuffered',
            'Instance,NonPublic').SetValue($Form, $true, $null)
    } catch { }
}

function Add-FormGradient {
    param([Parameter(Mandatory)]$Form, [string]$FromHex, [string]$ToHex)
    $c1 = ConvertTo-ThemeColor $FromHex
    $c2 = ConvertTo-ThemeColor $ToHex
    if (-not $c1 -or -not $c2) { return }
    Set-FormDoubleBuffered -Form $Form
    $Form.Add_Paint({
        param($src, $e)
        try {
            $r = $src.ClientRectangle
            if ($r.Width -le 0 -or $r.Height -le 0) { return }
            $b = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
                    $r, $c1, $c2,
                    [System.Drawing.Drawing2D.LinearGradientMode]::Vertical)
            $e.Graphics.FillRectangle($b, $r)
            $b.Dispose()
        } catch { }
    }.GetNewClosure())
}

# Console button geometry, shared rather than written as literals in each
# consumer. ToolManager's Sections dialog previews the size the console will
# actually draw, and a preview computed from a different pad than the console
# uses is a preview that lies -- which is worse than no preview, because it
# will be believed.
#
# 38 is not arbitrary. It was 32 until a two-line caption clipped along its top
# edge; it is two lines of the console font plus the button's own chrome.
$script:ConsoleBtnHeight = 38
$script:ConsoleBtnPad    = 4

# Spare width carried by every section panel, split evenly across both edges so
# it reads as a margin instead of collecting as a gap on the right. Even is
# deliberate -- an odd value cannot be halved without one side being a pixel
# wider than the other, which is exactly the asymmetry this exists to remove.
$script:ConsolePanelSlack = 6

function Get-ConsoleButtonWidth {
    <#
        The narrowest button width at which every caption still fits.

        Button width was a flat 150 for as long as the console has existed, and
        150 is not a measurement -- it is a number that happened to work at
        100% scaling in the font this console was written on. It has no
        relationship to the captions actually in Tools.json, so it is wrong in
        both directions at once: too wide for a deployment of short names, and
        too narrow the moment someone adds a long one.

        MEASURING RATHER THAN GUESSING also fixes display scaling for free. At
        125% or 150% the form's font grows but a hardcoded 150 does not, so
        captions that fit on the machine the toolkit was built on clip on a
        laptop. MeasureText is handed the REAL font, so whatever the shell is
        scaled to is already priced in.

        WHY THIS IS NOT "widest caption on one line". Some captions are meant
        to wrap -- 'Mailbox / Onedrive Perms' has always been two lines, and
        forcing it onto one would set the width for all forty-odd buttons and
        make the window considerably WIDER, which is the opposite of the point.
        So the test is not "does it fit on a line", it is "does it fit in the
        BOX", wrapping included.

        MaxLines is DERIVED from ButtonHeight rather than assumed, because a
        line of 9pt Microsoft Sans Serif is 15px at 100% scaling and 16px at
        certain DPI settings. Hardcoding "30px of text area" would silently
        mean two lines on one machine and one line on another -- and the
        one-line machine would not error, it would just quietly return a much
        larger width and produce a much wider console. A layout that changes
        shape depending on whose screen it is on is exactly the failure this
        function exists to remove.
    #>
    param(
        [string[]]$Captions,
        [System.Drawing.Font]$Font,
        [int]$ButtonHeight = 38,
        [int]$Minimum = 96,
        [int]$Maximum = 240,

        # Chrome between the button edge and its text. A themed button is
        # FlatStyle::Flat (1px border) and an unthemed one is Standard (3px
        # plus a focus rectangle), so this is sized for the WIDER of the two.
        # Being a few pixels generous costs a few pixels of window; being a few
        # pixels short clips a caption, and a clipped caption in a demo is not
        # recoverable by explaining it.
        [int]$Inset = 12,

        # 0 means derive from ButtonHeight.
        [int]$MaxLines = 0
    )

    $list = @($Captions | Where-Object { "$_".Trim() })
    if ($list.Count -eq 0) { return 150 }

    if (-not $Font) {
        $Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)
    }

    try {
        $flags = [System.Windows.Forms.TextFormatFlags]::WordBreak
        $unbounded = New-Object System.Drawing.Size(4000, 4000)

        # 'Hg' has both an ascender and a descender, so it measures the full
        # line box rather than the height of whichever letters happen to be in
        # a tool name.
        $lineH = [System.Windows.Forms.TextRenderer]::MeasureText(
                    'Hg', $Font, $unbounded, $flags).Height
        if ($lineH -lt 6) { $lineH = 15 }

        if ($MaxLines -lt 1) {
            $MaxLines = [int][Math]::Floor(($ButtonHeight - 6) / $lineH)
        }
        if ($MaxLines -lt 1) { $MaxLines = 1 }
        $allowedH = $lineH * $MaxLines

        for ($w = $Minimum; $w -le $Maximum; $w += 2) {
            $textW = $w - $Inset
            if ($textW -lt 24) { continue }

            $proposed = New-Object System.Drawing.Size($textW, 4000)
            $fits = $true
            foreach ($cap in $list) {
                $sz = [System.Windows.Forms.TextRenderer]::MeasureText(
                        "$cap", $Font, $proposed, $flags)

                # Width can come back GREATER than proposed: WordBreak breaks
                # between words, so a single word longer than the box overflows
                # rather than wrapping. That is the check that catches an
                # unbreakable caption -- height alone would call it a pass.
                if ($sz.Width -gt $textW -or $sz.Height -gt $allowedH) {
                    $fits = $false
                    break
                }
            }
            if ($fits) { return $w }
        }
        return $Maximum
    } catch {
        # Measuring is an optimisation, not a requirement. If GDI is
        # unavailable for any reason the console must still open, so fall back
        # to the width it used before this function existed.
        return 150
    }
}

function Test-ConsoleSectionOverflow {
    <#
        Does any section hold more tools than it declared room for?

        This is what "allow room for a scrollbar" is actually protecting
        against, so it is answered in one place and consumed twice: the console
        uses it to decide panel width when Sections.json has not stated a
        preference, and ToolManager's Sections dialog uses it to warn that
        unticking the box will wrap buttons.

        Counted from every tool, including ones a capability check will hide.
        Which tools are hidden is not known at layout time, and over-counting
        reserves space that proves unnecessary where under-counting wraps a
        button the first time a capability is switched on.
    #>
    param($Sections, $Tools, [int]$ButtonsPerRow)

    $names = @(@($Sections) | ForEach-Object { "$($_.Name)" } | Where-Object { $_ })
    if ($names.Count -eq 0) { return $false }
    if ($ButtonsPerRow -lt 1) { $ButtonsPerRow = 1 }

    $counts = @{}
    foreach ($n in $names) { $counts[$n] = 0 }
    foreach ($t in @($Tools)) {
        # Same fallback the console applies: an unknown or blank section goes to
        # the first one, so a tool cannot become invisible by being miscounted
        # here and mis-placed there.
        $s = "$($t.Section)".Trim()
        if (-not $counts.ContainsKey($s)) { $s = $names[0] }
        $counts[$s] = $counts[$s] + 1
    }

    foreach ($sec in @($Sections)) {
        $r = [int]$sec.Rows
        if ($r -lt 1) { $r = 1 }
        if ($counts["$($sec.Name)"] -gt ($r * $ButtonsPerRow)) { return $true }
    }
    return $false
}

function Get-ConsolePanelWidth {
    <#
        Width of a section panel holding ButtonsPerRow buttons.

        Lives here rather than in AdminConsole.ps1 because ToolManager's
        Sections dialog previews the same number, and two copies of layout
        arithmetic drift the moment one of them is adjusted (ARCHITECTURE 14n).

        THE SCROLLBAR ALLOWANCE IS CONDITIONAL, AND THE CONDITION IS GLOBAL.

        It used to be unconditional, which was the fix for a real bug: a panel
        that grows a scrollbar loses VerticalScrollBarWidth of client area, and
        the section that did so wrapped its last button onto a new row while the
        other four did not -- so one section looked special-cased when it was
        merely the only one full enough to notice.

        Reserving it always solved that and cost ~27px down the right-hand edge
        of every panel, permanently, because a FlowLayoutPanel is left-aligned
        and all the unused width collects on the right. Once button width became
        measured rather than fixed, sections stopped being full at all and the
        reservation became pure dead space.

        So the caller decides ONCE, for every panel, whether any section can
        overflow -- not per panel. That keeps the property that actually
        mattered (every panel is the same width, so no section can look
        different from its neighbours) while giving the width back in the normal
        case. Add enough tools to overfill a section and all panels widen
        together.

        The slack is what stops an exact fit. Without it the sum is precisely
        what a full row needs, and an exact fit is the same as no fit as soon as
        anything rounds: the form's AutoScaleMode scales control sizes at load,
        so six buttons can each round up a fraction and collectively exceed a
        panel that rounded down. It is smaller than the old +8 because the
        scrollbar -- the thing the old slack was really hedging against, being
        in device pixels while buttons scale with the font -- is no longer in
        the sum unless it is actually needed.

        AdminConsole splits this slack across both edges with panel Padding, so
        it reads as a margin rather than as a gap on one side.
    #>
    param(
        [int]$ButtonWidth,
        [int]$ButtonPad,
        [int]$ButtonsPerRow,
        [switch]$ReserveScrollBar
    )
    if ($ButtonsPerRow -lt 1) { $ButtonsPerRow = 1 }
    $cell = $ButtonWidth + ($ButtonPad * 2)

    $sb = 0
    if ($ReserveScrollBar) {
        $sb = 17
        try {
            $sb = [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth
        } catch { }
    }

    # +2 for the panel's FixedSingle border, one pixel each side.
    return ($cell * $ButtonsPerRow) + 2 + $sb + $script:ConsolePanelSlack
}

function Set-FlatStyleSafely {
    <#
        Sets FlatStyle only where the control actually has it.

        Not every control does -- TextBox, NumericUpDown and ProgressBar among
        them -- and assigning a property that does not exist is a terminating
        error that takes the whole theme pass down. The symptom is a tool that
        refuses to open with a message about a property, which reads like
        corruption rather than a styling problem.
    #>
    param([Parameter(Mandatory)]$Control)
    if ($Control.PSObject.Properties['FlatStyle']) {
        $Control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    }
}

function Set-ConsoleTheme {
    param(
        [Parameter(Mandatory)]$Form,
        $Theme,
        [switch]$NoGradient
    )
    if (-not $Theme) { $Theme = Get-ConsoleTheme }
    if (-not $Theme) { return }

    $formBack   = ConvertTo-ThemeColor $Theme.FormBack
    $formText   = ConvertTo-ThemeColor $Theme.FormText
    $panelBack  = ConvertTo-ThemeColor $Theme.PanelBack
    $groupText  = ConvertTo-ThemeColor $Theme.GroupText
    $btnBack    = ConvertTo-ThemeColor $Theme.ButtonBack
    $btnText    = ConvertTo-ThemeColor $Theme.ButtonText
    $btnBorder  = ConvertTo-ThemeColor $Theme.ButtonBorder
    $btnHover   = ConvertTo-ThemeColor $Theme.ButtonHover
    $btnDown    = ConvertTo-ThemeColor $Theme.ButtonDown
    $inputBack  = ConvertTo-ThemeColor $Theme.InputBack
    $inputText  = ConvertTo-ThemeColor $Theme.InputText

    # Selected radio/checkbox text. Falls back to the button border colour,
    # which is already the theme's accent, so an older Themes.json without this
    # role still gets a visible selection rather than none.
    $selText = ConvertTo-ThemeColor $Theme.SelectedText
    if (-not $selText) { $selText = ConvertTo-ThemeColor $Theme.ButtonBorder }

    # ServiceNow green. Defaults are used when the theme does not define them,
    # so the colour is consistent across every theme rather than disappearing
    # on the ones that predate this.
    $snBack   = ConvertTo-ThemeColor $Theme.ServiceNowBack
    $snText   = ConvertTo-ThemeColor $Theme.ServiceNowText
    $snHover  = ConvertTo-ThemeColor $Theme.ServiceNowHover
    $snBorder = ConvertTo-ThemeColor $Theme.ServiceNowBorder
    if (-not $snBack)   { $snBack   = [System.Drawing.Color]::FromArgb(10, 92, 66) }
    if (-not $snText)   { $snText   = [System.Drawing.Color]::FromArgb(255, 255, 255) }
    if (-not $snHover)  { $snHover  = [System.Drawing.Color]::FromArgb(14, 120, 86) }
    if (-not $snBorder) { $snBorder = [System.Drawing.Color]::FromArgb(22, 160, 114) }

    if ($formBack) { $Form.BackColor = $formBack }
    if ($formText) { $Form.ForeColor = $formText }

    if (-not $NoGradient -and $Theme.GradientTo -and $Theme.FormBack) {
        Add-FormGradient -Form $Form -FromHex $Theme.FormBack -ToHex $Theme.GradientTo
    }

    $apply = {
        param($parent)
        foreach ($c in $parent.Controls) {
            # Tag-based opt-out. Controls whose colour carries MEANING -- the
            # ExO/SPO connection state, the ServiceNow highlight -- set
            # Tag = 'NoTheme' and are skipped, so a theme can never disguise a
            # red "not connected" as decoration.
            if ("$($c.Tag)" -like '*NoTheme*') { continue }

            switch -Regex ($c.GetType().Name) {
                # Anchored. 'Button' unanchored also matches RadioButton, and
                # switch -Regex falls through every matching branch without a
                # break -- so a RadioButton ran the button styling AND its own,
                # and anything reaching a FlatStyle line that lacks the property
                # threw. Set-FlatStyleSafely below is the second guard.
                '^Button$' {
                    if ($btnBack) {
                        Set-FlatStyleSafely -Control $c
                        $c.BackColor = $btnBack
                        if ($btnText)   { $c.ForeColor = $btnText }
                        if ($btnBorder) { $c.FlatAppearance.BorderColor = $btnBorder }
                        if ($btnHover)  { $c.FlatAppearance.MouseOverBackColor = $btnHover }
                        if ($btnDown)   { $c.FlatAppearance.MouseDownBackColor = $btnDown }
                    }
                    # ServiceNow buttons get the ServiceNow green so the most
                    # used destination is findable without reading labels.
                    # Matched on the button TEXT, not Tag -- tool buttons carry
                    # their tool object in Tag, so that slot is taken.
                    if ($snBack -and "$($c.Text)" -match '(?i)service\s*now') {
                        Set-FlatStyleSafely -Control $c
                        $c.BackColor = $snBack
                        if ($snText)   { $c.ForeColor = $snText }
                        if ($snHover)  { $c.FlatAppearance.MouseOverBackColor = $snHover }
                        if ($snBorder) { $c.FlatAppearance.BorderColor = $snBorder }
                    }
                }
                'GroupBox' {
                    $c.BackColor = [System.Drawing.Color]::Transparent
                    if ($groupText) { $c.ForeColor = $groupText }
                }
                '^(RadioButton|CheckBox)$' {
                    # WinForms draws the radio dot and tick with system colours,
                    # which vanish against a dark BackColor -- the control is
                    # readable but you cannot tell which option is selected.
                    # Colour and weight carry the state instead of the glyph.
                    $c.BackColor = [System.Drawing.Color]::Transparent
                    Set-FlatStyleSafely -Control $c

                    # Fonts built once per control and stashed on it. Creating a
                    # Font on every state change would leak a GDI handle each
                    # time the user clicks.
                    Add-Member -InputObject $c -NotePropertyName '_FontRegular' `
                        -NotePropertyValue $c.Font -Force
                    Add-Member -InputObject $c -NotePropertyName '_FontBold' `
                        -NotePropertyValue (New-Object System.Drawing.Font($c.Font, [System.Drawing.FontStyle]::Bold)) -Force
                    Add-Member -InputObject $c -NotePropertyName '_ClrOn'  -NotePropertyValue $selText  -Force
                    Add-Member -InputObject $c -NotePropertyName '_ClrOff' -NotePropertyValue $formText -Force

                    # $this inside the handler is the control that raised it, so
                    # one shared scriptblock serves every radio and checkbox.
                    $c.Add_CheckedChanged({
                        if ($this.Checked) {
                            if ($this._ClrOn)   { $this.ForeColor = $this._ClrOn }
                            if ($this._FontBold) { $this.Font = $this._FontBold }
                        } else {
                            if ($this._ClrOff)      { $this.ForeColor = $this._ClrOff }
                            if ($this._FontRegular) { $this.Font = $this._FontRegular }
                        }
                    })

                    # Apply the current state now, so whatever is already
                    # selected is visible before the user touches anything.
                    if ($c.Checked) {
                        if ($selText) { $c.ForeColor = $selText }
                        $c.Font = $c._FontBold
                    } elseif ($formText) {
                        $c.ForeColor = $formText
                    }
                }
                '^(Label|LinkLabel)$' {
                    $c.BackColor = [System.Drawing.Color]::Transparent
                    if ($formText) { $c.ForeColor = $formText }
                }
                '^(TextBox|RichTextBox|ComboBox|ListBox|MaskedTextBox|NumericUpDown)$' {
                    if ($inputBack) { $c.BackColor = $inputBack }
                    if ($inputText) { $c.ForeColor = $inputText }
                }
                '^(ListView|TreeView)$' {
                    if ($inputBack) { $c.BackColor = $inputBack }
                    if ($inputText) { $c.ForeColor = $inputText }
                }
                '^DataGridView$' {
                    <#
                        A DataGridView does NOT paint its cells from BackColor
                        and ForeColor. Setting only those themed the strip of
                        empty space below the last row and left every cell on
                        its own defaults -- which is why ToolManager's 32 rows
                        were dark text on a dark background and needed
                        squinting. Cells come from DefaultCellStyle, headers
                        from ColumnHeadersDefaultCellStyle, and neither inherits
                        from the control.

                        EnableHeadersVisualStyles = $false is the non-obvious
                        one: while it is $true, Windows draws the headers with
                        the system visual style and silently ignores every
                        header colour set here.
                    #>
                    if ($inputBack) {
                        $c.BackgroundColor = $inputBack
                        $c.DefaultCellStyle.BackColor = $inputBack
                    }
                    if ($inputText) { $c.DefaultCellStyle.ForeColor = $inputText }

                    # Zebra striping. With three dozen rows this does more for
                    # readability than the text colour does -- it gives the eye
                    # a row boundary to track along.
                    if ($inputBack) {
                        $c.AlternatingRowsDefaultCellStyle.BackColor =
                            (Get-ShadedColor -Base $inputBack -Amount 14)
                        if ($inputText) { $c.AlternatingRowsDefaultCellStyle.ForeColor = $inputText }
                    }

                    if ($btnBack) { $c.DefaultCellStyle.SelectionBackColor = $btnBack }
                    $selFore = if ($selText) { $selText } elseif ($btnText) { $btnText } else { $null }
                    if ($selFore) { $c.DefaultCellStyle.SelectionForeColor = $selFore }

                    $c.EnableHeadersVisualStyles = $false
                    if ($panelBack) {
                        $c.ColumnHeadersDefaultCellStyle.BackColor = $panelBack
                        $c.RowHeadersDefaultCellStyle.BackColor    = $panelBack
                        $c.ColumnHeadersDefaultCellStyle.SelectionBackColor = $panelBack
                        $c.RowHeadersDefaultCellStyle.SelectionBackColor    = $panelBack
                    }
                    $hdrFore = if ($groupText) { $groupText } elseif ($formText) { $formText } else { $null }
                    if ($hdrFore) {
                        $c.ColumnHeadersDefaultCellStyle.ForeColor = $hdrFore
                        $c.RowHeadersDefaultCellStyle.ForeColor    = $hdrFore
                        $c.ColumnHeadersDefaultCellStyle.SelectionForeColor = $hdrFore
                        $c.RowHeadersDefaultCellStyle.SelectionForeColor    = $hdrFore
                    }
                    if ($btnBorder) { $c.GridColor = $btnBorder }

                    # Read-only columns get dimmed text so "you cannot edit
                    # this" reads as deliberate rather than as a rendering
                    # fault -- the same complaint that started this fix.
                    if ($inputText -and $inputBack) {
                        $dim = Get-BlendedColor -From $inputText -To $inputBack -Amount 0.45
                        foreach ($col in @($c.Columns)) {
                            if ($col.ReadOnly) { $col.DefaultCellStyle.ForeColor = $dim }
                        }
                    }
                }
                '^(Panel|FlowLayoutPanel|TableLayoutPanel|TabPage|SplitterPanel)$' {
                    if ($panelBack) { $c.BackColor = $panelBack }
                    if ($formText)  { $c.ForeColor = $formText }
                }
            }
            if ($c.Controls -and $c.Controls.Count -gt 0) { & $apply $c }
        }
    }
    & $apply $Form
}

function Get-CmdletParameterNames {
    <#
        Returns the parameter names a cmdlet actually declares.

        Exists because guessing is unreliable: -PageSize and
        -RecipientTypeDetails were both assumed onto Get-EXOCasMailbox and both
        rejected. The ExchangeOnlineManagement surface varies by version and
        tenant, so the command object in front of us is the only source of
        truth -- not the documentation.
    #>
    param([string]$Name)
    try {
        $cmd = Get-EXOCommand -Name $Name
        if ($cmd.Parameters) { return @($cmd.Parameters.Keys) }
    } catch { }
    try {
        $cmd = Get-Command -Name $Name -ErrorAction Stop
        if ($cmd.Parameters) { return @($cmd.Parameters.Keys) }
    } catch { }
    return @()
}

function Invoke-EXOBulkQuery {
    <#
        A bulk Exchange Online query that survives the two things that reliably
        break one: an unsupported property, and a dropped connection.

          * Only parameters the installed cmdlet declares are passed.
          * "InvalidProperties = X, Y" is parsed, the offenders dropped, retried.
          * Transport faults retry with exponential backoff.

        Note this does NOT fix an oversized response -- retrying an enormous
        query reproduces an enormous query. Callers dealing with thousands of
        objects should slice the request (by recipient type, by alias letter)
        and call this per slice.
    #>
    param(
        [Parameter(Mandatory)][string]$Cmdlet,
        [string[]]$Properties,
        [hashtable]$Extra = @{},
        [scriptblock]$Progress,
        [int]$MaxAttempts = 5
    )
    $supported = Get-CmdletParameterNames -Name $Cmdlet
    $remaining = @($Properties)

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $params = @{ ErrorAction = 'Stop' }
        if ($supported -contains 'ResultSize') { $params['ResultSize'] = 'Unlimited' }
        if ($remaining -and $supported -contains 'Properties') { $params['Properties'] = $remaining }
        foreach ($k in $Extra.Keys) {
            if ($supported -contains $k) { $params[$k] = $Extra[$k] }
        }
        try {
            return @(Invoke-EXOCommand -Name $Cmdlet -Parameters $params)
        } catch {
            $msg = "$($_.Exception.Message)"

            $m = [regex]::Match($msg, 'InvalidProperties\s*=\s*(.+?)\s*$')
            if ($m.Success) {
                $bad = $m.Groups[1].Value -split '\s*,\s*' | Where-Object { $_ }
                $remaining = @($remaining | Where-Object { $bad -notcontains $_ })
                if (-not $remaining) { throw }
                if ($Progress) { & $Progress "Dropped unsupported: $($bad -join ', '). Retrying..." }
                continue
            }

            $transient = 'underlying connection was closed|operation has timed out|' +
                         'unexpected error occurred on a receive|task was canceled|' +
                         'Unable to read data from the transport|forcibly closed'
            if (($msg -match $transient) -and $attempt -lt $MaxAttempts) {
                $wait = [math]::Pow(2, $attempt)
                if ($Progress) { & $Progress "Connection dropped. Retrying in $wait s..." }
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }
    throw "Could not query $Cmdlet after $MaxAttempts attempts."
}

function Get-AcceptedDomainList {
    <#
        Accepted domains for the connected tenant, lowercased.

        "External" is only meaningful relative to this list -- a forward to
        another accepted domain is internal mail, not exfiltration.
    #>
    param([switch]$Refresh)
    if (-not $Refresh -and $global:AdminConsoleAcceptedDomains) {
        return $global:AdminConsoleAcceptedDomains
    }
    try {
        # ResultSize is mandatory here, not tidiness. Get-AcceptedDomain
        # defaults to 1000 and a tenant with more than that gets a SILENTLY
        # TRUNCATED list plus a warning on the host stream. Truncation is the
        # dangerous half: this list is what "external" is measured against, so
        # a domain past the cut-off would make internal mail read as
        # exfiltration in the External Forwarding report.
        $d = @(Invoke-EXOCommand -Name 'Get-AcceptedDomain' -Parameters @{
                   ResultSize = 'Unlimited'; ErrorAction = 'Stop' } |
               ForEach-Object { "$($_.DomainName)".ToLower() } | Where-Object { $_ })
        if ($d.Count -gt 0) {
            $global:AdminConsoleAcceptedDomains = $d
            return $d
        }
    } catch { }
    return @()
}

function Get-SmtpDomain {
    param([string]$Address)
    if (-not $Address) { return '' }
    $a = "$Address"
    # Recipient values arrive in several shapes: plain SMTP, "Display Name
    # [SMTP:user@dom]", or an X500/legacy DN. Pull the last @-part of whatever
    # is there rather than assuming a clean address.
    $m = [regex]::Match($a, '([A-Za-z0-9._%+\-]+)@([A-Za-z0-9.\-]+\.[A-Za-z]{2,})')
    if ($m.Success) { return $m.Groups[2].Value.ToLower() }
    return ''
}

function New-ProgressReporter {
    <#
        Returns a scriptblock that reports progress to every place at once:
        a WinForms label, an optional progress bar, and Write-Progress in the
        PowerShell window.

        Long reports look identical to hung reports. A run that queries
        thousands of mailboxes can sit silent for minutes, and the only honest
        difference between "working" and "stalled" is whether something is
        still changing on screen.

        Usage:
            $p = New-ProgressReporter -Label $lblStatus -Bar $pb -Activity 'Protocols'
            & $p 'Querying UserMailbox...'
            & $p 'Half way' 50          # optional percent

        DoEvents is called deliberately: these tools are single-threaded, so
        without it the form never repaints and Windows greys it out as
        "Not Responding" -- which looks exactly like the crash we are trying
        to disprove.
    #>
    param(
        $Label,
        $Bar,
        [string]$Activity = 'Working',
        [switch]$NoConsole
    )
    $state = [pscustomobject]@{
        Label     = $Label
        Bar       = $Bar
        Activity  = $Activity
        NoConsole = [bool]$NoConsole
        Started   = Get-Date
    }
    return {
        param([string]$Message, [int]$Percent = -1)
        $elapsed = (Get-Date) - $state.Started
        $stamp   = '{0:mm\:ss}' -f $elapsed

        if ($state.Label) {
            $state.Label.Text = "[$stamp] $Message"
        }
        if ($state.Bar) {
            if ($Percent -ge 0) {
                $state.Bar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
                $state.Bar.Value = [Math]::Max(0, [Math]::Min(100, $Percent))
            } else {
                # Unknown duration -- marquee says "alive" without lying about
                # how far along it is.
                $state.Bar.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
            }
        }
        if (-not $state.NoConsole) {
            if ($Percent -ge 0) {
                Write-Progress -Activity $state.Activity -Status $Message -PercentComplete $Percent
            } else {
                Write-Progress -Activity $state.Activity -Status $Message
            }
        }
        [System.Windows.Forms.Application]::DoEvents()
    }.GetNewClosure()
}

function Complete-ProgressReporter {
    param([string]$Activity = 'Working', $Bar)
    try { Write-Progress -Activity $Activity -Completed } catch { }
    if ($Bar) {
        $Bar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
        $Bar.Value = 100
    }
}

function Add-ReportProgressBar {
    <#
        Standard progress bar for report windows. Placed by the caller; this
        just keeps the styling consistent across every report.
    #>
    param(
        [Parameter(Mandatory)]$Form,
        [int]$X = 12, [int]$Y = 0, [int]$Width = 720, [int]$Height = 14
    )
    $pb = New-Object System.Windows.Forms.ProgressBar
    $pb.Location = New-Object System.Drawing.Point($X, $Y)
    $pb.Size     = New-Object System.Drawing.Size($Width, $Height)
    $pb.Style    = [System.Windows.Forms.ProgressBarStyle]::Continuous
    $pb.Value    = 0
    $pb.Tag      = 'NoTheme'
    $Form.Controls.Add($pb)
    return $pb
}

#region Settings editor ------------------------------------------------------
# LIVES HERE BECAUSE IT HAS THREE CALLERS, not because settings are a "common"
# idea. It was inside ToolManager.ps1 while that was the only way in; the
# dashboard's Settings button and Settings.ps1 made that a trip through the tool
# registry to change an email address.
#
# Everything it calls was already in this file, so the move cost nothing. The
# alternative -- a second copy in Settings.ps1 -- is the one that costs: two
# editors of the same file drift, and the one you are not looking at is the one
# that wrote it.
function Show-SettingsEditor {
    <#
        Edits Settings.json: the per-MACHINE locations this toolkit writes to and
        reads from, plus the Edge profile names.

        NOT tenant facts. Those live in Config\Tenant.json and are edited by
        TenantManager. This comment used to say tenant constants stay hardcoded
        in Common.ps1, which was true before the tenant-agnostic refactor and
        has been false since -- one toolkit now ships to every tenant and the
        only thing separating them is Tenant.json.

        The split that matters: Tenant.json is the same on every machine at a
        tenant, Settings.json is different on every machine. What does NOT
        belong in either is anything in code that assumes a drive letter --
        'E:\Scripts\M365\Logs\Completed' works on the machine it was written on
        and fails silently on a server built differently, at the moment it
        tries to write.

        Blank means "beside the toolkit", which always exists.

        Returns $true if the file was written.
    #>
    $dlg                 = New-Object System.Windows.Forms.Form
    $dlg.Text            = 'Settings'
    $dlg.Size            = New-Object System.Drawing.Size(620, 550)
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false
    $dlg.StartPosition   = 'CenterParent'
    $dlg.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $current = Get-ToolkitSettings

    $lblInfo          = New-Object System.Windows.Forms.Label
    $lblInfo.Text     = 'Where this toolkit writes and reads. Leave a box blank to use a folder beside the toolkit -- nothing here assumes a particular drive exists.'
    $lblInfo.Location = New-Object System.Drawing.Point(12, 12)
    $lblInfo.Size     = New-Object System.Drawing.Size(584, 32)
    $dlg.Controls.Add($lblInfo)

    $grpPaths          = New-Object System.Windows.Forms.GroupBox
    $grpPaths.Text     = 'Folders'
    $grpPaths.Location = New-Object System.Drawing.Point(12, 52)
    $grpPaths.Size     = New-Object System.Drawing.Size(584, 200)
    $dlg.Controls.Add($grpPaths)

    $newFolderRow = {
        param($Parent, $LabelText, $Y, $Value, $Description)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $LabelText
        $l.Location = New-Object System.Drawing.Point(12, ($Y + 3))
        $l.Size = New-Object System.Drawing.Size(118, 20)
        $Parent.Controls.Add($l)

        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Location = New-Object System.Drawing.Point(134, $Y)
        $tb.Size = New-Object System.Drawing.Size(340, 22)
        $tb.Text = $Value
        $Parent.Controls.Add($tb)

        $b = New-Object System.Windows.Forms.Button
        $b.Text = 'Browse...'
        $b.Location = New-Object System.Drawing.Point(480, ($Y - 1))
        $b.Size = New-Object System.Drawing.Size(88, 24)
        $b.Add_Click({
            $fb = New-Object System.Windows.Forms.FolderBrowserDialog
            $fb.Description = $Description
            if ($tb.Text -and (Test-Path -LiteralPath $tb.Text)) { $fb.SelectedPath = $tb.Text }
            if ($fb.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $tb.Text = $fb.SelectedPath }
        }.GetNewClosure())
        $Parent.Controls.Add($b)
        return $tb
    }

    $txtReports = & $newFolderRow $grpPaths 'Reports:'         26 $current.ReportsPath        'Where report CSVs are written'
    $txtLogs    = & $newFolderRow $grpPaths 'Logs:'            60 $current.LogPath            'Where CompletedTasks.csv is written'
    $txtDisable = & $newFolderRow $grpPaths 'Retry queue:'      94 $current.RetryQueuePath     'Folder holding retry queues -- work a tool could not finish. Blank uses RetryQueue beside the toolkit'

    # Not a folder, so it does not go through $newFolderRow -- and it is per
    # OPERATOR rather than per machine, which is the point: notifications copy
    # to whoever ran the tool instead of to one hardcoded person.
    $lblOper          = New-Object System.Windows.Forms.Label
    $lblOper.Text     = 'Your email:'
    $lblOper.Location = New-Object System.Drawing.Point(12, 132)
    $lblOper.Size     = New-Object System.Drawing.Size(110, 20)
    $grpPaths.Controls.Add($lblOper)

    $txtOper          = New-Object System.Windows.Forms.TextBox
    $txtOper.Location = New-Object System.Drawing.Point(126, 129)
    $txtOper.Size     = New-Object System.Drawing.Size(330, 22)
    $txtOper.Text     = "$($current.OperatorEmail)"
    $grpPaths.Controls.Add($txtOper)

    $lblOperHint           = New-Object System.Windows.Forms.Label
    $lblOperHint.Text      = 'Copies of notifications sent by the disable tool go here. Leave blank to use the mail address on your own AD account; if that is empty, no copy is sent rather than one going to the wrong person.'
    $lblOperHint.Location  = New-Object System.Drawing.Point(126, 154)
    $lblOperHint.Size      = New-Object System.Drawing.Size(440, 30)
    $lblOperHint.ForeColor = [System.Drawing.Color]::DimGray
    $grpPaths.Controls.Add($lblOperHint)

    # SPO ADMIN URL AND EXCHANGE SERVER USED TO BE HERE. REMOVED 2026-08-13.
    #
    # Both are facts about the TENANT, and this file's own header draws that
    # line: Tenant.json holds domains, directories and servers; Settings.json
    # holds what an engineer prefers -- paths and browser profiles. An admin
    # centre host is not a preference.
    #
    # Having them in both places was not merely untidy. Get-ExchangeServerUri
    # read Settings FIRST, so a machine-local box silently overrode the tenant
    # fact, and two engineers at one tenant could be pointed at different
    # servers with both files looking correct. Worse on the SharePoint side:
    # Get-SpoAdminUrl never read Tenant.json at all, so the box in Tenant
    # Manager was inert and Connect SPO used a derived URL that had never
    # existed.
    #
    # Both now live in Tenant Manager only. Nothing is lost -- that is where
    # they were always supposed to be edited.
    $lblDisHint           = New-Object System.Windows.Forms.Label
    $lblDisHint.Text      = 'Retry queue: work a tool could not finish and will pick up again. Blank uses RetryQueue beside the toolkit and is created when needed; a path you enter here is never created, because an empty folder would read as an empty queue.'
    $lblDisHint.Location  = New-Object System.Drawing.Point(12, 258)
    $lblDisHint.Size      = New-Object System.Drawing.Size(584, 32)
    $lblDisHint.ForeColor = [System.Drawing.Color]::DimGray
    $dlg.Controls.Add($lblDisHint)

    $grpBrowser          = New-Object System.Windows.Forms.GroupBox
    $grpBrowser.Text     = 'Edge profiles  --  the profile NAME as shown in Edge, not "Profile 2"'
    $grpBrowser.Location = New-Object System.Drawing.Point(12, 296)
    $grpBrowser.Size     = New-Object System.Drawing.Size(584, 108)
    $dlg.Controls.Add($grpBrowser)

    $existing = Get-BrowserProfileMap
    $order    = @($existing.Keys | Sort-Object)
    $boxes    = @{}
    $y = 24
    foreach ($purpose in $order) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = "$purpose" + ':'
        $l.Location = New-Object System.Drawing.Point(12, ($y + 3))
        $l.Size = New-Object System.Drawing.Size(84, 20)
        $grpBrowser.Controls.Add($l)

        $cb = New-Object System.Windows.Forms.ComboBox
        $cb.DropDownStyle = 'DropDown'
        $cb.Location = New-Object System.Drawing.Point(100, $y)
        $cb.Size = New-Object System.Drawing.Size(200, 22)
        $cb.Text = "$($existing[$purpose])"
        $grpBrowser.Controls.Add($cb)
        $boxes[$purpose] = $cb
        $y += 26
    }

    # Offer the profiles Edge actually has: a name that does not exist falls back
    # to the last-used profile at run time, which usually works -- so the mistake
    # surfaces as a portal opening under the wrong account, not as an error.
    $lblProfWarn           = New-Object System.Windows.Forms.Label
    $lblProfWarn.Location  = New-Object System.Drawing.Point(310, 24)
    $lblProfWarn.Size      = New-Object System.Drawing.Size(262, 74)
    $lblProfWarn.ForeColor = [System.Drawing.Color]::DimGray
    $grpBrowser.Controls.Add($lblProfWarn)

    $edgeProfiles = @()
    try { $edgeProfiles = @(Get-EdgeProfiles | ForEach-Object { "$($_.DisplayName)" } | Where-Object { $_ }) } catch { }
    if (@($edgeProfiles).Count -gt 0) {
        foreach ($cb in $boxes.Values) {
            $edgeProfiles | Sort-Object -Unique | ForEach-Object { [void]$cb.Items.Add($_) }
        }
        $lblProfWarn.Text = "On this machine: $((@($edgeProfiles) | Sort-Object -Unique) -join ', ')"
    } else {
        $lblProfWarn.Text = 'No Edge profiles could be read here. Type the names exactly as they appear in Edge.'
    }

    $btnOK          = New-Object System.Windows.Forms.Button
    $btnOK.Text     = 'Save'
    $btnOK.Location = New-Object System.Drawing.Point(386, 450)
    $btnOK.Size     = New-Object System.Drawing.Size(100, 28)
    $btnOK.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
    $dlg.Controls.Add($btnOK)

    $btnCancel          = New-Object System.Windows.Forms.Button
    $btnCancel.Text     = 'Cancel'
    $btnCancel.Location = New-Object System.Drawing.Point(494, 450)
    $btnCancel.Size     = New-Object System.Drawing.Size(100, 28)
    $btnCancel.Add_Click({ $dlg.Close() })
    $dlg.Controls.Add($btnCancel)

    $btnOK.Add_Click({
        $dis = "$($txtDisable.Text)".Trim()
        if ($dis -and -not (Test-Path -LiteralPath $dis)) {
            $ans = [System.Windows.Forms.MessageBox]::Show(
                "The retry queue folder does not exist:`n`n$dis`n`nA path entered here is never created automatically, so nothing will be able to save a retry queue until it exists.`n`nSave it anyway?",
                'Settings', 'YesNo', 'Warning', 'Button2')
            if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }
        $oper = "$($txtOper.Text)".Trim()
        if ($oper -and $oper -notmatch '^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$') {
            [void][System.Windows.Forms.MessageBox]::Show(
                "That does not look like an email address:`n`n$oper", 'Settings', 'OK', 'Warning')
            return
        }
        $prof = @{}
        foreach ($k in $boxes.Keys) { $prof[$k] = "$($boxes[$k].Text)".Trim() }
        try {
            [void](Save-ToolkitSettings -ReportsPath "$($txtReports.Text)".Trim() `
                        -LogPath "$($txtLogs.Text)".Trim() `
                        -RetryQueuePath $dis `
                        -OperatorEmail "$($txtOper.Text)".Trim() `
                        -BrowserProfiles $prof)
            $dlg.Tag = 'saved'
            $dlg.Close()
        } catch {
            [void][System.Windows.Forms.MessageBox]::Show(
                "Could not save Settings.json:`n$($_.Exception.Message)",
                'Settings', 'OK', 'Error')
        }
    })

    Set-ConsoleTheme -Form $dlg
    [void]$dlg.ShowDialog()
    return ($dlg.Tag -eq 'saved')
}
#endregion

#region Report files ---------------------------------------------------------
<#
    Reading a report must not require running one.

    Every report window used to keep "Open Folder" and "Open CSV" disabled until
    a run completed, so looking at yesterday's output meant regenerating it --
    which on the slower reports is minutes of Exchange calls to read a file that
    was already on disk. The reports directory is the archive; these controls
    expose it.
#>

function Open-ReportFolder {
    <#
        Opens the reports directory, creating it if this is a first run.

        The directory is created rather than reported missing because an empty
        folder is the honest answer to "show me my reports" before any have been
        written -- and explorer.exe on a non-existent path fails silently, which
        would look like a broken button.
    #>
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
    Start-Process explorer.exe $Path
}

function Get-ReportFileList {
    <#
        Previous report runs, newest first.

        Sorting is by the timestamp in the FILENAME, not LastWriteTime, because
        the filename records when the report was *run*. A file that has been
        copied, restored from backup or touched by a sync client carries a
        mistaken mtime, and sorting by it would silently reorder history. Files
        that don't parse fall back to LastWriteTime rather than being dropped.

        Row counts are deliberately NOT computed here. Counting rows means
        opening every file on every refresh, which turns a folder of large CSVs
        into a visible stall. The caller reads the count for the ONE file the
        user selects.
    #>
    param(
        [Parameter(Mandatory)][string]$Directory,
        [string]$Prefix,
        # Almost every report is a CSV, so that stays the default and no
        # existing caller changes. Snapshot-style reports (an exported rule
        # collection, for one) are XML, and without this they wrote a file the
        # Previous runs list could not see.
        [string]$Extension = 'csv'
    )
    if (-not (Test-Path -LiteralPath $Directory)) { return @() }

    $ext = "$Extension".TrimStart('.')
    $pattern = if ($Prefix) { "$Prefix`_*.$ext" } else { "*.$ext" }
    $files = @(Get-ChildItem -LiteralPath $Directory -Filter $pattern -File -ErrorAction SilentlyContinue)

    $out = foreach ($f in $files) {
        # Prefix_yyyy-MM-dd_HHmm.csv
        $stamp = $null
        if ($f.BaseName -match '_(\d{4}-\d{2}-\d{2})_(\d{2})(\d{2})$') {
            $parsed = [datetime]::MinValue
            if ([datetime]::TryParseExact(
                    "$($Matches[1]) $($Matches[2]):$($Matches[3])",
                    'yyyy-MM-dd HH:mm',
                    [System.Globalization.CultureInfo]::InvariantCulture,
                    [System.Globalization.DateTimeStyles]::None,
                    [ref]$parsed)) {
                $stamp = $parsed
            }
        }
        $fromName = ($null -ne $stamp)
        if (-not $fromName) { $stamp = $f.LastWriteTime }

        # Everything before the _yyyy-MM-dd_HHmm tail is the report name.
        $report = $f.BaseName -replace '_\d{4}-\d{2}-\d{2}_\d{4}$', ''

        [PSCustomObject]@{
            Path          = $f.FullName
            Name          = $f.Name
            Report        = $report
            RunAt         = $stamp
            StampFromName = $fromName
            SizeKB        = [math]::Round($f.Length / 1KB, 1)
        }
    }
    return @($out | Sort-Object RunAt -Descending)
}

function Get-ReportRowCount {
    <#
        Data rows in a CSV -- lines minus the header. Returns $null rather than
        0 if the file can't be read, because "0 rows" and "couldn't tell" must
        not display the same way.
    #>
    param([Parameter(Mandatory)][string]$Path)
    try {
        $n = 0
        $reader = [System.IO.File]::OpenText($Path)
        try { while ($null -ne $reader.ReadLine()) { $n++ } } finally { $reader.Dispose() }
        if ($n -le 1) { return 0 }
        return $n - 1
    } catch { return $null }
}

function Add-ReportFileControls {
    <#
        A "Previous runs" row for a report window: label, dropdown of prior
        runs of THIS report, and an Open button.

        Returns an object carrying Refresh() and SelectPath() so the caller can
        repopulate after writing a new file. State lives on the returned object,
        not in a $script: variable, because these controls are used by scripts
        the console launches with '&' -- where Common.ps1's script scope is not
        the caller's script scope and any $script: read would resolve to $null.

        The list refreshes on Form.Activated as well, so a file written by
        another report (or dropped in by hand) appears on returning to the
        window. That removes the need for a Refresh button.
    #>
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Form]$Form,
        [Parameter(Mandatory)][string]$ReportDir,
        [Parameter(Mandatory)][string]$Prefix,
        [int]$X = 12,
        [int]$Y = 0,
        [int]$Width = 720,
        [System.Windows.Forms.Label]$StatusLabel,
        [string]$Extension = 'csv'
    )

    $lbl          = New-Object System.Windows.Forms.Label
    $lbl.Text     = 'Previous runs:'
    $lbl.Location = New-Object System.Drawing.Point($X, ($Y + 5))
    $lbl.Size     = New-Object System.Drawing.Size(90, 18)
    $Form.Controls.Add($lbl)

    $comboX = $X + 94
    $btnW   = 102
    $btnX   = $X + $Width - $btnW
    $comboW = $btnX - $comboX - 4

    $combo               = New-Object System.Windows.Forms.ComboBox
    $combo.DropDownStyle = 'DropDownList'
    $combo.Location      = New-Object System.Drawing.Point($comboX, ($Y + 2))
    $combo.Size          = New-Object System.Drawing.Size($comboW, 22)
    $Form.Controls.Add($combo)

    $btn          = New-Object System.Windows.Forms.Button
    $btn.Text     = 'Open'
    $btn.Location = New-Object System.Drawing.Point($btnX, $Y)
    $btn.Size     = New-Object System.Drawing.Size($btnW, 28)
    $Form.Controls.Add($btn)

    $ctl = [PSCustomObject]@{
        Label       = $lbl
        Combo       = $combo
        OpenButton  = $btn
        Directory   = $ReportDir
        Prefix      = $Prefix
        StatusLabel = $StatusLabel
        Extension   = $Extension
        Files       = @()
    }

    Add-Member -InputObject $ctl -MemberType ScriptMethod -Name Refresh -Value {
        $keep = $null
        if ($this.Combo.SelectedIndex -ge 0 -and $this.Combo.SelectedIndex -lt @($this.Files).Count) {
            $keep = $this.Files[$this.Combo.SelectedIndex].Path
        }

        $ext = if ("$($this.Extension)".Trim()) { "$($this.Extension)" } else { 'csv' }
        $this.Files = @(Get-ReportFileList -Directory $this.Directory -Prefix $this.Prefix -Extension $ext)
        $this.Combo.Items.Clear()

        if (@($this.Files).Count -eq 0) {
            [void]$this.Combo.Items.Add('(no previous runs in the reports folder)')
            $this.Combo.SelectedIndex = 0
            $this.Combo.Enabled       = $false
            $this.OpenButton.Enabled  = $false
            return
        }

        foreach ($f in $this.Files) {
            # A '~' marks a timestamp taken from the file's mtime because the
            # name didn't carry one -- so a copied or renamed file is visibly
            # less trustworthy in the list rather than silently equal.
            $mark = if ($f.StampFromName) { ' ' } else { '~' }
            [void]$this.Combo.Items.Add(
                ('{0}{1}    {2,8:N1} KB    {3}' -f $mark,
                    $f.RunAt.ToString('yyyy-MM-dd  HH:mm'), $f.SizeKB, $f.Name))
        }
        $this.Combo.Enabled      = $true
        $this.OpenButton.Enabled = $true

        $idx = 0
        if ($keep) {
            for ($i = 0; $i -lt @($this.Files).Count; $i++) {
                if ($this.Files[$i].Path -eq $keep) { $idx = $i; break }
            }
        }
        $this.Combo.SelectedIndex = $idx
    }

    Add-Member -InputObject $ctl -MemberType ScriptMethod -Name SelectPath -Value {
        param([string]$Path)
        $this.Refresh()
        for ($i = 0; $i -lt @($this.Files).Count; $i++) {
            if ($this.Files[$i].Path -eq $Path) { $this.Combo.SelectedIndex = $i; return }
        }
    }

    Add-Member -InputObject $ctl -MemberType ScriptMethod -Name SelectedPath -Value {
        if (@($this.Files).Count -eq 0) { return $null }
        $i = $this.Combo.SelectedIndex
        if ($i -lt 0 -or $i -ge @($this.Files).Count) { return $null }
        return $this.Files[$i].Path
    }

    # $ctl travels on .Tag, and these handlers deliberately do NOT use
    # .GetNewClosure().
    #
    # GetNewClosure binds a scriptblock to a NEW DYNAMIC MODULE. Local variables
    # are copied in -- which is why $ctl resolved -- but FUNCTION lookup inside
    # a module goes module scope then global, and never sees the functions this
    # file put in the caller's SCRIPT scope. So Get-ReportRowCount came back
    # "not recognized" from inside the handler while working fine everywhere
    # else. Same family as the $script: trap: the scope a scriptblock runs in is
    # not the scope it was written in.
    $btn.Tag   = $ctl
    $combo.Tag = $ctl

    $btn.Add_Click({
        $ctl = $this.Tag
        $p = $ctl.SelectedPath()
        if (-not $p) { return }
        if (-not (Test-Path -LiteralPath $p)) {
            # Deleted or moved since the list was built.
            if ($ctl.StatusLabel) { $ctl.StatusLabel.Text = 'That file is no longer there -- list refreshed.' }
            $ctl.Refresh()
            return
        }
        Start-Process $p
    })

    $combo.Add_SelectedIndexChanged({
        $ctl = $this.Tag
        if (-not $ctl) { return }
        if (-not $ctl.StatusLabel) { return }
        if (@($ctl.Files).Count -eq 0) { return }
        $i = $ctl.Combo.SelectedIndex
        if ($i -lt 0 -or $i -ge @($ctl.Files).Count) { return }
        $f = $ctl.Files[$i]

        # Clearing the combo resets SelectedIndex to -1, so this fires on every
        # Refresh -- including the one on Form.Activated. Reading a very large
        # CSV on each window focus would be a stall with no visible cause, so
        # above 20 MB the size is reported instead of the row count.
        if ($f.SizeKB -gt 20480) {
            $ctl.StatusLabel.Text = "Selected: $($f.Name) -- $([math]::Round($f.SizeKB/1024,1)) MB (too large to count rows quickly)."
            return
        }

        # Row counting means "lines minus the header", which is only true of a
        # CSV. An XML snapshot has no rows, and reporting a line count as one
        # would be a confident wrong number rather than a missing one.
        if ("$($ctl.Extension)".TrimStart('.') -ne 'csv') {
            $ctl.StatusLabel.Text = "Selected: $($f.Name) -- $([math]::Round($f.SizeKB,1)) KB."
            return
        }
        $rows = Get-ReportRowCount -Path $f.Path
        $ctl.StatusLabel.Text = if ($null -eq $rows) {
            "Selected: $($f.Name) -- row count unavailable."
        } else {
            "Selected: $($f.Name) -- $rows data row(s)."
        }
    })

    # The form's Tag holds a LIST: more than one report window could in
    # principle carry these controls, and overwriting Tag would silently
    # disconnect the first.
    if ($Form.Tag -is [System.Collections.IList]) { [void]$Form.Tag.Add($ctl) }
    else { $Form.Tag = [System.Collections.ArrayList]@($ctl) }

    $Form.Add_Activated({
        foreach ($c in @($this.Tag)) { if ($c -and $c.PSObject.Methods['Refresh']) { $c.Refresh() } }
    })

    $ctl.Refresh()
    return $ctl
}
#endregion

#region Tool capability detection --------------------------------------------
<#
    Grey out a tool the current window cannot actually run.

    This works only because AdminConsole launches tools with the call operator
    (& $path), so a tool runs in the SAME PROCESS as the console. Exchange
    Online and SharePoint sessions are per-process (see the architecture notes,
    section 4), so the console's connection state genuinely IS the tool's state.
    If tools were ever launched with Start-Process, every judgement here becomes
    a guess and this whole mechanism must be reconsidered.

    Requirements are DETECTED by reading each script rather than declared in
    Tools.json. A declared field would have to be maintained by hand for every
    tool and would silently go stale the moment a script's dependencies changed
    -- the same drift shape as the duplicated $DefaultDC. Reading the script
    keeps one source of truth: the code.

    KNOWN LIMITS, both deliberate:

    * Dot-sourced files are NOT followed. Every tool dot-sources Common.ps1,
      and Common.ps1 itself calls Get-ADUser, Get-SPOTenant and Exchange
      cmdlets -- following it would tag every tool as needing everything, which
      would grey out the entire console. A tool is judged on what IT calls.

    * Ordinary strings are NOT stripped, only comments and here-strings. A
      cmdlet named in a string is usually a real dependency (for example
      Get-CmdletParameterNames -Name 'Get-SPOSite'), so removing them would
      lose signal. A cmdlet named only in a comment WOULD be a false positive,
      which is why comments go.
#>

$script:CapabilityDefs = @(
    [PSCustomObject]@{
        Id       = 'ExO'
        Label    = 'Exchange Online'
        Patterns = @(
            '\b(?:Get|Set)-EXO[A-Za-z]+'
            '\bConnect-ExchangeOnline\b'
            '\bGet-ConnectionInformation\b'
            '\b(?:Get|Set)-Mailbox\b'
            '\bGet-CASMailbox\b'
            '\b(?:Get|Set|New|Remove)-(?:InboxRule|DistributionGroup|UnifiedGroup|MailContact)\b'
            '\bGet-UnifiedGroupLinks\b'
            '\b(?:Get|Add|Remove)-(?:MailboxPermission|RecipientPermission|MailboxFolderPermission)\b'
            '\bGet-(?:MailboxStatistics|MailboxAutoReplyConfiguration|MailboxRegionalConfiguration)\b'
            '\bGet-(?:Recipient|AcceptedDomain|OrganizationConfig)\b'
            # Open-ended on purpose: Microsoft versions these cmdlet names in
            # place. Get-MessageTrace was retired in favour of
            # Get-MessageTraceV2, and a trailing \b here matched neither the V2
            # form nor Get-MessageTraceDetailV2 -- so MessageTrace.ps1 scanned
            # as needing nothing at all. Assume any future V3 does the same.
            '\bGet-MessageTrace[A-Za-z0-9]*'
            '\b(?:Get|Release|Delete|Preview|Export)-QuarantineMessage[A-Za-z0-9]*'
            '\bStart-HistoricalSearch\b'
            # Get-TransportRule exists on-premises too, so strictly it proves
            # only "some Exchange". It sits under ExO because the only tool
            # calling it is cloud-only and routes through Invoke-EXOCommand.
            # Matching the wrong world would be bad; matching NOTHING is worse,
            # and that is what happened without this line -- a tool that
            # requires no capability never greys out, so it stays pressable
            # with no session and fails at the cmdlet instead of at the button.
            '\b(?:Get|Export|Import)-TransportRule[A-Za-z]*'
        )
    }
    [PSCustomObject]@{
        Id       = 'SPO'
        Label    = 'SharePoint Online'
        Patterns = @(
            '\b(?:Get|Set|Add|New|Remove|Connect|Disconnect|Test|Request)-SPO[A-Za-z]*'
            '\bConnect-PnPOnline\b'
            '\bGet-PnP[A-Za-z]+'
        )
    }
    [PSCustomObject]@{
        Id       = 'EMS'
        Label    = 'Exchange on-premises'
        Patterns = @(
            # Deliberately only the on-prem-EXCLUSIVE families. Get-Mailbox and
            # friends exist in both worlds and resolve by module load order, so
            # they prove nothing about which one is present.
            '\b(?:Get|Set|New|Remove|Enable|Disable)-RemoteMailbox\b'
            '\bGet-ExchangeServer\b'
            '\bUpdate-Recipient\b'
            # Set-ADServerSettings pins the Exchange shell to a domain
            # controller and exists ONLY on-premises. It is the marker for
            # tools that drive on-prem group membership: Add- and
            # Remove-DistributionGroupMember are exported by both worlds, so
            # they identify nothing, and a tool relying on them alone scans as
            # needing no connection at all and never greys out.
            '\bSet-ADServerSettings\b'
        )
    }
    [PSCustomObject]@{
        Id       = 'AD'
        Label    = 'On-premises Active Directory'
        Patterns = @(
            '\b(?:Get|Set|New|Remove|Add|Enable|Disable|Unlock|Search|Move|Rename)-AD[A-Za-z]+'
        )
    }
)

# Compiled once and reused. A [regex] built with the Compiled option costs
# more to construct and far less to run, which is the right trade for something
# invoked against every script in Tools.json.
$script:LineCommentRegex = New-Object System.Text.RegularExpressions.Regex(
    '^(?:[^''"#]|''(?:[^'']*''|[^'']*$)|"(?:[^"]*"|[^"]*$))*',
    [System.Text.RegularExpressions.RegexOptions]::Compiled)

function Remove-PowerShellComment {
    <#
        Strips block comments, line comments and here-strings from script text so
        a cmdlet mentioned in prose is not mistaken for a dependency.

        Here-strings go because they carry HTML and JavaScript in the web tools,
        and a '#' inside a CSS colour is not a comment. Line comments are removed
        only when the '#' is not inside a quoted string, so a URL fragment or a
        hex colour in ordinary code survives.
    #>
    param([Parameter(Mandatory)][string]$Text)

    # Here-strings first: their contents must not be parsed as anything.
    $t = [regex]::Replace($Text, "@`"[\s\S]*?`"@", '', 'Singleline')
    $t = [regex]::Replace($t,    "@'[\s\S]*?'@",   '', 'Singleline')
    # Block comments.
    $t = [regex]::Replace($t, '<#[\s\S]*?#>', '', 'Singleline')

    # One regex per line instead of a character-by-character loop.
    #
    # The loop it replaces ran in interpreted PowerShell over every character
    # of every registered script -- 438,000 of them across 31 tools -- on the
    # UI thread, before the console's form painted. That was most of the
    # cold-launch delay.
    #
    # Reads as: from the start of the line, consume anything that is not a
    # quote or a hash, OR a complete quoted string, OR an unterminated quote
    # running to end of line. The match therefore stops at the first hash that
    # is NOT inside a string, which is exactly where the loop cut.
    #
    # The unterminated-quote alternatives are not decoration. A double-quoted
    # string may legally span lines in PowerShell, so a line can open a quote
    # and never close it; without those branches the match would stop at the
    # opening quote and silently truncate real code.
    #
    # Verified line-for-line identical to the old loop across all 47 scripts
    # in the toolkit before replacing it.
    $out = New-Object System.Text.StringBuilder
    foreach ($line in ($t -split "`n")) {
        [void]$out.AppendLine($script:LineCommentRegex.Match($line).Value)
    }
    return $out.ToString()
}

function Get-CapabilityCachePath {
    <#
        Where the capability scan results are remembered between windows.

        LOCALAPPDATA, deliberately NOT the toolkit folder. The toolkit is meant
        to run from a share, and a cache file beside it would be written by
        several operators at once -- contention over a file that exists purely
        to save a second. Per-user and per-machine also matches what the cache
        describes: this machine's copy of these scripts.
    #>
    $dir = Join-Path $env:LOCALAPPDATA 'M365AdminToolkit'
    if (-not (Test-Path -LiteralPath $dir)) {
        try { New-Item -Path $dir -ItemType Directory -Force | Out-Null } catch { return $null }
    }
    return (Join-Path $dir 'capability-cache.json')
}

function Import-ScriptCapabilityCache {
    <#
        Loads the on-disk cache into the in-process one, at most once.

        Every failure path here is a silent fall-through to an empty cache. A
        corrupt or unreadable cache must cost a rescan, never a broken console
        -- this is an optimisation, and an optimisation that can fail the
        program is not one.
    #>
    if ($global:ToolCapabilityCacheLoaded) { return }
    $global:ToolCapabilityCacheLoaded = $true
    if (-not $global:ToolCapabilityCache) { $global:ToolCapabilityCache = @{} }

    $path = Get-CapabilityCachePath
    if (-not $path -or -not (Test-Path -LiteralPath $path)) { return }
    try {
        $j = Get-Content -LiteralPath $path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        foreach ($p in $j.PSObject.Properties) {
            $global:ToolCapabilityCache[$p.Name] = [PSCustomObject]@{
                Path     = "$($p.Value.Path)"
                Found    = [bool]$p.Value.Found
                # @() matters: ConvertFrom-Json hands back a bare string for a
                # one-element array, and every caller does .Count on this.
                Requires = @($p.Value.Requires)
            }
        }
    } catch { }
}

function Save-ScriptCapabilityCache {
    <#
        Writes the cache back, only if something was actually scanned.

        Called once by the console after it has built its buttons rather than
        on every entry, so a cold start is one write instead of thirty-one.
        A tool that never calls this loses nothing but the saving.

        Keys carry the file's last-write ticks, so an edited script simply
        misses and rescans. Entries whose key no longer matches a current file
        are dropped on save, which is what stops the file growing forever as
        scripts are revised.
    #>
    if (-not $global:ToolCapabilityCacheDirty) { return }
    $path = Get-CapabilityCachePath
    if (-not $path) { return }

    try {
        $keep = @{}
        foreach ($k in $global:ToolCapabilityCache.Keys) {
            $parts = "$k" -split '\|'
            if ($parts.Count -ne 2) { continue }
            try {
                $fi = Get-Item -LiteralPath $parts[0] -ErrorAction Stop
                if ("$($fi.LastWriteTimeUtc.Ticks)" -eq $parts[1]) {
                    $keep[$k] = $global:ToolCapabilityCache[$k]
                }
            } catch { }
        }
        $json = [PSCustomObject]$keep | ConvertTo-Json -Depth 4
        $enc  = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($path, ($json -replace "`r?`n", "`r`n"), $enc)
        $global:ToolCapabilityCacheDirty = $false
    } catch { }
}

# Cmdlets that CREATE a connection rather than consume one. A script calling
# any of these provides that capability, so it is never greyed for lacking it.
#
# ANCHORED AT STATEMENT POSITION -- (?m)^\s* -- and that is the whole trick.
# Remove-PowerShellComment strips comments but NOT string literals, and several
# report tools carry an error message reading "...run Connect-SPOService
# first." An unanchored pattern matched that sentence and marked those tools as
# providing SharePoint, which would have left them enabled with no connection
# behind them. A mention is not a call; only a call starts a line.
$script:CapabilityProvides = @(
    # The optional '$x =' allows for the assignment form -- Connect-ExSE does
    # '$session = New-PSSession @params', which a bare start-of-line anchor
    # would miss, leaving the on-premises connect button greyed for wanting the
    # connection it creates.
    @{ Id = 'ExO'; Pattern = '(?m)^\s*(\$\w+\s*=\s*)?Connect-ExchangeOnline\b' }
    @{ Id = 'SPO'; Pattern = '(?m)^\s*(\$\w+\s*=\s*)?Connect-SPOService\b' }
    @{ Id = 'EMS'; Pattern = '(?m)^\s*(\$\w+\s*=\s*)?(Connect-ExchangeServer|New-PSSession)\b' }
)

function Get-ScriptCapabilityRequirement {

    <#
        Which capabilities a script needs, inferred from the cmdlets it calls.

        Cached by path AND last-write time, so editing a script re-scans it but
        a relaunch does not re-read two dozen files. The cache is global rather
        than $script: because the console and the tools it launches share a
        process but not a script scope (section 10), and it is BACKED BY DISK so
        a brand-new PowerShell window does not pay the scan either -- which is
        what it used to do, every time, on the UI thread before the form
        painted.
    #>
    param([Parameter(Mandatory)][string]$Path)

    Import-ScriptCapabilityCache

    try {
        $fi = Get-Item -LiteralPath $Path -ErrorAction Stop
    } catch {
        # A missing script is not an unmet capability -- it is a broken
        # Tools.json entry, which the console reports separately.
        return [PSCustomObject]@{ Path = $Path; Found = $false; Requires = @() }
    }

    $key = '{0}|{1}' -f $fi.FullName, $fi.LastWriteTimeUtc.Ticks
    if ($global:ToolCapabilityCache.ContainsKey($key)) {
        return $global:ToolCapabilityCache[$key]
    }

    $needs = @()
    try {
        $code = Remove-PowerShellComment -Text (Get-Content -LiteralPath $fi.FullName -Raw -ErrorAction Stop)
        foreach ($cap in $script:CapabilityDefs) {
            foreach ($p in $cap.Patterns) {
                if ([regex]::IsMatch($code, $p, 'IgnoreCase')) { $needs += $cap.Id; break }
            }
        }

        # A script that ESTABLISHES a connection cannot be gated on already
        # having it. Connect-ExO.ps1 calls Connect-ExchangeOnline, the scanner
        # read that as "needs Exchange Online", and the console greyed out the
        # one button that could have fixed it -- you could not connect because
        # you were not connected.
        #
        # Providing a capability outranks appearing to need it. Written as a
        # pattern rather than a filename so any future connect script inherits
        # this without anyone remembering to add it.
        foreach ($prov in $script:CapabilityProvides) {
            if ([regex]::IsMatch($code, $prov.Pattern, 'IgnoreCase')) {
                $needs = @($needs | Where-Object { $_ -ne $prov.Id })
            }
        }
    } catch { }

    $result = [PSCustomObject]@{
        Path     = $fi.FullName
        Found    = $true
        Requires = @($needs)
    }
    $global:ToolCapabilityCache[$key] = $result
    $global:ToolCapabilityCacheDirty = $true
    return $result
}

#region Tool help ---------------------------------------------------------------
<#
    A one-screen answer to "what does this tool do".

    Not documentation. The ScriptNotes files are documentation and are written
    for whoever maintains the toolkit. This is for the person who just clicked
    a button called "Returning User" and would like to know what happens next,
    before it happens.

    Two rules shape the design:

    1. THE TEXT LIVES IN THE SCRIPT. A Help.json would drift the first time
       somebody edited a script without editing the config, and would not
       survive the script being copied into another tenant's toolkit. A marker
       block in the header travels with the code:

           # TOOLKIT-HELP-BEGIN
           # Restores a user who was disabled in error.
           # TOOLKIT-HELP-END

    2. THE "RUNS AGAINST" LINE IS DERIVED, NOT TYPED. Get-ScriptCapabilityRequirement
       already classifies every script by which cmdlets it calls -- that is what
       greys the console buttons out. Reusing it means the help and the grey-out
       cannot disagree, and the line cannot go stale when the script changes.
       Typing it by hand would create a second source of truth for one fact,
       which is how the DC-selector and the write-scope bugs both happened.
#>

#region PIM activation timer ----------------------------------------------------
<#
    A local record of when PIM was last activated, and therefore how long is
    left before it has to be done again.

    THIS IS A NOTE, NOT A FACT. Entra is not consulted -- the tenant will not
    consent to the Graph permissions that would make it authoritative, and the
    portal does not put the Active-assignments tab in a URL either. What it
    knows is when the Activate button was pressed, plus a fixed duration.

    That is honest and useful because activation always goes through the
    button. If it ever does not, the countdown is wrong -- so the display
    always shows WHEN it was recorded, never a bare number. A stale timestamp
    is visibly stale; a bare "2h left" would not be.

    Stored per user under LOCALAPPDATA, NOT in Config\. Config ships to other
    engineers, and this is personal, machine-local state -- shipping it would
    hand someone else a countdown for an activation they never made.
#>

function Get-PimStatePath {
    $dir = Join-Path $env:LOCALAPPDATA 'M365AdminToolkit'
    if (-not (Test-Path -LiteralPath $dir)) {
        try { New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null } catch { }
    }
    return (Join-Path $dir 'pim-activation.json')
}

function Get-PimDurationHours {
    # 4 hours is this tenant's PIM policy. In Settings so a tenant with a
    # different maximum does not need a code change.
    $s = $null
    try { $s = Get-ToolkitSettings } catch { }
    if ($s -and $s.PSObject.Properties['PimDurationHours']) {
        $h = 0
        if ([int]::TryParse("$($s.PimDurationHours)", [ref]$h) -and $h -gt 0) { return $h }
    }
    return 4
}

function Get-PimActivation {
    <#
        When activation was last recorded, or $null.

        ONE CLOCK, not two. Groups and roles were tracked separately because
        they are activated separately, but the number that mattered was always
        the shorter of the two -- and two countdowns in a header is two things
        to read where one decision gets made: activate again, or don't.

        MIGRATION: a state file written by the two-clock version has Groups
        and/or Roles and no Activated. The LATER of the pair is taken rather
        than discarded, so upgrading mid-session does not silently reset a
        countdown that is still running.
    #>
    $p = Get-PimStatePath
    if (-not (Test-Path -LiteralPath $p)) { return $null }

    $parse = {
        param([string]$Text)
        if (-not "$Text".Trim()) { return $null }
        $dt = [DateTimeOffset]::MinValue
        # Round-trip format, parsed as round-trip. Parsing a stored timestamp
        # with local-culture rules is how a file written in one region becomes
        # unreadable in another.
        if ([DateTimeOffset]::TryParse("$Text", [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::RoundtripKind, [ref]$dt)) { return $dt }
        return $null
    }

    try {
        $j = Get-Content -LiteralPath $p -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    } catch { return $null }

    if ($j.PSObject.Properties['Activated']) {
        $v = & $parse "$($j.Activated)"
        if ($v) { return $v }
    }

    $legacy = @()
    foreach ($k in 'Groups', 'Roles') {
        if ($j.PSObject.Properties[$k]) {
            $v = & $parse "$($j.$k)"
            if ($v) { $legacy += $v }
        }
    }
    if ($legacy.Count -eq 0) { return $null }
    return (@($legacy | Sort-Object)[-1])
}

function Set-PimActivated {
    <# Records activation as of now, and returns the moment recorded. #>
    $at = [DateTimeOffset]::Now
    $p  = Get-PimStatePath
    try {
        # Only Activated is written. The legacy keys are not carried forward --
        # a file holding both would let the old countdown reappear if an older
        # copy of the toolkit read it.
        [PSCustomObject]@{ Activated = $at.ToString('o') } |
            ConvertTo-Json | Set-Content -LiteralPath $p -Encoding UTF8 -ErrorAction Stop
    } catch { }
    return $at
}

function Clear-PimActivation {
    <# Forgets the recorded activation. #>
    $p = Get-PimStatePath
    try {
        [PSCustomObject]@{ Activated = $null } |
            ConvertTo-Json | Set-Content -LiteralPath $p -Encoding UTF8 -ErrorAction Stop
    } catch { }
}

function Get-PimRemaining {
    <#
        Time left, or $null if never recorded. NEGATIVE when expired -- the
        caller decides how to say so, and "expired 20 minutes ago" is more
        useful than "expired".
    #>
    $at = Get-PimActivation
    if (-not $at) { return $null }
    return ($at.AddHours((Get-PimDurationHours)) - [DateTimeOffset]::Now)
}

function Format-PimSpan {
    <#
        [Math]::Floor, NOT [int].

        The PowerShell [int] cast ROUNDS -- and rounds half to even -- it does
        not truncate. [int](239 / 60) is 4, so 3h59m displayed as "4h 59m" and
        looked like a five-hour timer on a four-hour policy. Reported
        2026-08-07 as "it's counting down from 5 hours".

        Anywhere a total is split into units, the quotient must be floored.
    #>
    param([TimeSpan]$Span)
    $m = [int][Math]::Floor([Math]::Abs($Span.TotalMinutes))
    $h = [int][Math]::Floor($m / 60); $m = $m % 60
    if ($h -gt 0) { return "{0}h {1:00}m" -f $h, $m }
    return "{0}m" -f $m
}

function Get-PimStatusLine {
    <#
        One short line for the console header. Empty string when nothing has
        ever been recorded, so the label can simply hide rather than showing
        a meaningless zero state.
    #>
    $r = Get-PimRemaining
    if ($null -eq $r) { return '' }
    if ($r.TotalSeconds -le 0) { return 'PIM: EXPIRED' }
    return "PIM: $(Format-PimSpan -Span $r)"
}

function Get-PimStatusColour {
    <#
        Green with time to spare, amber under 30 minutes, red once expired.
        Thirty minutes because that is roughly the point at which starting
        something long becomes a bad idea.
    #>
    $r = Get-PimRemaining
    if ($null -eq $r) { return [System.Drawing.Color]::DimGray }
    if ($r.TotalSeconds -le 0)  { return [System.Drawing.Color]::FromArgb(255, 105, 97) }
    if ($r.TotalMinutes -le 30) { return [System.Drawing.Color]::FromArgb(230, 170, 60) }
    return [System.Drawing.Color]::FromArgb(80, 220, 130)
}

#endregion

#region Version and integrity ---------------------------------------------------
<#
    Two different questions, and they need two different answers.

    "WHAT VERSION ARE YOU ON?" is answered by a number -- Version.json at the
    toolkit root.

    "DOES YOUR COPY MATCH MINE?" is NOT answered by that number. A version
    string is a CLAIM: edit a script, forget to bump it, and the toolkit now
    reports a version it is not. The only honest answer is a per-file hash,
    generated from the master and compared on the copy. A hash cannot be wrong
    by omission -- forgetting to update it makes the check FAIL, which is the
    right direction to fail in.

    So Version.json carries both: one version, and a SHA-256 per code file.

    WHAT IS HASHED: .ps1 files, at the toolkit root and under Projects.
    WHAT IS NOT:    Config\ (Tools.json, Settings.json and friends are meant to
                    differ per engineer -- hashing them would report drift on
                    every machine and train everyone to ignore the result),
                    reports, logs, exports, and Version.json itself.

    Regenerate on the master with Publish-Toolkit.ps1 after any code
    change. That is a discipline, and the check compensates for slips: it
    reports the manifest as STALE when any .ps1 is newer than the manifest.
#>

function Get-ToolkitVersionPath {
    return (Join-Path $PSScriptRoot 'Version.json')
}

function Get-ToolkitVersion {
    <#
        Returns the version object, or a clearly-marked unknown. Never throws:
        a missing manifest is the normal state of a working copy that has not
        been stamped yet, not an error worth stopping the console for.
    #>
    $out = [PSCustomObject]@{
        Version  = 'unversioned'
        Released = $null
        Files    = @{}
        Found    = $false
    }
    $p = Get-ToolkitVersionPath
    if (-not (Test-Path -LiteralPath $p)) { return $out }
    try {
        $j = Get-Content -LiteralPath $p -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        $files = @{}
        if ($j.PSObject.Properties['Files'] -and $j.Files) {
            foreach ($n in $j.Files.PSObject.Properties.Name) { $files[$n] = "$($j.Files.$n)" }
        }
        return [PSCustomObject]@{
            Version  = if ($j.PSObject.Properties['Version']) { "$($j.Version)" } else { 'unversioned' }
            Released = if ($j.PSObject.Properties['Released']) { "$($j.Released)" } else { $null }
            Files    = $files
            Found    = $true
        }
    } catch { return $out }
}

function Get-ManifestExcludedFile {
    <#
        MASTER-ONLY TOOLS, LEFT OUT OF THE MANIFEST ON PURPOSE.

        These two exist only on the master and Deploy.ps1 does not copy them.
        While the manifest listed them, the master's Version.json and a
        deployed copy's described DIFFERENT file lists -- so the master's could
        never be handed over directly. Every patch had to travel with a manifest
        generated by a full deploy, or the recipient's self-test opened with two
        MISSING rows for files their copy is not supposed to have.

        Excluding them makes the two manifests identical, which is what makes
        "here is the changed script and here is Version.json" a workable way to
        ship a fix.

        THE COST, STATED RATHER THAN ASSUMED: Deploy.ps1 and Publish-Toolkit.ps1
        are no longer covered by tamper detection. An edit to either goes
        unreported. That is acceptable because there is exactly one master and
        it is the machine of the person who would be making the edit -- but it
        IS a hole, and it is here rather than buried in a filter expression.

        A FUNCTION, NOT A $script: VARIABLE. Deploy.ps1 dot-sources Common.ps1
        and needs the same answer; a function always resolves in the file that
        defined it, while $script: state has already caught this toolkit out
        once when it was read from a different scope.
    #>
    return @('Publish-Toolkit.ps1', 'Deploy.ps1')
}

function Get-ToolkitCodeFile {
    <#
        The files the manifest covers. One definition, used by both the
        generator and the checker -- if these two ever disagreed, every copy
        would report drift that is not there.

        The exclusion is applied HERE, which means the excluded files are absent
        from the on-disk list as well as from the manifest. That symmetry is the
        whole point: filtering only the generator would leave them on disk and
        unlisted, and Test-ToolkitIntegrity would report them as EXTRA -- the
        same false alarm in the other direction.
    #>
    $root = $PSScriptRoot
    $skip = Get-ManifestExcludedFile
    $list = New-Object System.Collections.ArrayList
    foreach ($f in (Get-ChildItem -LiteralPath $root -Filter '*.ps1' -File -ErrorAction SilentlyContinue)) {
        if ($skip -contains $f.Name) { continue }
        [void]$list.Add($f.FullName)
    }
    $proj = Join-Path $root 'Projects'
    if (Test-Path -LiteralPath $proj) {
        foreach ($f in (Get-ChildItem -LiteralPath $proj -Filter '*.ps1' -File -Recurse -ErrorAction SilentlyContinue)) {
            [void]$list.Add($f.FullName)
        }
    }
    return @($list.ToArray() | Sort-Object)
}

function Get-ToolkitRelativePath {
    param([Parameter(Mandatory)][string]$FullPath)
    $root = $PSScriptRoot.TrimEnd('\')
    if ($FullPath.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) {
        return $FullPath.Substring($root.Length).TrimStart('\')
    }
    return (Split-Path -Leaf $FullPath)
}

function Get-FileHashHex {
    param([Parameter(Mandatory)][string]$Path)
    try { return (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash.ToUpperInvariant() }
    catch { return $null }
}

function Test-ToolkitIntegrity {
    <#
        Compares this copy against the manifest.

        Returns counts plus the actual lists, because "3 files differ" is a
        headline and "which three" is the answer.
    #>
    $v = Get-ToolkitVersion
    $modified = New-Object System.Collections.ArrayList
    $missing  = New-Object System.Collections.ArrayList
    $extra    = New-Object System.Collections.ArrayList
    $matched  = 0
    $stale    = $false

    if (-not $v.Found) {
        return [PSCustomObject]@{
            Version = $v.Version; Released = $v.Released; HasManifest = $false
            Matched = 0; Modified = @(); Missing = @(); Extra = @(); Stale = $false
            Summary = 'No Version.json -- this copy is not stamped.'
        }
    }

    $onDisk = @{}
    foreach ($f in (Get-ToolkitCodeFile)) { $onDisk[(Get-ToolkitRelativePath -FullPath $f)] = $f }

    foreach ($rel in $v.Files.Keys) {
        if (-not $onDisk.ContainsKey($rel)) { [void]$missing.Add($rel); continue }
        $h = Get-FileHashHex -Path $onDisk[$rel]
        if ($h -and $h -eq $v.Files[$rel]) { $matched++ } else { [void]$modified.Add($rel) }
    }
    foreach ($rel in $onDisk.Keys) {
        if (-not $v.Files.ContainsKey($rel)) { [void]$extra.Add($rel) }
    }

    # Timestamps, kept as an EXPLANATION and never as a verdict.
    #
    # A file newer than the manifest but hashing the same is byte-for-byte
    # identical -- touched, not changed. Treating that as a problem would turn
    # a perfect copy amber on any copy method that resets timestamps, which is
    # exactly the false alarm this region's header warns against. So Stale
    # never fails a clean check on its own; it only explains a failing one
    # ("that is why -- the master was edited and not re-stamped").
    try {
        $mTime = (Get-Item -LiteralPath (Get-ToolkitVersionPath)).LastWriteTimeUtc
        foreach ($f in $onDisk.Values) {
            if ((Get-Item -LiteralPath $f).LastWriteTimeUtc -gt $mTime) { $stale = $true; break }
        }
    } catch { }

    $summary =
        if ($modified.Count -eq 0 -and $missing.Count -eq 0 -and $extra.Count -eq 0) {
            "v$($v.Version) -- all $matched files match"
        } else {
            "v$($v.Version) -- $($modified.Count) modified, $($missing.Count) missing, $($extra.Count) extra"
        }
    if ($stale -and ($modified.Count -or $missing.Count -or $extra.Count)) {
        $summary += '  (manifest is older than the code -- re-stamp the master)'
    }

    return [PSCustomObject]@{
        Version = $v.Version; Released = $v.Released; HasManifest = $true
        Matched = $matched
        Modified = @($modified.ToArray()); Missing = @($missing.ToArray()); Extra = @($extra.ToArray())
        Stale = $stale
        Summary = $summary
    }
}
#endregion

function Get-ToolHelpText {
    <#
        Reads the TOOLKIT-HELP block out of a script.

        Returns $null when there is none, so callers can decide whether to
        show a button at all rather than showing an empty dialog.
    #>
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try {
        $lines = Get-Content -LiteralPath $Path -ErrorAction Stop
    } catch { return $null }

    $out = New-Object System.Collections.ArrayList
    $inBlock = $false
    foreach ($l in $lines) {
        if ($l -match 'TOOLKIT-HELP-BEGIN') { $inBlock = $true; continue }
        if ($l -match 'TOOLKIT-HELP-END')   { break }
        if ($inBlock) {
            # Strip the leading '#' and ONE following space. Stripping all
            # leading whitespace would flatten deliberate indentation, and
            # these blocks use it for bullets.
            $text = $l -replace '^\s*#', ''
            if ($text.StartsWith(' ')) { $text = $text.Substring(1) }
            [void]$out.Add($text)
        }
    }
    if ($out.Count -eq 0) { return $null }
    return (($out -join "`r`n").Trim())
}

<#
    A SECOND, stricter classification -- used only by the help line.

    $script:CapabilityDefs answers "might this script need X?", and is
    deliberately generous: over-claiming makes the console grey a button that
    would have worked, which is a minor annoyance. Under-claiming lets someone
    click a tool that then fails, which is worse. Generous is correct there.

    Stating a fact to a person is the opposite problem, and reusing the
    generous list produced two wrong answers on the first test:

      * Bulk-DLMembership.ps1 claimed Exchange Online, because Get-DistributionGroup
        is in the ExO pattern list. That cmdlet exists in BOTH worlds -- it is
        the exact ambiguity that cost most of 2026-08-06 -- and the tool is
        on-premises only.
      * HeaderAnalyzer.ps1 claimed Active Directory, because it contained a
        function called Get-AddressDomain and 'Get-AD' + 'dressDomain' matches
        the AD cmdlet pattern. The tool connects to nothing at all.

    So this list contains only markers that are EXCLUSIVE to one world, plus a
    whitelist of real AD nouns rather than an open-ended -AD* pattern. Anything
    ambiguous is reported as ambiguous, which is the honest answer and is also
    useful: "Exchange, whichever you are connected to" is a true statement about
    a tool that calls Get-Mailbox and nothing more specific.
#>
$script:HelpCapabilityDefs = @(
    [PSCustomObject]@{
        Id       = 'ExO'
        Label    = 'Exchange Online'
        Patterns = @(
            '\b(?:Get|Set)-EXO[A-Za-z]+'
            '\bConnect-ExchangeOnline\b'
            '\bGet-ConnectionInformation\b'
            '\bInvoke-EXOCommand\b'
            '\bGet-UnifiedGroup[A-Za-z]*'
            '\bGet-MessageTrace[A-Za-z0-9]*'
            '\b(?:Get|Release|Delete|Preview|Export)-QuarantineMessage[A-Za-z0-9]*'
            '\bStart-HistoricalSearch\b'
        )
    }
    [PSCustomObject]@{
        Id       = 'SPO'
        Label    = 'SharePoint Online'
        Patterns = @(
            '\b(?:Get|Set|Add|New|Remove|Connect|Disconnect|Test|Request)-SPO[A-Za-z]*'
            '\bConnect-PnPOnline\b'
            '\bGet-PnP[A-Za-z]+'
        )
    }
    [PSCustomObject]@{
        Id       = 'EMS'
        Label    = 'Exchange on-premises'
        Patterns = @(
            '\b(?:Get|Set|New|Remove|Enable|Disable)-RemoteMailbox\b'
            '\bGet-ExchangeServer\b'
            '\bGet-ExchangeCertificate\b'
            '\bSet-ADServerSettings\b'
            '\bUpdate-Recipient\b'
            '\bInvoke-OnPremCommand\b'
            # Exchange cmdlets despite the AD prefix, and on-premises only.
            '\b(?:Get|Add|Remove)-ADPermission\b'
        )
    }
    [PSCustomObject]@{
        Id       = 'AD'
        Label    = 'On-premises Active Directory'
        Patterns = @(
            # Whitelisted nouns, not -AD[A-Za-z]+. The open form matches any
            # local function whose name happens to begin Get-AD, which is how
            # an offline tool ended up claiming a directory connection.
            '\b(?:Get|Set|New|Remove|Add|Enable|Disable|Unlock|Search|Move|Rename)-AD(?:User|Group|GroupMember|PrincipalGroupMembership|Object|Computer|OrganizationalUnit|Domain|DomainController|RootDSE|Account|AccountPassword)\b'
        )
    }
    [PSCustomObject]@{
        Id       = 'Graph'
        Label    = 'Microsoft Entra (Microsoft Graph)'
        Patterns = @(
            # Added 2026-08-06 after PimActivation.ps1 reported "connects to no
            # directory and sends nothing anywhere" -- while talking to Graph.
            #
            # An absent capability does not read as "unknown", it reads as
            # "nothing", and "nothing" is the most reassuring sentence this
            # tool can print. A missing entry is therefore not a gap, it is a
            # false statement of safety. Anything added to the toolkit that
            # reaches a new service needs a row here on the same day.
            '\bConnect-MgGraph\b'
            '\bInvoke-MgGraphRequest\b'
            '\bGet-Mg[A-Z][A-Za-z]*'
            'graph\\.microsoft\\.com'
        )
    }
)

# Cmdlets that exist identically in BOTH Exchanges. Their presence proves the
# script talks to Exchange; it proves nothing about WHICH.
$script:AmbiguousExchangePattern =
    '\b(?:Get|Set|New|Remove|Add|Enable|Disable)-(?:Mailbox|MailboxPermission|MailboxFolderPermission|MailboxStatistics|CASMailbox|Recipient|RecipientPermission|DistributionGroup|DistributionGroupMember|MailContact|MailUser|InboxRule|AcceptedDomain|OrganizationConfig|MailboxAutoReplyConfiguration)\b'

function Get-ToolRunsAgainstText {
    <#
        The derived first line: which directories this script actually talks to.

        Derived rather than typed so it cannot go stale when the script changes,
        and so it cannot disagree with itself. See $script:HelpCapabilityDefs
        for why it does NOT reuse the grey-out classification.
    #>
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return 'Runs against: unknown -- the script could not be read.'
    }
    try {
        $raw  = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        $code = Remove-PowerShellComment -Text $raw
    } catch {
        return 'Runs against: unknown -- the script could not be read.'
    }

    # An explicit override, for the case derivation genuinely cannot answer:
    #
    #     # TOOLKIT-HELP-RUNS: <text>
    #
    # This is NOT a licence to type what could be derived -- the whole point of
    # deriving is that it cannot go stale. It exists because AdminConsole.ps1
    # is a launcher: it runs against whatever you connect it to, and scanning
    # its own code describes the launcher's plumbing rather than the answer to
    # the question. Use it where the truth is not in the file's cmdlets, and
    # nowhere else.
    if ($raw -match '(?im)^\s*#\s*TOOLKIT-HELP-RUNS:\s*(.+?)\s*$') {
        return "Runs against: $($Matches[1])"
    }

    $labels = New-Object System.Collections.ArrayList
    $ids    = New-Object System.Collections.ArrayList
    foreach ($cap in $script:HelpCapabilityDefs) {
        foreach ($p in $cap.Patterns) {
            if ([regex]::IsMatch($code, $p, 'IgnoreCase')) {
                [void]$labels.Add("$($cap.Label)"); [void]$ids.Add("$($cap.Id)"); break
            }
        }
    }

    # A script that declares its own process is on-premises by construction --
    # that declaration exists precisely because it must not share a process
    # with the cloud. It settles the ambiguous case without guessing.
    $declaresOnPrem = [regex]::IsMatch($code, 'TOOLKIT-REQUIRES:\s*OwnProcess', 'IgnoreCase')
    if (-not $declaresOnPrem) {
        try {
            $head = (Get-Content -LiteralPath $Path -TotalCount 40 -ErrorAction Stop) -join "`n"
            $declaresOnPrem = [regex]::IsMatch($head, 'TOOLKIT-REQUIRES:\s*OwnProcess', 'IgnoreCase')
        } catch { }
    }

    $ambiguous = [regex]::IsMatch($code, $script:AmbiguousExchangePattern, 'IgnoreCase')
    if ($ambiguous -and -not ($ids -contains 'ExO') -and -not ($ids -contains 'EMS')) {
        if ($declaresOnPrem) {
            [void]$labels.Add('Exchange on-premises')
        } else {
            # Honest, and genuinely informative: these cmdlets go wherever the
            # session points, which is a real property of the tool rather than
            # a gap in the analysis.
            [void]$labels.Add('Exchange (on-premises or Online -- whichever this session is connected to)')
        }
    }

    if ($labels.Count -eq 0) {
        return 'Runs against: nothing -- this tool connects to no directory and sends nothing anywhere.'
    }
    return ('Runs against: ' + (($labels | Select-Object -Unique) -join ', '))
}

function Show-ToolHelp {
    <#
        The dialog. Derived line, a rule, then the script's own words.

        Read-only multiline TextBox rather than a MessageBox: help that runs
        past a few lines gets a scrollbar instead of a dialog taller than the
        screen, and the text stays selectable so it can be pasted into a
        ticket.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Title = ''
    )

    $help = Get-ToolHelpText -Path $Path
    $runs = Get-ToolRunsAgainstText -Path $Path
    if (-not $Title) { $Title = [System.IO.Path]::GetFileNameWithoutExtension($Path) }

    if (-not $help) {
        $help = "No overview has been written for this tool yet.`r`n`r`nAdd one between TOOLKIT-HELP-BEGIN and TOOLKIT-HELP-END in the script header and it will appear here."
    }

    $body = "$runs`r`n" + ('-' * 68) + "`r`n`r`n$help"

    $f               = New-Object System.Windows.Forms.Form
    $f.Text          = "About: $Title"
    $f.Size          = New-Object System.Drawing.Size(620, 460)
    $f.StartPosition = 'CenterParent'
    $f.MinimizeBox   = $false
    $f.MaximizeBox   = $false
    $f.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $tb            = New-Object System.Windows.Forms.TextBox
    $tb.Multiline  = $true
    $tb.ReadOnly   = $true
    $tb.ScrollBars = 'Vertical'
    $tb.WordWrap   = $true
    $tb.Location   = New-Object System.Drawing.Point(12, 12)
    $tb.Size       = New-Object System.Drawing.Size(580, 360)
    $tb.Font       = New-Object System.Drawing.Font('Consolas', 9)
    $tb.BackColor  = [System.Drawing.Color]::WhiteSmoke
    $tb.ForeColor  = [System.Drawing.Color]::Black
    $tb.Tag        = 'NoTheme'
    $tb.Text       = $body
    # Caret at the top. A ReadOnly TextBox set after creation otherwise opens
    # scrolled to the end on longer text.
    $tb.Select(0, 0)
    $f.Controls.Add($tb)

    $ok          = New-Object System.Windows.Forms.Button
    $ok.Text     = 'Close'
    $ok.Location = New-Object System.Drawing.Point(492, 382)
    $ok.Size     = New-Object System.Drawing.Size(100, 30)
    # Resolved from the sender, not from a captured $f. Same reasoning as
    # above -- no variable capture, nothing to rebind, nothing to lose.
    $ok.Add_Click({ $this.FindForm().Close() })
    $f.Controls.Add($ok)
    $f.AcceptButton = $ok
    $f.CancelButton = $ok

    Set-ConsoleTheme -Form $f
    [void]$f.ShowDialog()
    $f.Dispose()
}

function Add-ToolHelpButton {
    <#
        Puts a '?' in the top-right corner of a form, in the same place on
        every tool.

        Consistency is the entire point -- a help button you have to look for
        is one nobody presses. It anchors Top+Right so it stays in the corner
        if the form is resized, and it is added LAST so it draws above any
        control it overlaps (WinForms z-order is creation order: earlier is
        on top).

        Tag carries 'NoTheme' plus the script path and title, pipe-separated.
        The theme's opt-out is a substring test, so the button still keeps its
        own colours while Tag doubles as the handler's data channel.
    #>
    param(
        [Parameter(Mandatory)]$Form,
        [Parameter(Mandatory)][string]$ScriptPath,
        [string]$Title = ''
    )

    $btn          = New-Object System.Windows.Forms.Button
    $btn.Text     = '?'
    $btn.Name     = 'btnToolHelp'
    $btn.Size     = New-Object System.Drawing.Size(28, 26)

    <#
        Finds a clear spot rather than assuming the corner is free.

        A static sweep on 2026-08-06 found the corner position sitting on real
        content on 15 of 40 forms -- a text box in HeaderAnalyzer, a grid in
        QuickLaunchTool, a button in Test-Toolkit, and the right-hand end of a
        status label on a dozen report tools. Fixing fifteen layouts by hand
        would fix today and not tomorrow.

        Done at run time on purpose. Only here can the ACTUAL rendered text
        width be measured; a source scan has to estimate it, and estimating it
        is what made the first two passes of that sweep wrong in both
        directions -- once by ignoring that a Label clips at its own Width, and
        once by counting an overlap with a label's empty right-hand end as a
        collision.

        A Label is an obstacle only as far as its text actually reaches.
        Everything else is an obstacle for its whole width, because anything
        can be drawn anywhere inside it.
    #>
    $y = 8
    $x = $Form.ClientSize.Width - 40
    foreach ($c in $Form.Controls) {
        if ("$($c.Tag)" -like '*ToolkitLogo*') { $x = $c.Left - 8 - $btn.Width; break }
    }

    $obstacles = New-Object System.Collections.ArrayList
    foreach ($c in $Form.Controls) {
        if ("$($c.Name)" -eq 'btnToolHelp') { continue }
        if ("$($c.Tag)" -like '*ToolkitLogo*') { continue }
        if (-not $c.Visible -and $c.GetType().Name -ne 'Label') { continue }

        $bounds = $c.Bounds
        if ($c -is [System.Windows.Forms.Label] -and -not $c.AutoSize) {
            try {
                $tw = [System.Windows.Forms.TextRenderer]::MeasureText($c.Text, $c.Font).Width
                if ($tw -lt $c.Width) {
                    $bounds = New-Object System.Drawing.Rectangle($c.Left, $c.Top, ([Math]::Max($tw, 1)), $c.Height)
                }
            } catch { }
        }
        [void]$obstacles.Add($bounds)
    }

    # Try the corner, then a short walk left.
    #
    # The walk rarely helps and that is worth knowing: on most of these forms
    # the top row is ONE label spanning the full width, so there is no
    # horizontal gap to find. It is kept for the few forms where a small nudge
    # does clear a button or a short label.
    $limit = $Form.ClientSize.Width - 200
    $start = $x
    $found = $false
    while ($x -gt $limit) {
        $rect = New-Object System.Drawing.Rectangle($x, $y, $btn.Width, $btn.Height)
        $clash = $false
        foreach ($o in $obstacles) { if ($rect.IntersectsWith($o)) { $clash = $true; break } }
        if (-not $clash) { $found = $true; break }
        $x -= 8
    }

    if (-not $found) {
        # No gap anywhere along the top row, so make one: widen the form and put
        # the button in the new margin.
        #
        # The alternatives were both worse. Sitting on top of a control hides
        # something the operator needs. Shrinking the offending label clips real
        # text -- and the labels involved are the ones already running to the
        # window edge, so they are exactly the ones with no slack to give.
        #
        # Widening only affects controls anchored Right, which follow the edge
        # as intended; absolutely positioned controls stay put and the new space
        # becomes an empty gutter, which is precisely what is wanted.
        $gutter = $btn.Width + 16
        $Form.ClientSize = New-Object System.Drawing.Size(
            ($Form.ClientSize.Width + $gutter), $Form.ClientSize.Height)
        $x = $Form.ClientSize.Width - $btn.Width - 12

        # A logo already placed sits at the old right edge; move it along.
        foreach ($c in $Form.Controls) {
            if ("$($c.Tag)" -like '*ToolkitLogo*') {
                $c.Left = $c.Left + $gutter
                $x = $c.Left - 8 - $btn.Width
            }
        }
    }

    $btn.Location = New-Object System.Drawing.Point($x, $y)
    $btn.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $btn.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif', 10, [System.Drawing.FontStyle]::Bold)
    $btn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btn.BackColor = [System.Drawing.Color]::FromArgb(52, 88, 148)
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(120, 170, 235)
    $btn.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(72, 118, 190)

    # NO .GetNewClosure(). It looks like the right way to capture $ScriptPath
    # and it is the reason the first version threw
    # "Show-ToolHelp is not recognized".
    #
    # GetNewClosure rebinds the scriptblock to a NEW dynamic module scope whose
    # parent is the global scope. Common.ps1 is dot-sourced into each tool's
    # SCRIPT scope, so its functions are not global -- and the closure
    # therefore cannot see them. A plain scriptblock keeps the session state it
    # was created in, where Show-ToolHelp is perfectly visible.
    #
    # The data rides on Tag instead. Set-ConsoleTheme's opt-out test is
    # "$($c.Tag) -like '*NoTheme*'", a substring match, so a composite string
    # still opts out -- which is what makes this slot usable rather than
    # 'taken'.
    $btn.Tag = "NoTheme|$ScriptPath|$Title"
    $btn.Add_Click({
        $parts = "$($this.Tag)" -split '\|', 3
        if ($parts.Count -ge 2) {
            Show-ToolHelp -Path $parts[1] -Title $(if ($parts.Count -ge 3) { $parts[2] } else { '' })
        }
    })

    $tip = New-Object System.Windows.Forms.ToolTip
    $tip.SetToolTip($btn, 'What does this tool do?')

    $Form.Controls.Add($btn)
    $btn.BringToFront()
    return $btn
}
#endregion

function Get-ExchangeModuleMap {
    <#
        Which loaded module owns the on-premises verbs, and which owns the
        cloud ones.

        Some scripts genuinely need BOTH Exchanges in one process -- the hybrid
        disable pipeline reads and writes on-premises AD and RemoteMailbox
        objects and then acts on the cloud mailbox. Isolating them into their
        own process (TOOLKIT-REQUIRES: OwnProcess) would break them, so they
        need the other answer: say which Exchange you mean, every time.

        Both worlds export Get-Mailbox, Set-Mailbox, Get-DistributionGroup and
        dozens more, and the last import wins the bare name. A hybrid script
        that calls the bare name is relying on connection ORDER -- which works
        right up until someone connects in the other order, and then writes
        land in the wrong directory with no error.

        Telling the modules apart, version-independently:

          on-premises -- exports Get-RemoteMailbox. The *-RemoteMailbox family
                         is on-prem-exclusive, which is the same property
                         Test-EmsAvailable relies on.
          cloud       -- exports Get-Mailbox but NOT Get-RemoteMailbox.

        Deliberately not matched on module NAME. Implicit remoting generates
        names like tmpEXO_a1b2c3 and they are not stable across versions or
        sessions.
    #>
    param([switch]$Refresh)

    if (-not $Refresh -and $script:ExchangeModuleMap) { return $script:ExchangeModuleMap }

    $onPrem = $null; $cloud = $null
    foreach ($m in @(Get-Module)) {
        $ec = $m.ExportedCommands
        if (-not $ec -or -not $ec.ContainsKey('Get-Mailbox')) { continue }
        if ($ec.ContainsKey('Get-RemoteMailbox')) {
            if (-not $onPrem) { $onPrem = $m }
        } else {
            if (-not $cloud) { $cloud = $m }
        }
    }

    $script:ExchangeModuleMap = [PSCustomObject]@{
        OnPrem = $onPrem
        Cloud  = $cloud
        Both   = [bool]($onPrem -and $cloud)
    }
    return $script:ExchangeModuleMap
}

function Invoke-ExchangeCommand {
    <#
        Run an Exchange cmdlet against a NAMED directory, not against whichever
        module happens to have won the name.

            Invoke-ExchangeCommand -World OnPrem -Command 'Set-Mailbox' -Params @{
                Identity = $smtp; HiddenFromAddressListsEnabled = $true
            }

        When only one Exchange is loaded this resolves to the same command the
        bare name would have. The value is entirely in the case where both are
        loaded -- and in the case where a reader has to work out, months later,
        which directory a line was aimed at. The answer is now in the line.

        Throws rather than falling back if the requested world is absent. A
        hybrid script that silently ran its on-premises half against the cloud
        is exactly the failure being prevented here, and a fallback would
        reintroduce it.
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('OnPrem', 'Cloud')][string]$World,
        [Parameter(Mandatory)][string]$Command,
        [hashtable]$Params = @{}
    )

    $map = Get-ExchangeModuleMap
    $mod = if ($World -eq 'OnPrem') { $map.OnPrem } else { $map.Cloud }

    if (-not $mod) {
        # Re-read once before giving up: a connect may have happened after the
        # map was first built, and the cache would still say absent.
        $map = Get-ExchangeModuleMap -Refresh
        $mod = if ($World -eq 'OnPrem') { $map.OnPrem } else { $map.Cloud }
    }
    if (-not $mod) {
        $what = if ($World -eq 'OnPrem') { 'on-premises Exchange' } else { 'Exchange Online' }
        throw "$Command was requested against $what, which is not connected in this process. Refusing to run it against the other directory."
    }

    $qualified = Get-Command -Name "$($mod.Name)\$Command" -ErrorAction SilentlyContinue
    if (-not $qualified) {
        throw "$Command is not exported by the $World Exchange module ($($mod.Name))."
    }
    & $qualified @Params
}

function Test-ExoInProcess {
    <#
        Is Exchange Online ALSO live in this process?

        This is not the same question as Test-EmsAvailable, and conflating the
        two cost three rounds of misdiagnosis on 2026-08-06.

        Test-EmsAvailable answers "are the on-premises cmdlets loaded". It can
        be TRUE while Exchange Online is loaded as well -- and the two share
        names. Get-DistributionGroup, Remove-DistributionGroupMember,
        Set-Mailbox and dozens more exist in both worlds. Whichever module was
        imported last owns the name, so a tool can report "connected to
        on-premises Exchange", genuinely have EMS available, and still have
        every call routed to the cloud.

        What it looks like when that happens: an on-premises object is found
        and identified correctly, and the write fails with

            The operation on Identity "X" failed because it's out of the
            current user's write scope

        which is Exchange Online refusing to modify a directory-synchronised
        object. The sentence continues "...because the object is being
        synchronized from your on-premises organization", and that tail is what
        gives it away -- so never truncate an error message in a log window.

        A tool that writes on-premises must therefore check BOTH: that EMS is
        available, and that EXO is not competing for the same verbs.
    #>
    # v3+ answers directly. This is the reliable path.
    $gci = Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue
    if ($gci) {
        try {
            $conns = @(Get-ConnectionInformation -ErrorAction Stop)
            $live  = @($conns | Where-Object { $_.State -eq 'Connected' -or -not $_.State })
            if ($live.Count -gt 0) {
                return [PSCustomObject]@{
                    Present = $true
                    Detail  = "Exchange Online is connected in this process ($(($live | ForEach-Object { $_.Organization }) -join ', ')). It shares cmdlet names with on-premises Exchange."
                }
            }
        } catch { }
    }

    # v2 and remote-session fallback: an implicit remoting session pointed at
    # the service. Matching on the URI rather than the name because the
    # generated session names are not stable.
    try {
        $ps = @(Get-PSSession -ErrorAction Stop | Where-Object {
            $_.ConfigurationName -eq 'Microsoft.Exchange' -and
            "$($_.Runspace.ConnectionInfo.ConnectionUri)" -match 'outlook\.office365\.com|outlook\.office\.com|ps\.compliance\.protection'
        })
        if ($ps.Count -gt 0) {
            return [PSCustomObject]@{
                Present = $true
                Detail  = 'An Exchange Online remoting session is open in this process. It shares cmdlet names with on-premises Exchange.'
            }
        }
    } catch { }

    return [PSCustomObject]@{ Present = $false; Detail = 'No Exchange Online connection in this process.' }
}

function Test-EmsAvailable {
    <#
        On-premises Exchange, tested via *-RemoteMailbox.

        That family exists ONLY on-prem, so its presence is unambiguous -- which
        is the same property that makes NewUser.ps1 safe to run beside an open
        Exchange Online session (section 2). Testing with Get-Mailbox would
        prove nothing, because both worlds export it.
    #>
    $c = Get-Command -Name 'New-RemoteMailbox' -ErrorAction SilentlyContinue
    if (-not $c) {
        return [PSCustomObject]@{
            Available = $false
            Detail    = 'Not available -- the *-RemoteMailbox cmdlets are not loaded in this process. Open an Exchange Management Shell window, or import the on-prem session.'
        }
    }

    # Loaded is not the same as working. Test-AdAvailable already learned this
    # and answers it with Get-ADRootDSE; without the equivalent here, an EMS
    # session whose server has gone away reports as connected and then fails on
    # the first Set-RemoteMailbox -- which in a disable batch means every user
    # fails at step one and the retry queue fills with the same misleading error.
    #
    # Get-RemoteMailbox is on-prem-exclusive, so this proves WHICH Exchange
    # answered as well as THAT one did. -ResultSize 1 keeps it cheap on an org
    # with thousands of mailboxes.
    #
    # -WarningAction SilentlyContinue is REQUIRED, not tidiness. Asking for one
    # row out of thousands makes Exchange say
    #
    #     WARNING: There are more results available than are currently
    #     displayed. To view them, increase the value for the ResultSize
    #     parameter.
    #
    # which is true, expected, and exactly what was asked for -- but it prints
    # on the host stream at every launch and every connection repaint, so it
    # reads as a fault in the console rather than as Exchange answering the
    # question correctly. The warning cannot be avoided by raising ResultSize;
    # it is caused by LIMITING the query at all, so the only fix is to silence
    # this one call. See ARCHITECTURE section 14b -- this is the opposite trap
    # to the one Get-AcceptedDomain has, and confusing the two cost a day.
    try {
        $null = Get-RemoteMailbox -ResultSize 1 -ErrorAction Stop -WarningAction SilentlyContinue
        return [PSCustomObject]@{ Available = $true; Detail = 'Exchange on-premises is loaded and answering.' }
    } catch {
        return [PSCustomObject]@{
            Available = $false
            Detail    = "Not available -- the on-prem cmdlets are loaded but the Exchange server did not answer: $($_.Exception.Message)"
        }
    }
}

function Test-AdAvailable {
    <#
        The ActiveDirectory module plus a domain controller that answers.

        Get-ADRootDSE is the cheapest possible round trip -- it proves the
        module is loaded AND a DC is reachable, which "is the module present"
        alone does not.
    #>
    # No configured directories means this tenant has no on-premises AD, which
    # is a different answer from "the module is missing". Saying so stops every
    # AD tool reporting a fault in a tenant that simply does not have one.
    if (-not (Test-TenantCapability -Name 'OnPremAD')) {
        return [PSCustomObject]@{
            Available = $false
            Detail    = 'Not configured -- Tenant.json does not list on-premises AD for this tenant.'
        }
    }
    $probeServer = Get-TenantDirectoryServer
    if (-not $probeServer) {
        return [PSCustomObject]@{
            Available = $false
            Detail    = 'Not configured -- Tenant.json has OnPremAD enabled but lists no directories.'
        }
    }

    if (-not (Get-Command -Name 'Get-ADUser' -ErrorAction SilentlyContinue)) {
        return [PSCustomObject]@{
            Available = $false
            Detail    = 'Not available -- the ActiveDirectory module is not loaded. Run Import-Module ActiveDirectory.'
        }
    }
    try {
        $null = Get-ADRootDSE -Server $probeServer -ErrorAction Stop
        return [PSCustomObject]@{ Available = $true; Detail = "Active Directory reachable via $probeServer." }
    } catch {
        return [PSCustomObject]@{
            Available = $false
            Detail    = "Not available -- the module is loaded but $probeServer did not answer: $($_.Exception.Message)"
        }
    }
}

#region Gallery modules ------------------------------------------------------
<#
    THE ASSUMPTION THIS REGION EXISTS TO REMOVE: that Exchange Online and
    SharePoint Online management are already installed and imported.

    True on the machines of the team who wrote the toolkit, and false the first
    time it is handed to anyone else. The connect scripts used to print the
    Install-Module line and stop, which is correct behaviour and no help at all
    to an operator who has been told to stop when a tool says no.

    DELIBERATELY ONLY THE TWO PSGALLERY MODULES. ActiveDirectory comes from
    RSAT and Exchange on-premises from the management tools installer; both
    need administrator rights and neither comes from a gallery, so an install
    button for them would be a lie. Those two stay detect-and-report.
#>

function Get-ToolkitModuleDefinition {
    <#
        The gallery modules this tenant needs, derived from Capabilities rather
        than from a hybrid/cloud switch.

        Keyed off Tenant.json for the same reason the directory list is: a
        tenant with SharePoint switched off should not be nagged about a module
        it will never load, and adding a tenant should not mean editing code.
    #>
    param([string]$Id)

    $defs = @()

    if (Test-TenantCapability -Name 'ExchangeOnline') {
        $defs += [PSCustomObject]@{
            Id      = 'ExO'
            Label   = 'Exchange Online'
            Name    = 'ExchangeOnlineManagement'
            # The cmdlet whose PRESENCE proves the module is usable in this
            # process. Cheap to test and, unlike a module-name lookup, it
            # answers the question the caller actually has.
            Probe   = 'Connect-ExchangeOnline'
            Purpose = 'Every Exchange Online tool, and the connection pill in the console header.'
        }
    }

    if (Test-TenantCapability -Name 'SharePointOnline') {
        $defs += [PSCustomObject]@{
            Id      = 'SPO'
            Label   = 'SharePoint Online'
            Name    = 'Microsoft.Online.SharePoint.PowerShell'
            Probe   = 'Connect-SPOService'
            Purpose = 'Orphaned OneDrive, External Sharing, and the OneDrive half of mailbox access.'
        }
    }

    if ($Id) { return @($defs | Where-Object { $_.Id -eq $Id }) | Select-Object -First 1 }
    return @($defs)
}

function Test-ToolkitModule {
    <#
        Is this module usable, merely installed, or absent?

        THE ORDER OF THE TWO CHECKS IS THE WHOLE DESIGN.

        Get-Command is a hashtable lookup against what is already loaded --
        microseconds, and it answers the common case, which is "yes, it is
        right here". Get-Module -ListAvailable walks every path in
        PSModulePath and reads every manifest it finds; on a machine carrying
        several versions of a large module that is seconds, every time.

        So the expensive call runs ONLY when the cheap one has already said no
        -- which is the same moment the caller is about to do something slow
        anyway. That ordering is what removes the need to cache the answer to
        disk: there is no recurring cost to avoid, and therefore nothing to go
        stale, which is the failure mode a cached boolean would have brought
        with it.

        Three states, not two. "Installed but not imported" is a different
        problem from "not installed" and has a different fix -- one is a free
        Import-Module, the other is a download -- and collapsing them into a
        boolean is what makes a tool offer to install something that is already
        sitting on the disk.
    #>
    param([Parameter(Mandatory)][string]$Id)

    $def = Get-ToolkitModuleDefinition -Id $Id
    if (-not $def) {
        return [PSCustomObject]@{
            Id = $Id; Label = $Id; Name = ''; State = 'NotRequired'; Version = ''
            Detail = 'Not required -- this tenant does not have the matching capability enabled.'
        }
    }

    $result = [PSCustomObject]@{
        Id      = $def.Id
        Label   = $def.Label
        Name    = $def.Name
        State   = 'Missing'
        Version = ''
        Detail  = ''
    }

    # 1. Loaded and usable. The cheap path, and the usual answer.
    $cmd = Get-Command -Name $def.Probe -ErrorAction SilentlyContinue
    if ($cmd) {
        $result.State = 'Imported'
        $ver = ''
        if ($cmd.Module -and $cmd.Module.Version) { $ver = "$($cmd.Module.Version)" }
        elseif ($cmd.Version)                      { $ver = "$($cmd.Version)" }
        $result.Version = $ver
        # Built in two statements rather than one interpolated string. The
        # single-line form needed double quotes nested inside a $() inside a
        # double-quoted string -- legal, but unreadable, and it is the kind of
        # line that gets broken by the next person who edits it.
        $result.Detail = 'Loaded and ready.'
        if ($ver) { $result.Detail = "Loaded and ready (version $ver)." }
        return $result
    }

    # 2. On disk but not in this process. Only now do we pay for the scan.
    try {
        $avail = @(Get-Module -ListAvailable -Name $def.Name -ErrorAction SilentlyContinue |
                    Sort-Object Version -Descending)
        if ($avail.Count -gt 0) {
            $result.State   = 'Installed'
            $result.Version = "$($avail[0].Version)"
            $result.Detail  = ("Installed (version $($result.Version)) but not imported into this process. " +
                               "It will be imported when you connect.")
            if ($avail.Count -gt 1) {
                $result.Detail += " $($avail.Count) versions are present; the newest is used."
            }
            return $result
        }
    } catch {
        $result.State  = 'Unknown'
        $result.Detail = "Could not read the module list: $($_.Exception.Message)"
        return $result
    }

    $result.State  = 'Missing'
    $result.Detail = "$($def.Name) is not installed for this user."
    return $result
}

function Import-ToolkitModule {
    <#
        Imports the module so the CONSOLE can see it, not just this script.

        -Global is load-bearing, and Connect-ExSE.ps1 already records why: a
        plain Import-Module inside a script imports into that script's scope.
        The script then reports success, returns, its scope goes away, and the
        very next thing the console does says the module is not available. The
        two statements are both true and read as a contradiction.
    #>
    param([Parameter(Mandatory)][string]$Id)

    $def = Get-ToolkitModuleDefinition -Id $Id
    if (-not $def) { return [PSCustomObject]@{ Ok = $false; Detail = "Unknown module id '$Id'." } }

    try {
        Import-Module -Name $def.Name -Global -DisableNameChecking -ErrorAction Stop
        $after = Test-ToolkitModule -Id $Id
        if ($after.State -eq 'Imported') {
            return [PSCustomObject]@{ Ok = $true; Detail = "Imported $($def.Name) $($after.Version)." }
        }
        return [PSCustomObject]@{
            Ok = $false
            Detail = "Import-Module reported success but $($def.Probe) is still not available."
        }
    } catch {
        return [PSCustomObject]@{ Ok = $false; Detail = "Could not import $($def.Name): $($_.Exception.Message)" }
    }
}

function Install-ToolkitModule {
    <#
        Installs from PSGallery for the CURRENT USER, then imports.

        NEVER CALLED WITHOUT ASKING FIRST. Installing software on someone's
        machine unannounced is a conversation with their desktop team that
        nobody wants to have on the toolkit's behalf. See
        Request-ToolkitModuleInstall.

        THE THREE THINGS THAT MAKE A CLEAN MACHINE FAIL, all of which surface
        as a PROMPT rather than an error -- and a prompt behind a window that
        nobody is looking at is indistinguishable from a hang:

          TLS 1.2      PowerShell 5.1 does not negotiate it by default, and
                       PSGallery requires it. The symptom is "Unable to resolve
                       package source", which reads like DNS and is not.
          NuGet        The provider is bootstrapped on first use, with a
                       confirmation prompt.
          Trust        PSGallery is Untrusted by default, so Install-Module
                       asks whether you are sure.

        All three are handled here rather than left to chance, and -Force plus
        -Confirm:$false means nothing waits on a keystroke.

        The repository is named explicitly. A machine with an internal gallery
        registered as well would otherwise resolve by whatever order the
        repositories happen to be in.
    #>
    param([Parameter(Mandatory)][string]$Id)

    $def = Get-ToolkitModuleDefinition -Id $Id
    if (-not $def) { return [PSCustomObject]@{ Ok = $false; Detail = "Unknown module id '$Id'." } }

    $steps = @()

    try {
        # TLS 1.2, added to whatever is already enabled rather than replacing
        # it -- overwriting the value would drop TLS 1.3 on a machine that has
        # it, to fix a problem that machine does not have.
        $cur = [Net.ServicePointManager]::SecurityProtocol
        if (-not ($cur -band [Net.SecurityProtocolType]::Tls12)) {
            [Net.ServicePointManager]::SecurityProtocol = $cur -bor [Net.SecurityProtocolType]::Tls12
            $steps += 'Enabled TLS 1.2 for this process.'
        }
    } catch {
        $steps += "Could not set TLS 1.2 ($($_.Exception.Message)). The download may fail."
    }

    try {
        $nuget = Get-PackageProvider -Name 'NuGet' -ListAvailable -ErrorAction SilentlyContinue
        if (-not $nuget) {
            $steps += 'Installing the NuGet provider (required before anything can come from PSGallery).'
            Install-PackageProvider -Name 'NuGet' -MinimumVersion '2.8.5.201' `
                -Scope CurrentUser -Force -ErrorAction Stop | Out-Null
        }
    } catch {
        return [PSCustomObject]@{
            Ok = $false
            Detail = ("The NuGet provider could not be installed, so nothing can be downloaded from " +
                      "PSGallery: $($_.Exception.Message)")
            Steps = $steps
        }
    }

    try {
        $steps += "Installing $($def.Name) from PSGallery for the current user only."
        Install-Module -Name $def.Name -Repository 'PSGallery' -Scope CurrentUser `
            -Force -Confirm:$false -ErrorAction Stop
    } catch {
        return [PSCustomObject]@{
            Ok = $false
            Detail = "Install-Module failed for $($def.Name): $($_.Exception.Message)"
            Steps = $steps
        }
    }

    $imp = Import-ToolkitModule -Id $Id
    $steps += $imp.Detail
    return [PSCustomObject]@{ Ok = $imp.Ok; Detail = $imp.Detail; Steps = $steps }
}

function Request-ToolkitModuleInstall {
    <#
        Asks, and returns $true only on an explicit Yes.

        A DIALOG RATHER THAN A TYPED ANSWER, because the operators these
        guardrails were built for read a console window and stop. A modal box
        with two buttons cannot be missed and cannot be answered by accident;
        Enter on a Read-Host can.

        Loads WinForms itself if the caller has not -- the connect scripts are
        console tools and have no reason to carry the assemblies otherwise.
        Test-ConsoleLaunch does the same thing for the same reason.

        Defaults to No. If the box fails to appear at all, the answer is no
        install, which is the safe direction to be wrong in.
    #>
    param([Parameter(Mandatory)][string]$Id)

    $def = Get-ToolkitModuleDefinition -Id $Id
    if (-not $def) { return $false }

    $msg = ("$($def.Label) needs the $($def.Name) module, and it is not installed." +
            [Environment]::NewLine + [Environment]::NewLine +
            "Install it now from the PowerShell Gallery?" +
            [Environment]::NewLine + [Environment]::NewLine +
            "It installs for your account only, needs no administrator rights, and " +
            "changes nothing for anyone else who uses this machine." +
            [Environment]::NewLine + [Environment]::NewLine +
            "Used by: $($def.Purpose)")

    try {
        if (-not ('System.Windows.Forms.MessageBox' -as [type])) {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        }
        $answer = [System.Windows.Forms.MessageBox]::Show(
            $msg, "Install $($def.Name)?",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        return ($answer -eq [System.Windows.Forms.DialogResult]::Yes)
    } catch {
        return $false
    }
}

function Get-ToolkitModuleStateMap {
    <#
        Every required gallery module and its state, cached for the life of the
        window.

        CACHED HERE AND NOWHERE ELSE. Test-ToolkitModule is cheap when the
        answer is "imported" and expensive only when it is not -- fine for a
        connect script, which runs once and is about to do something slower.
        The dashboard is the opposite: it repaints on open and on Refresh, and
        would pay the PSModulePath scan every time for a module that is missing
        precisely because it is missing.

        In memory, not on disk. A file would have to answer "is this still
        true on this machine, for this user, today", and the honest answer is
        that it cannot -- which is the same reason Capabilities is a claim and
        Test-EmsAvailable is a probe.

        The VERSION is carried, not just a boolean. "Installed" stops being
        useful the moment two versions disagree about a parameter, and six
        side-by-side copies of a module is not a hypothetical here.
    #>
    param([switch]$Refresh)

    if ($global:ToolkitModuleStateMap -and -not $Refresh) { return $global:ToolkitModuleStateMap }

    $map = @{}
    foreach ($def in @(Get-ToolkitModuleDefinition)) {
        $map[$def.Id] = Test-ToolkitModule -Id $def.Id
    }
    $global:ToolkitModuleStateMap = $map
    return $map
}
#endregion

function Get-CapabilityStateMap {
    <#
        Probes all four capabilities once and caches the answer for the life of
        the window. Pass -Refresh after connecting to something.

        NOTE the limit this cannot cross: a stale Exchange Online session
        reports as connected and then fails on the next real call with a
        transport error (section 11). So an enabled button means "nothing is
        obviously missing", not "this will work". Greying catches not-connected;
        it cannot catch connected-but-dead.
    #>
    param([switch]$Refresh, [switch]$SkipSpo)

    if ($global:CapabilityStateMap -and -not $Refresh) { return $global:CapabilityStateMap }

    $map = @{}

    $exo = Test-ExoConnectionState
    $map['ExO'] = [PSCustomObject]@{
        Id = 'ExO'; Label = 'Exchange Online'
        Available = [bool]$exo.Connected; Unknown = $false; Detail = "$($exo.Detail)"
    }

    # SharePoint is the expensive probe -- Get-SPOTenant is a network round trip
    # -- so the console skips it on the startup pass. Unknown is NOT the same as
    # unavailable: an unchecked capability must not grey anything out, or every
    # SharePoint tool would be dead until something happened to check.
    if ($SkipSpo) {
        # Carry a real previous answer forward rather than downgrading it to
        # Unknown. Without this, any cheap refresh would un-grey the SharePoint
        # tools it had correctly greyed a moment earlier -- the buttons would
        # flicker between blocked and available depending on which refresh ran
        # last, which is worse than either answer.
        $prior = $null
        if ($global:CapabilityStateMap) { $prior = $global:CapabilityStateMap['SPO'] }
        if ($prior -and -not $prior.Unknown) {
            $map['SPO'] = $prior
        } else {
            $map['SPO'] = [PSCustomObject]@{
                Id = 'SPO'; Label = 'SharePoint Online'
                Available = $true; Unknown = $true
                Detail    = 'Not checked yet. Tools that need SharePoint are left enabled rather than greyed on a guess.'
            }
        }
    } else {
        $spo = Test-SpoConnectionState
        $map['SPO'] = [PSCustomObject]@{
            Id = 'SPO'; Label = 'SharePoint Online'
            Available = [bool]$spo.Connected; Unknown = $false; Detail = "$($spo.Detail)"
        }
    }

    $ems = Test-EmsAvailable
    $map['EMS'] = [PSCustomObject]@{
        Id = 'EMS'; Label = 'Exchange on-premises'
        Available = $ems.Available; Unknown = $false; Detail = $ems.Detail
    }

    $ad = Test-AdAvailable
    $map['AD'] = [PSCustomObject]@{
        Id = 'AD'; Label = 'On-premises Active Directory'
        Available = $ad.Available; Unknown = $false; Detail = $ad.Detail
    }

    # ---- Tenant configuration gate ---------------------------------------
    #
    # The probes above answer "can I reach it right now". Tenant.json answers
    # "does this tenant have it at all". Both are real and they are different
    # questions, so this gate can only ever force a capability OFF -- config
    # cannot conjure a connection, and a probe cannot invent a service the
    # tenant does not own.
    #
    # The distinction is the point. Before this, a cloud-only tenant showed
    # "Exchange on-prem: Not Connected" beside a row of greyed buttons, which
    # reads as a fault to be fixed. It is not a fault; there is nothing there.
    $capMap = @{ ExO = 'ExchangeOnline'; SPO = 'SharePointOnline'
                 EMS = 'OnPremExchange'; AD  = 'OnPremAD' }
    foreach ($id in @($map.Keys)) {
        if (-not $capMap.ContainsKey($id)) { continue }
        $has = $true
        try { $has = Test-TenantCapability -Name $capMap[$id] } catch { $has = $true }
        if (-not $has) {
            $map[$id] = [PSCustomObject]@{
                Id            = $map[$id].Id
                Label         = $map[$id].Label
                Available     = $false
                Unknown       = $false
                NotConfigured = $true
                Detail        = "Not present in this tenant -- Tenant.json does not list $($map[$id].Label)."
            }
        } else {
            Add-Member -InputObject $map[$id] -NotePropertyName 'NotConfigured' `
                       -NotePropertyValue $false -Force
        }
    }

    $global:CapabilityStateMap = $map
    return $map
}

function Get-ToolAvailability {
    <#
        Whether a tool can run right now, and a tooltip explaining any 'no'.

        FAILS OPEN. A script that cannot be read, or whose scan found nothing,
        is reported as available. Every existing Tools.json entry predates this
        feature, so a mechanism that blocked anything it could not classify
        would grey out the whole console the first time it ran -- turning a
        convenience into an outage.
    #>
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [hashtable]$StateMap
    )

    if (-not $StateMap) { $StateMap = Get-CapabilityStateMap }

    $req = Get-ScriptCapabilityRequirement -Path $ScriptPath
    $missing = @()
    $unknown = @()
    foreach ($id in @($req.Requires)) {
        if (-not $StateMap.ContainsKey($id)) { continue }
        if ($StateMap[$id].Unknown)          { $unknown += $StateMap[$id]; continue }
        if (-not $StateMap[$id].Available)   { $missing += $StateMap[$id] }
    }

    if (@($missing).Count -eq 0) {
        $needText = if (@($req.Requires).Count -gt 0) {
            'Needs: ' + ((@($req.Requires) | ForEach-Object { $StateMap[$_].Label }) -join ', ')
        } else {
            'No connection required.'
        }
        foreach ($u in $unknown) { $needText += "`r`n$($u.Label) -- $($u.Detail)" }
        return [PSCustomObject]@{
            Enabled = $true; Requires = @($req.Requires); Missing = @(); Reason = $needText
        }
    }

    # Two different reasons, two different sentences. "Not connected" invites
    # someone to go and connect; "not present in this tenant" tells them not to
    # bother. Collapsing them into one message is how an engineer spends ten
    # minutes trying to reach a server that does not exist.
    $absent = @($missing | Where-Object { $_.PSObject.Properties['NotConfigured'] -and $_.NotConfigured })
    $down   = @($missing | Where-Object { -not ($_.PSObject.Properties['NotConfigured'] -and $_.NotConfigured) })
    $lines = @()
    if ($down.Count -gt 0) {
        $lines += 'Unavailable in this window:'
        foreach ($m in $down) { $lines += "  $($m.Label) -- $($m.Detail)" }
    }
    if ($absent.Count -gt 0) {
        $lines += 'Not available in this tenant:'
        foreach ($m in $absent) { $lines += "  $($m.Label)" }
        $lines += 'This tool can be removed from Tools.json for this deployment.'
    }
    return [PSCustomObject]@{
        Enabled  = $false
        Requires = @($req.Requires)
        Missing  = @($missing | ForEach-Object { $_.Id })
        Reason   = ($lines -join "`r`n")
    }
}
#endregion

#region Browser profiles -----------------------------------------------------
<#
    Different destinations need different identities. ServiceNow wants the
    everyday account; the M365 portals want the admin account. Opening both
    with Start-Process lands them in whichever profile Edge last used, which
    means a sign-in prompt, or worse, the wrong tenant.

    Map purpose -> Edge profile DISPLAY name (the name shown in Edge's profile
    switcher). The display name is deliberately what goes here: Edge's command
    line wants the *directory* name ('Profile 1'), which is an implementation
    detail nobody remembers and which changes if profiles are recreated.
    Get-EdgeProfileDirectory does that translation.
#>
$script:BrowserProfiles = @{
    ServiceNow = 'ServiceNow'   # ticket queue
    Admin      = 'Admin'        # M365 / Entra / Defender portals -- elevated
    GWA        = 'GWA'          # everyday account, for future links
}

function Get-ExchangeServerUri {
    <#
        The remote PowerShell endpoint on the on-premises Exchange server.

        TENANT.JSON ONLY, corrected 2026-08-13.

        Settings.json used to be read FIRST, which made a machine-local
        preference override a tenant fact -- so two engineers at the same
        tenant could be pointed at different Exchange servers with both files
        looking correct, and nothing on screen would disagree. Same correction
        as Get-SpoAdminUrl and for the same reason: an Exchange server URI
        describes the ORGANISATION, not the person at the keyboard.

        The setting is gone rather than demoted. A value that is read but never
        written is a trap for whoever finds it in the file later.

        Accepts either a bare host name or a full URI, because both are things
        a person reasonably types, and rejecting one of them would be a support
        call rather than a feature.

        http, not https, and that is not an oversight. Microsoft: "The
        ConnectionUri value uses http, not https because the session is
        encrypted by the Kerberos token in the payload. Kerberos encryption
        removes the need to encrypt the session again using https."
    #>
    param()

    $configured = ''
    try { $configured = "$(Get-TenantExchangeServerUri)".Trim() } catch { }
    if (-not $configured) { return $null }

    if ($configured -match '^https?://') { return $configured.TrimEnd('/') + '/' }
    return "http://$configured/PowerShell/"
}

function Get-SpoAdminUrl {
    <#
        The SharePoint admin centre URL for this tenant.

        Settings.json first, then derived. Deriving it is safe because the admin
        centre host is not a free choice -- it is always
        https://<tenant>-admin.sharepoint.com -- so the only unknown is the
        tenant label, and there are two places that already know it.

        Session organisation is preferred over the mail domain because the
        SharePoint tenant label follows the *.onmicrosoft.com prefix, not the
        vanity mail domain. They match here; they do not always, and a tenant
        that renamed its mail domain would otherwise get a URL that has never
        existed.

        The setting exists because derivation cannot cover a vanity SharePoint
        host or a tenant whose SPO label genuinely differs. Set it and nothing
        is guessed.

        TENANT.JSON ONLY, CORRECTED 2026-08-13. This read Settings.json and
        then guessed, and never looked at Tenant.json at all -- so
        SharePoint.AdminUrl, which Tenant Manager writes and displays, was
        inert. An engineer could type the correct admin centre into the tool
        whose whole job is tenant facts, save it, and watch Connect SPO use a
        derived host that had never existed.

        The admin centre is a fact about the TENANT, not a preference of the
        engineer, and this file's header has said which file holds which since
        the beginning. Settings.json no longer carries it at all -- the box is
        gone from the Settings dialog and the key is gone from the schema, so
        there is exactly one place to look and one place to change it.
    #>
    param()

    $configured = ''
    try { $configured = Get-TenantSpoAdminUrl } catch { }

    if ($configured) {
        if ($configured -notmatch '^https?://') { $configured = "https://$configured" }
        return $configured.TrimEnd('/')
    }

    $label = ''
    $org = Get-SessionOrganization
    if ($org) { $label = ("$org" -split '\.')[0] }
    if (-not $label) { $label = ((Get-MailDomain) -split '\.')[0] }
    if (-not $label) { return $null }

    return "https://$label-admin.sharepoint.com"
}

# Show-DirectoryChooser was REMOVED on 2026-08-13 along with the ScriptOnPrem
# "AD twin" feature it existed to serve. See ARCHITECTURE section 30.
#
# It asked which directory a paired tool should act on. Nothing calls it now:
# pairs became groups, and the group chooser names the actual tools rather than
# asking about an abstraction the operator could not see.
#
# Kept as a note rather than silently deleted because a function removed from a
# shared file is the kind of thing someone re-adds from memory six months later
# without knowing why it went.

function Get-BrowserProfileMap {
    <#
        Purpose -> Edge profile display name, with Settings.json overriding per
        purpose. Merged rather than replaced, so naming only one purpose in
        Settings.json leaves the others on their defaults instead of deleting
        them -- a partial config should narrow behaviour, not remove it.

        These are per-MACHINE: the profile names are whatever Edge was set up
        with, and a name that does not exist falls back to the last-used profile
        -- which usually works, so the mistake shows up as the wrong account
        rather than as an error.
    #>
    $map = @{}
    foreach ($k in $script:BrowserProfiles.Keys) { $map[$k] = $script:BrowserProfiles[$k] }
    $s = Get-ToolkitSettings
    if ($s.BrowserProfiles) {
        foreach ($p in $s.BrowserProfiles.PSObject.Properties) {
            if ("$($p.Value)".Trim()) { $map[$p.Name] = "$($p.Value)".Trim() }
        }
    }
    return $map
}

function Get-EdgeExecutable {
    foreach ($p in @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles          'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:LOCALAPPDATA          'Microsoft\Edge\Application\msedge.exe')
    )) {
        if ($p -and (Test-Path -LiteralPath $p)) { return $p }
    }
    return $null
}

function Get-EdgeProfiles {
    <#
        Lists Edge profiles as DisplayName -> Directory.

        Edge stores this in its 'Local State' JSON under
        profile.info_cache, keyed by directory name with the display name
        inside. Run this once to see what your profiles are actually called:

            . .\Common.ps1 ; Get-EdgeProfiles
    #>
    $state = Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\User Data\Local State'
    if (-not (Test-Path -LiteralPath $state)) { return @() }
    try {
        $j = Get-Content -LiteralPath $state -Raw | ConvertFrom-Json
        $cache = $j.profile.info_cache
        if (-not $cache) { return @() }
        return @(
            foreach ($dir in $cache.PSObject.Properties.Name) {
                [PSCustomObject]@{
                    DisplayName = "$($cache.$dir.name)"
                    Directory   = $dir
                }
            }
        )
    } catch { return @() }
}

function Get-EdgeProfileDirectory {
    <#
        Resolves an Edge profile display name to its directory name.
        Accepts a directory name directly too, so 'Profile 1' also works.
    #>
    param([string]$Name)
    if (-not $Name) { return $null }
    $all = Get-EdgeProfiles
    if ($all.Count -eq 0) { return $null }

    $hit = $all | Where-Object { $_.DisplayName -ieq $Name } | Select-Object -First 1
    if ($hit) { return $hit.Directory }
    $hit = $all | Where-Object { $_.Directory -ieq $Name } | Select-Object -First 1
    if ($hit) { return $hit.Directory }
    # Partial match last: 'GWA' should find 'GWA (work)'.
    $hit = $all | Where-Object { $_.DisplayName -like "*$Name*" } | Select-Object -First 1
    if ($hit) { return $hit.Directory }
    return $null
}

function Start-UrlInBrowserProfile {
    <#
        Opens a URL in a named Edge profile, falling back cleanly.

        Fallback order, and it matters that this never hard-fails: a browser
        that will not open is a worse outcome than one that opens in the wrong
        profile, because the second is obvious and recoverable.
          1. Edge with the resolved profile directory
          2. Edge with no profile switch
          3. Start-Process (default browser)

        Returns a status string so the caller can say which happened rather
        than leaving the user guessing why they got a sign-in prompt.
    #>
    param(
        [Parameter(Mandatory)][string]$Url,
        [string]$ProfileName,
        [string]$Purpose
    )
    if (-not $ProfileName -and $Purpose) {
        $profileMap = Get-BrowserProfileMap
        if ($profileMap.ContainsKey($Purpose)) {
            $ProfileName = $profileMap[$Purpose]
        } else {
            # A typo'd purpose would otherwise open in the last-used profile
            # and look like it worked -- which is the exact failure this
            # function exists to prevent.
            Write-Warning ("Unknown browser purpose '{0}'. Known: {1}" -f
                           $Purpose, (($profileMap.Keys | Sort-Object) -join ', '))
        }
    }

    $edge = Get-EdgeExecutable
    if ($edge -and $ProfileName) {
        $dir = Get-EdgeProfileDirectory -Name $ProfileName
        if ($dir) {
            try {
                # The quotes around $dir are load-bearing.
                #
                # Edge profile directories are named 'Profile 2', 'Profile 5'
                # and so on -- they contain a SPACE. Start-Process joins
                # -ArgumentList entries with spaces and does not quote them, so
                # an unquoted value arrives at Edge as two arguments:
                #   --profile-directory=Profile   and   2
                # Edge then creates a new profile directory literally called
                # 'Profile', and treats the leftover '2' as a URL -- which it
                # resolves as the IP 0.0.0.2. Every purpose ends up in that same
                # new profile, which is exactly the wrong-identity failure this
                # function exists to prevent.
                Start-Process -FilePath $edge `
                    -ArgumentList @("--profile-directory=`"$dir`"", "`"$Url`"")
                return "Opened in Edge profile '$ProfileName' ($dir)."
            } catch { }
        } else {
            $names = (Get-EdgeProfiles | ForEach-Object { $_.DisplayName }) -join ', '
            try {
                Start-Process -FilePath $edge -ArgumentList @("`"$Url`"")
                return "Edge profile '$ProfileName' not found. Opened in the last-used profile. Available: $names"
            } catch { }
        }
    }
    if ($edge) {
        try {
            Start-Process -FilePath $edge -ArgumentList @("`"$Url`"")
            return 'Opened in Edge (no profile specified).'
        } catch { }
    }
    Start-Process $Url
    return 'Opened in the default browser.'
}
#endregion

function Get-TenantLogoPath {
    <#
        Convention over configuration: a logo is used if one is present, and
        the tenant-specific name wins over the generic one. Nothing is
        hardcoded to a tenant, so the same folder serves any of them.

            Config\Images\logo.<tenant-domain>.png    preferred
            Config\Images\logo.png                    fallback
    #>
    param([string]$Domain)

    $dir = Get-ConfigImageFolder
    if (-not (Test-Path -LiteralPath $dir)) { return $null }
    if (-not $Domain) { $Domain = Get-ConnectedTenantDomain }

    if ($Domain) {
        foreach ($ext in @('png', 'jpg', 'gif', 'bmp')) {
            $p = Join-Path $dir "logo.$Domain.$ext"
            if (Test-Path -LiteralPath $p) { return $p }
        }
    }
    foreach ($ext in @('png', 'jpg', 'gif', 'bmp')) {
        $p = Join-Path $dir "logo.$ext"
        if (Test-Path -LiteralPath $p) { return $p }
    }
    return $null
}

function Add-FormLogo {
    param(
        [Parameter(Mandatory)]$Form,
        [int]$MaxHeight = 40,
        [ValidateSet('TopRight', 'BottomRight')][string]$Corner = 'BottomRight',
        [int]$Margin = 10
    )
    $path = Get-TenantLogoPath
    if (-not $path) { return $null }
    try {
        $img = [System.Drawing.Image]::FromFile($path)

        # MaxHeight is a REQUEST, not a size.
        #
        # Doubling the logo on 2026-08-06 turned out to be impossible to do
        # blindly: at 2x it landed on the tab strip in Disable-User, on a button
        # in LogViewer, and on the input box in AdminConsole. Every form has a
        # different amount of room in that corner, and a constant cannot know
        # which.
        #
        # So it shrinks until it fits. The form gets the biggest logo its own
        # layout allows, up to what was asked for, and nothing is ever covered
        # by a decoration. Making room is then a deliberate layout change rather
        # than something discovered by looking at a screenshot.
        $h = $MaxHeight
        $w = [int]($img.Width * ($h / $img.Height))
        while ($h -gt 16) {
            $w  = [int]($img.Width * ($h / $img.Height))
            $tx = $Form.ClientSize.Width - $w - $Margin
            $ty = if ($Corner -eq 'TopRight') { $Margin }
                  else { $Form.ClientSize.Height - $h - $Margin }
            $rect = New-Object System.Drawing.Rectangle($tx, $ty, $w, $h)

            $clash = $false
            foreach ($c in $Form.Controls) {
                # The help button is allowed to be here -- it gets moved out of
                # the way below rather than being treated as an obstacle.
                if ("$($c.Name)" -eq 'btnToolHelp') { continue }
                if ($rect.IntersectsWith($c.Bounds)) { $clash = $true; break }
            }
            if (-not $clash) { break }
            $h -= 4
        }
        $w = [int]($img.Width * ($h / $img.Height))

        $pb = New-Object System.Windows.Forms.PictureBox
        $pb.Image    = $img
        $pb.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
        $pb.Size     = New-Object System.Drawing.Size($w, $h)
        $pb.BackColor = [System.Drawing.Color]::Transparent
        # Tagged so Add-ToolHelpButton can find it and sit to its left.
        $pb.Tag      = 'NoTheme|ToolkitLogo'
        $pb.Name     = 'pbToolkitLogo'

        $x = $Form.ClientSize.Width - $w - $Margin
        $y = if ($Corner -eq 'TopRight') { $Margin }
             else { $Form.ClientSize.Height - $h - $Margin }
        $pb.Location = New-Object System.Drawing.Point($x, $y)
        $pb.Anchor   = if ($Corner -eq 'TopRight') {
                           [System.Windows.Forms.AnchorStyles]::Top -bor
                           [System.Windows.Forms.AnchorStyles]::Right
                       } else {
                           [System.Windows.Forms.AnchorStyles]::Bottom -bor
                           [System.Windows.Forms.AnchorStyles]::Right
                       }
        $Form.Controls.Add($pb)
        $pb.BringToFront()

        # If the help button was added first, move it clear. Handling both
        # orders means neither function has to be called before the other,
        # which is one less rule to remember and get wrong.
        if ($Corner -eq 'TopRight') {
            foreach ($c in $Form.Controls) {
                if ("$($c.Name)" -eq 'btnToolHelp') {
                    $c.Location = New-Object System.Drawing.Point(($x - $c.Width - 8), $c.Top)
                    $c.BringToFront()
                }
            }
        }
        return $pb
    } catch { return $null }
}
#endregion

#region NEW -- connection state probes (for AdminConsole's Connect buttons) --
function Test-ExoConnectionState {
    param()
    $r = [ordered]@{ Connected = $false; Detail = ''; Organization = $null }

    $probe = @(Get-Command -Name 'Get-Mailbox' -All -ErrorAction SilentlyContinue |
        Where-Object { $_.ModuleName -like 'tmpEXO_*' -or $_.Source -like 'tmpEXO_*' -or
                       $_.ModuleName -eq 'ExchangeOnlineManagement' -or $_.Source -eq 'ExchangeOnlineManagement' })
    if ($probe.Count -eq 0) {
        $r.Detail = 'Not connected -- no Exchange Online cmdlets in this PowerShell process. Connect-ExchangeOnline is per-process, so connecting in another window does not count.'
        return [PSCustomObject]$r
    }

    $gci = Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue
    if (-not $gci) {
        $r.Connected = $true
        $r.Detail = 'Exchange Online cmdlets are loaded. This module version has no Get-ConnectionInformation, so live session state could not be verified.'
        return [PSCustomObject]$r
    }
    try {
        $conns = @(& $gci -ErrorAction Stop)
        if ($conns.Count -eq 0) {
            $r.Detail = 'Not connected -- cmdlets are loaded but there is no active session (it may have expired). Run Connect-ExchangeOnline again.'
            return [PSCustomObject]$r
        }
        $healthy = $conns | Where-Object {
            (-not $_.PSObject.Properties['State']       -or "$($_.State)"       -eq 'Connected') -and
            (-not $_.PSObject.Properties['TokenStatus'] -or "$($_.TokenStatus)" -ne 'Expired')
        }
        $chosen = $conns | Select-Object -First 1
        if ($chosen.PSObject.Properties['Organization']) { $r.Organization = "$($chosen.Organization)" }
        if (@($healthy).Count -gt 0) {
            $r.Connected = $true
            $r.Detail = "Connected to Exchange Online$(if ($r.Organization) { " ($($r.Organization))" })."
        } else {
            $r.Detail = 'Not connected -- the existing session reports itself expired or disconnected. Run Connect-ExchangeOnline again.'
        }
    } catch {
        $r.Connected = $true
        $r.Detail = "Exchange Online cmdlets are loaded; Get-ConnectionInformation could not be read ($($_.Exception.Message))."
    }
    return [PSCustomObject]$r
}

function Get-SessionOrganization {
    <#
        The connected tenant's organisation name, straight from the Exchange
        Online module. Local call, no network -- the module already holds this
        from Connect-ExchangeOnline.

        Returns the *.onmicrosoft.com routing domain, which is nobody's mail
        address but is an unambiguous tenant identifier. Good enough to answer
        "am I where I think I am"; not what you want on screen if a vanity
        domain is available.
    #>
    $gci = Get-Command -Name 'Get-ConnectionInformation' -ErrorAction SilentlyContinue
    if (-not $gci) { return $null }
    try {
        $c = @(& $gci -ErrorAction Stop) | Select-Object -First 1
        if ($c -and $c.PSObject.Properties['Organization']) {
            $o = "$($c.Organization)".Trim()
            if ($o) { return $o }
        }
    } catch { }
    return $null
}

function Test-TenantDomainMismatch {
    <#
        Replaces a check that used to fall out of enumerating every accepted
        domain: is the configured mail domain actually this tenant's?

        The old form confirmed the configured domain appeared in the accepted
        list. This asks the session directly instead -- no network, and it
        catches the case that actually hurts, which is being connected to the
        wrong tenant entirely rather than having a typo in a setting.

        Deliberately loose. A tenant's vanity domain rarely resembles its
        onmicrosoft prefix, so only an obvious disagreement is reported, and
        it is reported as a question rather than an error. Returns $null when
        there is nothing to say.
    #>
    param()

    $org = Get-SessionOrganization
    if (-not $org) { return $null }

    # The parentheses in (Get-ToolkitSettings).MailDomain are load-bearing.
    # Without them PowerShell parses the whole thing as a COMMAND NAME rather
    # than a property access, throws CommandNotFound straight into the catch,
    # and yields nothing. The symptom would be Settings.json being ignored
    # while the code constant still worked -- close to invisible.
    $configured = ''
    try { $configured = "$((Get-ToolkitSettings).MailDomain)".Trim() } catch { }
    if (-not $configured) { $configured = "$((Get-TenantConfig).PrimaryDomain)".Trim() }
    if (-not $configured) { return $null }

    # An onmicrosoft prefix usually contains the organisation's short name, so
    # a shared leading label is treated as agreement.
    $orgLabel = ("$org" -split '\.')[0]
    $cfgLabel = ("$configured" -split '\.')[0]
    if (-not $orgLabel -or -not $cfgLabel) { return $null }
    if ($orgLabel -ieq $cfgLabel) { return $null }
    if ($orgLabel -like "*$cfgLabel*" -or $cfgLabel -like "*$orgLabel*") { return $null }
    if ("$org" -ieq "$configured") { return $null }

    return "Mail domain is set to '$configured' but this session is connected to '$org'. If that is not the tenant you meant, reconnect before making changes."
}

function Get-ConnectedTenantDomain {
    <#
        Which tenant is this console pointed at.

        THIS RUNS AT LAUNCH, and again on every connection-button repaint, so
        its cost is paid over and over. It must not enumerate anything.

        Three sources, cheapest first:

          1. Tenant.json / Settings.json -- already known, zero calls.
          2. Get-ConnectionInformation -- the module's own record of the
             session. LOCAL, no network round trip, and already read once by
             Test-ExoConnectionState during the same startup.
          3. Get-AcceptedDomain -- the network. Only reached when the first two
             are blank.

        The earlier version did (3) unconditionally and then filtered the
        result by the configured domain -- fetching a thousand-plus domains to
        confirm a value it already had, on every repaint. That is where the
        console's startup lag came from.

        Losing the list also loses the implicit check that the configured
        domain really belongs to the connected tenant. Get-ConnectionInformation
        answers that better and for free -- see Test-TenantDomainMismatch.
    #>
    param()
    # Global for the same reason as the theme cache: child script scopes do
    # not inherit $script: state from the console that launched them.
    if ($global:AdminConsoleTenantDomain) { return $global:AdminConsoleTenantDomain }
    if ($script:ConnectedTenantDomainCache) { return $script:ConnectedTenantDomainCache }

    # (1) Configured. Settings.json wins over the code-level constant, matching
    # every other consumer of the mail domain.
    # The parentheses in (Get-ToolkitSettings).MailDomain are load-bearing.
    # Without them PowerShell parses the whole thing as a COMMAND NAME rather
    # than a property access, throws CommandNotFound straight into the catch,
    # and yields nothing. The symptom would be Settings.json being ignored
    # while the code constant still worked -- close to invisible.
    $configured = ''
    try { $configured = "$((Get-ToolkitSettings).MailDomain)".Trim() } catch { }
    if (-not $configured) { $configured = "$((Get-TenantConfig).PrimaryDomain)".Trim() }
    if ($configured) {
        $script:ConnectedTenantDomainCache = $configured
        $global:AdminConsoleTenantDomain   = $configured
        return $configured
    }

    # (2) The session's own record. No network.
    $org = Get-SessionOrganization
    if ($org) {
        $script:ConnectedTenantDomainCache = $org
        $global:AdminConsoleTenantDomain   = $org
        return $org
    }

    # (3) Network, last resort.
    try {
        # ResultSize: see Get-AcceptedDomainList. Without it the cmdlet stops
        # at 1000 domains, and the console's whole identity -- tenant label,
        # theme binding, directory detection -- is decided from whatever
        # arbitrary slice came back first.
        $accepted = @(Invoke-EXOCommand -Name 'Get-AcceptedDomain' -Parameters @{
            ResultSize = 'Unlimited'; ErrorAction = 'Stop' })
        if ($accepted.Count -eq 0) { return $null }

        $names = @($accepted | ForEach-Object { "$($_.DomainName)" })

        # Order matters, and 'first non-onmicrosoft' is deliberately NOT first.
        # In a tenant with hundreds of accepted domains that is whichever one
        # Exchange happened to return first -- a branch office, not the
        # organisation. The tenant's own Default answers the question properly;
        # it is only skipped when it is the .onmicrosoft.com fallback, which
        # identifies the tenant correctly but reads as noise on screen.
        # No configured-domain test here -- this branch is only reached when
        # the configured domain was blank, so that comparison could only ever
        # match nothing.
        $pick = $null
        if (-not $pick) {
            $def = $accepted | Where-Object { $_.Default -eq $true } | Select-Object -First 1
            if ($def -and "$($def.DomainName)" -notlike '*.onmicrosoft.com') {
                $pick = "$($def.DomainName)"
            }
        }
        if (-not $pick) {
            $pick = $names | Where-Object { $_ -notlike '*.onmicrosoft.com' } | Select-Object -First 1
        }
        if (-not $pick) {
            $def = $accepted | Where-Object { $_.Default -eq $true } | Select-Object -First 1
            if ($def) { $pick = "$($def.DomainName)" }
        }
        if (-not $pick) { $pick = $names[0] }

        if ($pick) {
            $script:ConnectedTenantDomainCache = $pick
            $global:AdminConsoleTenantDomain  = $pick
            return $pick
        }
    } catch {
    }
    return $null
}

function Clear-ConnectedTenantDomainCache {
    $script:ConnectedTenantDomainCache = $null
    $global:AdminConsoleTenantDomain  = $null
}

function Test-SpoConnectionState {
    param()
    $r = [ordered]@{ Connected = $false; Detail = '' }
    if (-not (Get-Command -Name 'Get-SPOTenant' -ErrorAction SilentlyContinue)) {
        $r.Detail = 'Not connected -- SharePoint Online cmdlets are not loaded in this process. Import Microsoft.Online.SharePoint.PowerShell and run Connect-SPOService.'
        return [PSCustomObject]$r
    }
    try {
        $null = Get-SPOTenant -ErrorAction Stop
        $r.Connected = $true
        $r.Detail = 'Connected to SharePoint Online.'
    } catch {
        $r.Detail = "Not connected -- SPO cmdlets are loaded but the session is missing or expired: $($_.Exception.Message)"
    }
    return [PSCustomObject]$r
}
#endregion

#region NEW -- shared "enter a target" prompt for standalone tool launches --
function Show-TargetPromptDialog {
    param(
        [string]$InitialSmtp,
        [string]$Title = 'Enter Target'
    )
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title
    $dlg.ClientSize = New-Object System.Drawing.Size(464, 140)
    $dlg.StartPosition = 'CenterScreen'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Text = 'Mailbox or Group (email address):'
    $lblUser.Location = New-Object System.Drawing.Point(12, 12)
    $lblUser.AutoSize = $true
    $dlg.Controls.Add($lblUser)
    $txtUser = New-Object System.Windows.Forms.TextBox
    $txtUser.Location = New-Object System.Drawing.Point(12, 30)
    $txtUser.Size = New-Object System.Drawing.Size(440, 22)
    $txtUser.Font = New-Object System.Drawing.Font('Consolas', 9)
    $txtUser.Text = $InitialSmtp
    $dlg.Controls.Add($txtUser)

    # The hint that used to sit here has been removed, not reworded.
    #
    # It said the directory is worked out automatically. That is true, and it
    # is also not the operator's problem: they typed an address, and how the
    # toolkit finds it is an implementation detail. On a cloud-only tenant it
    # was worse than noise -- it described machinery that does not exist there.
    #
    # A dialog with one input box does not need a sentence explaining itself.

    $lblErr = New-Object System.Windows.Forms.Label
    $lblErr.Location = New-Object System.Drawing.Point(12, 78)
    $lblErr.Size = New-Object System.Drawing.Size(440, 20)
    $lblErr.ForeColor = [System.Drawing.Color]::Red
    $dlg.Controls.Add($lblErr)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Location = New-Object System.Drawing.Point(280, 100)
    $btnCancel.Size = New-Object System.Drawing.Size(84, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton = $btnCancel
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'OK'
    $btnOk.Location = New-Object System.Drawing.Point(368, 100)
    $btnOk.Size = New-Object System.Drawing.Size(84, 28)
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton = $btnOk

    $btnOk.Add_Click({
        $smtp = Get-TargetSmtp -Text $txtUser.Text
        if (-not $smtp) { $lblErr.Text = 'Enter a mailbox or group email address.'; return }
        $dlg.Tag = $smtp
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $dlg.Close()
    })

    $result = $dlg.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    return $dlg.Tag
}

function Resolve-ToolTarget {
    param(
        [string]$SmtpAddress,
        [string]$Environment,
        [string]$DomainController,
        [string]$ToolTitle = 'Enter Target',

        # A CLOUD-ONLY OBJECT IN A HYBRID TENANT IS NOT AN ERROR.
        #
        # The block below already forgives a cloud-only TENANT. It does not
        # forgive a cloud-only OBJECT in a tenant that also has AD -- and those
        # are now routine here: a distribution group created in Exchange Online
        # has no on-premises object at all, so the lookup misses and the address
        # is rejected as a typo. 'AP-3255@...' is a real, working group and the
        # console refused to open it.
        #
        # This will get worse rather than better. Every new group is created in
        # the cloud from now on, and every one of them fails this test.
        #
        # OPT-IN rather than the default, because the on-premises tools NEED the
        # refusal: AD-GroupEditor with a null Environment has no directory to
        # write to, and failing at the front door is far better than failing
        # halfway through. Cloud tools pass this switch; on-premises tools must
        # not.
        [switch]$CloudOk
    )
    $smtp = Get-TargetSmtp -Text $SmtpAddress

    # Valid directory Ids come from Tenant.json, not from a pair of literals.
    # This read as a membership test against two literal directory names, so
    # a tenant that
    # named its directories anything else failed this test every time and fell
    # through to a lookup it did not need.
    $validIds = @()
    try { $validIds = @(Get-TenantDirectory | ForEach-Object { "$($_.Id)" }) } catch { }
    $envOk = $Environment -and ($validIds -contains $Environment)

    if ($smtp -and $envOk) {
        return [PSCustomObject]@{
            Smtp        = $smtp
            Environment = $Environment
            ADServer    = Get-ADServerString -Environment $Environment -DomainController $DomainController
        }
    }
    if (-not $smtp) {
        $smtp = Show-TargetPromptDialog -InitialSmtp $SmtpAddress -Title $ToolTitle
        if (-not $smtp) { return $null }
    }

    # A CLOUD-ONLY TENANT HAS NOTHING TO LOOK UP, AND THAT IS NOT A FAILURE.
    #
    # This used to demand an on-premises hit before returning, so on a tenant
    # with no AD every address was rejected as "not found" and tools like Out
    # of Office could never resolve anything at all. The address was fine; the
    # question was wrong.
    #
    # Cloud tools need the ADDRESS. Environment and ADServer exist for the
    # on-premises ones, and null is the honest answer when there is no
    # on-premises directory to name.
    $hasAd = $true
    try { $hasAd = Test-TenantCapability -Name 'OnPremAD' } catch { $hasAd = $true }
    if (-not $hasAd) {
        return [PSCustomObject]@{ Smtp = $smtp; Environment = $null; ADServer = $null }
    }

    $envResult = Find-TargetEnvironment -Smtp $smtp
    if (-not $envResult.Environment) {
        # See -CloudOk above. The caller has told us it can work from the
        # address alone, so a missing on-premises object means "cloud-only",
        # not "wrong address".
        if ($CloudOk) {
            return [PSCustomObject]@{ Smtp = $smtp; Environment = $null; ADServer = $null }
        }
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        [void][System.Windows.Forms.MessageBox]::Show(
            "'$smtp' was not found in the on-premises directory.`n`nDouble-check the address and try again.",
            $ToolTitle, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $null
    }
    return [PSCustomObject]@{
        Smtp        = $smtp
        Environment = $envResult.Environment
        ADServer    = $envResult.ADServer
    }
}

function Resolve-ToolTargetOnDemand {
    <#
        Resolve-ToolTarget for a tool whose address box is ON ITS OWN FORM.

        THE ON-PREMISES TOOLS CANNOT SIMPLY DROP Resolve-ToolTarget the way the
        cloud ones did. They need -ADServer: which directory holds this account
        and which controller to talk to. Without it there is nothing to write
        to, and failing at the front door beats failing halfway through a
        multi-step edit. The refusal is a feature, not friction.

        What DOES have to change is when it runs. It used to run once, before
        the window existed. Now the address can change at any time, so it runs
        again on every lookup -- and the one thing it must not do in that
        position is open the modal prompt it opens for an empty address. The
        box on the form IS that prompt.

        So: empty in, $null out, no dialog. Anything else is handed to
        Resolve-ToolTarget unchanged, warning box and all.

        -Environment is deliberately NOT forwarded. It is an answer, not a
        question: passing the one resolved for the PREVIOUS address would make
        Resolve-ToolTarget trust it and skip the lookup, quietly pointing the
        new address at the old directory.
    #>
    param(
        [string]$Text,
        [string]$ToolTitle = 'Enter Target',
        [string]$DomainController
    )
    $smtp = Get-TargetSmtp -Text $Text
    if (-not $smtp) { return $null }
    return Resolve-ToolTarget -SmtpAddress $smtp -DomainController $DomainController -ToolTitle $ToolTitle
}
#endregion
```


<br>

==============================================================================

# ◆ FILE 2 OF 16 — `Test-Toolkit.ps1`

> **1950 lines.** THE VERIFIER, already built. Run it first, extend it -- do NOT reimplement it from the spec. It lacks PS7002, GRF001 and FLT001 because it was written on the target platform, where none of those classes of defect could occur.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```powershell
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
        # only, so a licence group name sat hardcoded in two scripts that
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
        '\b[a-z0-9-]+\.(?:corp|internal|ad|local|intranet|lan)\.[a-z.]+\b'
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

                # WORD BOUNDARY, not substring. A short TenantName matched a
                # longer ordinary word -- and @{ Expression = ... } appears in
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

        if ($bare -notmatch '\btry\b' -and $bare -match '\b(?:Get|Set|New|Remove|Add|Enable|Disable)-(?:Mailbox|Recipient|DistributionGroup|UnifiedGroup|RemoteMailbox|InboxRule|MailboxPermission|RecipientPermission|ADUser|ADGroup)\w*\b') {
            Add-Finding -Severity WARN -File $File -Line $line -Rule 'HDL001' -Message ('Click handler calls Exchange/AD directly with no try/catch. An exception here is caught by .NET, not PowerShell, and the operator sees assembly versions instead of a line number. JUDGEMENT: a handler that delegates to a function which catches its own errors is fine -- check where the call goes before adding a try here.')
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
```


<br>

==============================================================================

# ◆ FILE 3 OF 16 — `AdminConsole.ps1`

> **2232 lines.** The launcher: connection state in the header, the PIM countdown, capability-driven greying, JSON-driven section/button layout, and the modal-loop deferral (constraint B2).
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```powershell
#Requires -Version 5.1
<#
    AdminConsole.ps1
    Layer 1 launcher -- Mailbox / Group lookup + dynamic tool menu.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-RUNS: whatever you connect it to. The status row at the top of this window shows what it currently has.
# TOOLKIT-HELP-BEGIN
# The launcher. Look up a mailbox or group, then use the tools below.
#
# ---------------------------------------------------------------------------
# WHAT THIS TENANT HAS
# ---------------------------------------------------------------------------
#
# The status row shows only what this tenant actually has. A tenant with no
# on-premises Exchange has no EMS indicator at all -- not a grey one. Config\
# Tenant.json decides that, and TenantManager edits it.
#
# (This used to say ToolManager > Settings, which cannot edit Tenant.json and
# never could. That dialog writes Settings.json -- engineer and machine
# preferences. Tenant.json is tenant FACTS and had no editor at all until
# TenantManager was written.)
#
# So the row describes THIS deployment. If something you expect is missing from
# it, the capability is switched off in Tenant.json rather than broken.
#
# ---------------------------------------------------------------------------
# CONNECT FIRST
# ---------------------------------------------------------------------------
#
# Tools stay greyed until this window holds the connection they need. The
# connection belongs to THIS window, not to your account.
#
# If your admin rights are activated on demand rather than standing, activate
# BEFORE connecting. A session made first holds the token it was issued with,
# and that token does not contain access granted afterwards -- everything then
# fails as though the activation never happened. Activate, then reconnect.
#
# The PIM row and the Activate button appear only where Tenant.json says this
# tenant uses on-demand rights.
#
# ---------------------------------------------------------------------------
# CHECK THE TENANT BEFORE YOU WRITE
# ---------------------------------------------------------------------------
#
# One toolkit serves several tenants and looks the same in all of them. The
# tenant name is in the top row, and a colour can be bound per tenant in
# ThemeManager -- colour tells you where you are faster than reading does.
#
# A (?) beside the tenant name means the connected domain does not match what
# Tenant.json expects. That is a question, not a fault -- a tenant with several
# accepted domains can disagree legitimately -- but answer it before a write.
#
# ---------------------------------------------------------------------------
# WHY A BUTTON IS GREYED
# ---------------------------------------------------------------------------
#
# Three reasons, and they want three different responses:
#
#   * WRONG OBJECT TYPE. The tool cannot act on what you just looked up. Look
#     up something else and it comes back.
#
#   * NOT CONNECTED. It needs a service you have not connected in this window.
#     Connect it -- and if rights are activated on demand, activate first.
#
#   * NOT AVAILABLE IN THIS TENANT. The tool needs something this tenant does
#     not have. Connecting will not help. That tool can be removed from
#     Tools.json for this deployment, and the tooltip says so.
#
# Hover a greyed button and the tooltip says which of the three it is.
#
# ---------------------------------------------------------------------------
#
# The line at the bottom shows the toolkit version and whether the files still
# match it. Amber means this copy differs from the one it was stamped from --
# run Tester for the list.
#
# Every tool has its own ? button explaining what it does and what it touches.
# TOOLKIT-HELP-END

param(
    [string]$ToolsConfigPath = ''
)

#region Relaunch trampoline ----------------------------------------------------
<#
    THE CONSOLE RE-OPENS ITSELF IN A LOOP, NOT BY RECURSION.

    Why it re-opens at all: a Connect script has to run in THIS process, or the
    session it establishes belongs to a process that is about to exit and this
    console stays as disconnected as it was. So the window closes, the connect
    runs on the same thread, and the console comes back.

    HOW IT USED TO COME BACK, AND WHY THAT WAS WRONG. The bottom of this file
    called '& $PSCommandPath' -- the script invoking itself. That works, but
    every connect left the process one stack frame deeper, and the frames only
    unwound when the console was finally closed. A depth counter capped it at 8
    and printed "not reopening again", which is a guard doing its job and an
    operator being told to go away.

    Nobody hit it while PowerShell was restarted daily; the counter reset with
    the process. It surfaced the day the console first stayed open overnight --
    ExchangeOnlineManagement v3 keeps a REST-backed session alive, so the
    process now outlives the working day and the connects accumulate.

    The fix is a trampoline. The FIRST invocation becomes a loop and does no
    work of its own; it calls the script again and keeps calling it for as long
    as the body asks to be re-opened. Every real invocation therefore runs at
    exactly one level deep, forever, and the depth guard is gone because there
    is no depth to guard.

    $global: rather than $script:, because the two invocations are separate
    script scopes in one process and a $script: variable would not be visible
    across them. Cleared in a finally so an unhandled error inside the console
    cannot leave the flag set and strand the next launch inside a loop.
#>
if (-not $global:ToolkitConsoleTrampoline) {
    $global:ToolkitConsoleTrampoline = $true
    try {
        do {
            # Reset BEFORE the call, not after. The body sets this when it
            # wants to come back; reading a stale value would re-open a console
            # nobody asked for.
            $global:ToolkitConsoleRelaunch = $false
            & $PSCommandPath @PSBoundParameters
        } while ($global:ToolkitConsoleRelaunch)
    } finally {
        $global:ToolkitConsoleTrampoline = $null
        $global:ToolkitConsoleRelaunch   = $null
    }
    return
}
#endregion

#region Startup breadcrumbs ----------------------------------------------------
<#
    A trail on disk, written before anything else can fail.

    WHY THIS EXISTS. The console occasionally appears and vanishes, and the
    second launch is always fine (first seen 2026-08-13). Nothing was recorded,
    so there was nothing to work from.

    The first attempt at this wrote the reason into the console's own output
    box, which was useless: if the window dies, the message dies with it. That
    is the whole failure mode. A diagnostic has to outlive the thing it is
    diagnosing, so this writes to a FILE.

    IT LOGS PROGRESS, NOT JUST ERRORS. An exception that never surfaces still
    leaves the last breadcrumb behind, so the file answers WHERE it stopped even
    when it cannot say why. "Reached: dot-sourcing Common" followed by nothing
    is a complete answer.

    DELIBERATELY SELF-CONTAINED. No Common.ps1, no Get-LogPath, no toolkit
    helpers -- those are the very things that might not have loaded. Only
    $PSScriptRoot and built-ins, and every write is wrapped so that a failure to
    log can never itself take the console down.

    The file is truncated at 200 lines on each launch, so it stays readable and
    cannot grow without bound.
#>
$script:StartupLogPath = Join-Path $PSScriptRoot 'Logs\ConsoleStartup.log'

function Write-StartupBreadcrumb {
    param([string]$Text)
    try {
        $dir = Split-Path -Parent $script:StartupLogPath
        if (-not (Test-Path -LiteralPath $dir)) {
            New-Item -Path $dir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        $line = '{0}  pid {1}  {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $PID, $Text
        Add-Content -LiteralPath $script:StartupLogPath -Value $line -Encoding UTF8 -ErrorAction Stop
    } catch { }
}

function Write-StartupFailure {
    param([string]$Stage, $ErrorRecord)
    try {
        Write-StartupBreadcrumb -Text "FAILED at: $Stage"
        if ($ErrorRecord) {
            Write-StartupBreadcrumb -Text "    $($ErrorRecord.Exception.GetType().FullName): $($ErrorRecord.Exception.Message)"
            # The stack is what turns "it broke" into a line number, and it is
            # the first thing lost when an exception is caught and summarised.
            $st = "$($ErrorRecord.ScriptStackTrace)"
            foreach ($l in ($st -split "`r?`n")) { if ("$l".Trim()) { Write-StartupBreadcrumb -Text "    $l" } }
        }
    } catch { }
}

# Keep the file short. Reading a startup log means reading the LAST run, and a
# thousand old lines is how the relevant twelve get missed.
try {
    if (Test-Path -LiteralPath $script:StartupLogPath) {
        $existing = @(Get-Content -LiteralPath $script:StartupLogPath -ErrorAction Stop)
        if ($existing.Count -gt 200) {
            Set-Content -LiteralPath $script:StartupLogPath -Value $existing[-200..-1] -Encoding UTF8 -ErrorAction Stop
        }
    }
} catch { }

Write-StartupBreadcrumb -Text '--- launch ---'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Write-StartupBreadcrumb -Text 'WinForms assemblies loaded'

try {
    . (Join-Path $PSScriptRoot 'Common.ps1')
    Write-StartupBreadcrumb -Text 'Common.ps1 dot-sourced'
} catch {
    Write-StartupFailure -Stage 'dot-sourcing Common.ps1' -ErrorRecord $_
    throw
}

# THE MARKER THAT SAYS "the console started this".
#
# Set once, here, before any tool can be launched -- and before the form is
# built, so there is no path where a button exists and the marker does not.
#
# A GUID rather than a fixed word, so it cannot be guessed from reading one
# tool. That is not a security property -- anyone can read Common.ps1 and see
# that any value passes -- it just means the value in a process listing is not
# a reusable password.
#
# Environment variables are inherited by child processes, which is why this
# covers both launch paths: the call operator (same process) and Start-Process
# (new process, inherits). Nothing needs passing through the command line.
#
# It dies with this process. A prompt opened afterwards does not have it.
$env:ToolkitConsoleSession = [guid]::NewGuid().ToString()
Write-StartupBreadcrumb -Text 'console session marker set'
#endregion

# Config now lives in Config\. Resolved after Common.ps1 loads, because the
# resolver is defined there -- a param default cannot call it.
if (-not $ToolsConfigPath) { $ToolsConfigPath = Get-ConfigFilePath -Name 'Tools.json' }

# One-time tidy-up for installs that predate Config\. Runs before anything reads
# a config file, and reports what it did rather than moving files silently.
$script:ConfigMigrationNotes = @(Invoke-ConfigMigration)

# No tenant configured -- say so, loudly, before drawing anything.
#
# Test-TenantCapability returns $false when Tenant.json is absent, and the
# capability gate HIDES tools that are off. So without this the console opens
# looking finished, with most of its buttons simply not there, and nothing
# anywhere saying why. That is indistinguishable from a broken install, and it
# is the first thing a new tenant would ever see.
#
# The toolkit deliberately ships WITHOUT a Tenant.json -- only
# Config\Tenant.TEMPLATE.json -- because shipping one populated with somebody
# else's tenant is worse: it looks configured and is wrong.
if (-not (Test-Path -LiteralPath (Get-ConfigFilePath -Name 'Tenant.json'))) {
    $tm = Join-Path $PSScriptRoot 'TenantManager.ps1'
    $msg = "No Config\Tenant.json, so this deployment does not know which tenant it is." +
           [Environment]::NewLine + [Environment]::NewLine +
           "Until it exists, every capability reads as OFF and the tools that need one are HIDDEN -- the console will look half empty." +
           [Environment]::NewLine + [Environment]::NewLine
    if (Test-Path -LiteralPath $tm) {
        $msg += "Open Tenant Manager now to set it up?"
        $ans = [System.Windows.Forms.MessageBox]::Show($msg, 'Tenant not configured',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Its own process: this console has not drawn its form yet, and a
            # modal dialog on this thread would block the setup window it just
            # offered to open. See 14d.
            Start-Process -FilePath (Get-Command powershell.exe).Source `
                -ArgumentList "-ExecutionPolicy Bypass -File `"$tm`""
            return
        }
    } else {
        $msg += "Run TenantManager.ps1 to set it up."
        [void][System.Windows.Forms.MessageBox]::Show($msg, 'Tenant not configured',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
    }
}

#region Tools.json load -------------------------------------------------------
function Get-ToolDefinitions {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Tools.json not found at '$Path'. Use ToolManager.ps1 to create one, or copy the seed Tools.json shipped with this console."
    }
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        $tools = @($raw | ConvertFrom-Json -ErrorAction Stop)
    } catch {
        throw "Tools.json at '$Path' could not be parsed as JSON: $($_.Exception.Message)"
    }
    return @($tools | Sort-Object { if ($_.PSObject.Properties['Order']) { [int]$_.Order } else { 999 } })
}
#endregion

#region Main window ------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Admin Console"
$form.Size = New-Object System.Drawing.Size(840, 808)
$form.StartPosition = "Manual"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

$lblExo = New-Object System.Windows.Forms.Label
$lblExo.Tag = 'NoTheme'
$lblExo.Location = New-Object System.Drawing.Point(12, 8)
$lblExo.Size = New-Object System.Drawing.Size(156, 20)
$lblExo.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$lblExo.Text = "ExO: ..."
$form.Controls.Add($lblExo)

$lblSpo = New-Object System.Windows.Forms.Label
$lblSpo.Tag = 'NoTheme'
$lblSpo.Location = New-Object System.Drawing.Point(170, 8)
$lblSpo.Size = New-Object System.Drawing.Size(156, 20)
$lblSpo.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$lblSpo.Text = "SPO: ..."
$form.Controls.Add($lblSpo)

$lblTenant = New-Object System.Windows.Forms.Label
$lblTenant.Tag = 'NoTheme'
$lblTenant.Location = New-Object System.Drawing.Point(330, 8)
$lblTenant.Size = New-Object System.Drawing.Size(330, 20)
$lblTenant.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$lblTenant.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblTenant)

# On-premises indicators share the second row with the hint text. The
# capability map has always probed EMS and AD -- that is what greys the on-prem
# tools out -- but nothing displayed the answer, so a missing on-prem session
# showed up as a mysteriously dead button rather than as a status you could read.
$lblEms = New-Object System.Windows.Forms.Label
$lblEms.Tag = 'NoTheme'
$lblEms.Location = New-Object System.Drawing.Point(12, 30)
$lblEms.Size = New-Object System.Drawing.Size(156, 20)
$lblEms.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$lblEms.Text = "EMS: ..."
$form.Controls.Add($lblEms)

$lblAd = New-Object System.Windows.Forms.Label
$lblAd.Tag = 'NoTheme'
$lblAd.Location = New-Object System.Drawing.Point(170, 30)
$lblAd.Size = New-Object System.Drawing.Size(156, 20)
$lblAd.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 10, [System.Drawing.FontStyle]::Bold)
$lblAd.Text = "AD: ..."
$form.Controls.Add($lblAd)

$lblWarn = New-Object System.Windows.Forms.Label
$lblWarn.Tag = 'NoTheme'
# Short, and stopping at x=640, because everything from there to the right edge
# is reserved for the logo and the '?'.
#
# It was 482px wide reaching x=812, which is why the logo had nowhere to grow
# and the '?' kept being pushed into it. The four status labels above already
# name exactly what is not connected, so this only has to point at the buttons.
$lblWarn.Text = "Use Connect below."
$lblWarn.Location = New-Object System.Drawing.Point(508, 31)
$lblWarn.Size = New-Object System.Drawing.Size(132, 18)
$lblWarn.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Italic)
$lblWarn.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblWarn)

# PIM countdown. Row 2, DIRECTLY BELOW the tenant label -- both start at x=330.
#
# It was in the top-right corner, where the '?' sat on top of it. The help
# button positions itself at run time and treats a Label as an obstacle only as
# far as its TEXT reaches -- and at the moment the button is placed this label
# is still empty and hidden, so it measured as zero-width and was not avoided.
# Content that appears later cannot be dodged by a layout decision made
# earlier; the fix is to put it somewhere the corner never reaches.
#
# Hidden until something has actually been recorded, so a console that has
# never used the Activate button shows nothing rather than a blank promise.
$lblPim           = New-Object System.Windows.Forms.Label
$lblPim.Tag       = 'NoTheme'
$lblPim.Location  = New-Object System.Drawing.Point(330, 31)

# AUTOSIZE, not a fixed 170px.
#
# The line holds Groups AND Roles once both timers are running, and 170px fits
# only the first. Roles was not hidden -- it was off the right-hand edge of a
# label too small for its own text, which looks identical to "not set".
#
# Update-StatusRowLayout positions the labels after it by adding $lbl.Width, so
# a fixed width that disagrees with the content also mis-places everything to
# its right. AutoSize keeps Width honest and the row flows correctly.
$lblPim.AutoSize  = $true
$lblPim.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$lblPim.Visible   = $false
$form.Controls.Add($lblPim)

function Update-PimLabel {
    <#
        TAKES THE CONTROLS AS PARAMETERS. It used to read $lblPim and $toolTip
        from the enclosing scope, and that was a real fault rather than a
        stylistic one.

        This function is called from a 60-second Timer tick. A Timer tick
        scriptblock is not a closure, so PowerShell resolves its variables
        against whatever scope is CURRENT when the tick fires -- and the
        console launches tools with '& $scriptPath', which gives the tool a
        CHILD scope of this one. While a tool's ShowDialog() pumps messages on
        this same thread, the timer keeps ticking, and any name the tool has
        assigned to shadows the console's.

        Dashboard.ps1 opens with

            $lblPim = $null
            if ($script:HasPim) { $lblPim = Add-DashLine ... }

        so in a tenant WITHOUT PIM the timer found the dashboard's $null and
        threw "The property 'Visible' cannot be found on this object" -- from
        inside this function's own catch, which then threw as well, so nothing
        contained it and the operator got a JIT-debugging dialog captioned
        'Admin Dashboard'.

        IN A TENANT WITH PIM IT DID NOT THROW, WHICH IS WORSE. The dashboard's
        $lblPim was a real Label, so the console's countdown was written into
        the DASHBOARD's label instead of its own -- silently, for as long as
        both existed.

        Passing the controls in removes the name lookup entirely. The timer
        carries them on its own .Tag; see the wiring at the bottom of the file.
    #>
    param(
        [System.Windows.Forms.Label]$Lbl,
        $Tip
    )
    if (-not $Lbl) { return }
    try {
        $line = Get-PimStatusLine
        if (-not $line) { $Lbl.Visible = $false; return }
        $Lbl.Text      = $line
        $Lbl.ForeColor = Get-PimStatusColour
        $Lbl.Visible   = $true
        if ($Tip) {
            $Tip.SetToolTip($Lbl, 'Recorded by the Activate button, not read from Entra. Clear it there if it looks wrong.')
        }
    } catch { $Lbl.Visible = $false }
}

# Lay the status row out from whatever is actually showing.
#
# A capability this tenant does not have is HIDDEN, not greyed. Grey still
# occupies space and still invites a glance -- and "EMS: n/a" in a cloud-only
# tenant is a label that will never say anything else. The row should describe
# this tenant, not the union of every tenant.
#
# Hiding leaves a hole, so the survivors are repositioned rather than left with
# a gap where a control used to be. Fixed coordinates were the reason the old
# row could not lose a member.
#
# Tool BUTTONS keep their grey. That is different: a greyed button says "this
# tool is here and cannot run", which is exactly the prompt to take it out of
# Tools.json. A status label has no such job.
function Update-StatusRowLayout {
    param([System.Windows.Forms.Label[]]$Row, [int]$Y, [int]$Gap = 8)
    $x = 12
    foreach ($lbl in $Row) {
        if (-not $lbl -or -not $lbl.Visible) { continue }
        $lbl.Location = New-Object System.Drawing.Point($x, $Y)
        $x += $lbl.Width + $Gap
    }
}

function Update-TenantLabel {
    <#
        The label now comes from configuration or from the session's own
        record, never from enumerating accepted domains -- see
        Get-ConnectedTenantDomain. That makes it free to call, which matters
        because Update-ConnectionButtonColors clears the cache and calls it
        again on every repaint.

        Because it can now report a CONFIGURED value rather than a verified
        one, the mismatch check is what keeps it honest: connected to the wrong
        tenant, the label alone would look perfectly correct.
    #>
    $d = Get-ConnectedTenantDomain
    if ($d) {
        $lblTenant.Text = "Tenant: $d"
        $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(80, 220, 130)
    } else {
        $lblTenant.Text = "Tenant: not connected"
        $lblTenant.ForeColor = [System.Drawing.Color]::DimGray
    }

    # Amber, not red, and it does not block anything. A tenant with several
    # accepted domains can disagree with its onmicrosoft prefix legitimately,
    # so this is a question worth asking rather than a fault worth stopping on.
    $mismatch = Test-TenantDomainMismatch
    if ($mismatch) {
        $lblTenant.Text = "$($lblTenant.Text)  (?)"
        $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(176, 96, 0)
        $toolTip.SetToolTip($lblTenant, $mismatch)
    } else {
        $toolTip.SetToolTip($lblTenant, '')
    }
}

$script:DefaultDC = (Get-TenantDirectoryServer)

$lblInput = New-Object System.Windows.Forms.Label
$lblInput.Text = "Enter mailbox or group email address ($($(Get-MailDomain)) optional):"
$lblInput.Location = New-Object System.Drawing.Point(12, 50)
$lblInput.AutoSize = $true
$form.Controls.Add($lblInput)
$txtInput = New-Object System.Windows.Forms.TextBox
$txtInput.Location = New-Object System.Drawing.Point(12, 70)
$txtInput.Size = New-Object System.Drawing.Size(800, 20)
$txtInput.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtInput)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Look Up"
$btnRun.Location = New-Object System.Drawing.Point(12, 98)
$btnRun.Size = New-Object System.Drawing.Size(110, 28)
$form.Controls.Add($btnRun)
$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Clear"
$btnClear.Location = New-Object System.Drawing.Point(130, 98)
$btnClear.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClear)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(228, 98)
$btnClose.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClose)
$form.CancelButton = $btnClose

# Find box, on the SAME ROW as Look Up / Clear / Close.
#
# Placed here rather than on a row of its own because the console is already
# tall enough to be capped against the working area on a 1080p screen, and a
# search box is not worth 30px of everything else. The buttons end at x=318 and
# the panel is nearly 800 wide, so the space to the right was doing nothing.
#
# Width is set later, with the other PanelW-derived controls -- PanelW is not
# known until Sections.json has been read, and referencing it here would give a
# $null that WinForms turns into a width of zero.
$lblFind = New-Object System.Windows.Forms.Label
$lblFind.Text = 'Find tool:'
$lblFind.Location = New-Object System.Drawing.Point(340, 103)
$lblFind.Size = New-Object System.Drawing.Size(62, 18)
$form.Controls.Add($lblFind)

$txtFind = New-Object System.Windows.Forms.TextBox
$txtFind.Location = New-Object System.Drawing.Point(404, 99)
$txtFind.Size = New-Object System.Drawing.Size(200, 24)
$form.Controls.Add($txtFind)

$btnFindClear = New-Object System.Windows.Forms.Button
$btnFindClear.Text = 'Show all'
$btnFindClear.Location = New-Object System.Drawing.Point(612, 98)
$btnFindClear.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnFindClear)

$lblFindCount = New-Object System.Windows.Forms.Label
$lblFindCount.Location = New-Object System.Drawing.Point(710, 103)
$lblFindCount.Size = New-Object System.Drawing.Size(120, 18)
$form.Controls.Add($lblFindCount)

$lblResult = New-Object System.Windows.Forms.Label
$lblResult.Text = "Not looked up yet."
$lblResult.Location = New-Object System.Drawing.Point(12, 136)
$lblResult.Size = New-Object System.Drawing.Size(800, 20)
$lblResult.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblResult)

$txtOut = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline = $true
$txtOut.ReadOnly = $true
$txtOut.ScrollBars = "Vertical"
$txtOut.Location = New-Object System.Drawing.Point(12, 160)
$txtOut.Size = New-Object System.Drawing.Size(800, 140)
$txtOut.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtOut.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.Controls.Add($txtOut)

# Sections are built from Sections.json rather than hardcoded, so a new
# section can be added without touching this script. Geometry is derived:
# every section is ButtonsPerRow buttons wide and as many rows as it declares,
# and the form grows or shrinks to fit whatever total that comes to.
$script:SectionsConfigPath = Get-ConfigFilePath -Name 'Sections.json'
$script:BtnH = $script:ConsoleBtnHeight; $script:BtnPad = $script:ConsoleBtnPad

# The tool list is loaded HERE, several hundred lines above where the buttons
# are built, because button width is measured from the captions and the
# measurement has to happen before any Size() is computed.
#
# This is as high as it can go. The failure path reports into $txtOut, so
# $txtOut has to exist first -- moving this any earlier would turn "Tools.json
# is missing" into a null reference, which is a worse message for the same
# problem.
$script:AllTools = @()
try {
    $script:AllTools = Get-ToolDefinitions -Path $ToolsConfigPath
} catch {
    $script:AllTools = @()
    $txtOut.Text = $_.Exception.Message
}

# PanelW is DERIVED, not chosen, and it allows for a scrollbar that may not be
# there.
#
# It was a flat 800. A button occupies BtnW + Margin on both sides = 158, so a
# row of five needs 790 and the panel's client area was 798 -- five fitted, with
# eight pixels to spare. Add a vertical scrollbar and the client area drops by
# SystemInformation.VerticalScrollBarWidth (17 at 100% DPI) to 781, nine pixels
# short, so the fifth button wrapped to the next row.
#
# That is why ONE section looked different: Admin Tools holds 26 buttons in a
# 4-row panel and is the only section over its declared capacity, so it is the
# only one with a scrollbar. The setting said five across and four of the five
# sections agreed, which made it look like Admin Tools was special-cased
# somewhere. It was not -- it was the only one wide enough to notice.
#
# Reserving the scrollbar width unconditionally costs 17px of window and makes
# the row count independent of how full a section happens to be. A layout that
# changes shape when content grows is a layout that will surprise someone
# mid-demo.
function Get-SectionDefinitions {
    $default = [PSCustomObject]@{
        ButtonsPerRow = 5
        Sections = @(
            [PSCustomObject]@{ Name = 'Connect Scripts'; Rows = 1 }
            [PSCustomObject]@{ Name = 'Admin Tools';     Rows = 4 }
            [PSCustomObject]@{ Name = 'Daily Scripts';   Rows = 1 }
        )
    }
    if (-not (Test-Path -LiteralPath $script:SectionsConfigPath)) { return $default }
    try {
        $cfg = Get-Content -LiteralPath $script:SectionsConfigPath -Raw | ConvertFrom-Json
        if (-not $cfg.Sections -or @($cfg.Sections).Count -eq 0) { return $default }
        if (-not $cfg.ButtonsPerRow) {
            Add-Member -InputObject $cfg -NotePropertyName ButtonsPerRow -NotePropertyValue 5 -Force
        }
        return $cfg
    } catch { return $default }
}

$script:SectionCfg    = Get-SectionDefinitions
$script:SectionPanels = @{}

# Labels and declared row counts are kept alongside the panels so the search
# filter can collapse a whole section when nothing in it matches. Without the
# label reference a filtered console shows a heading over an empty box, which
# reads as "these tools disappeared" rather than "nothing here matched".
$script:SectionLabels = @{}
$script:SectionRows   = @{}
$script:SectionStartY = 308
$script:SectionNames  = @(@($script:SectionCfg.Sections) | ForEach-Object { "$($_.Name)" })

$script:PerRow  = [int]$script:SectionCfg.ButtonsPerRow
if ($script:PerRow -lt 1) { $script:PerRow = 5 }

# Button width is MEASURED, not chosen.
#
# It was a flat 150 -- a number that fitted the captions of the day, in the
# font of the machine this was written on, at 100% scaling. None of those three
# things is a constant, so 150 was wrong in both directions at once: wider than
# short captions need, and too narrow for any long one added later.
#
# Get-ConsoleButtonWidth (Common.ps1) returns the narrowest width at which
# every caption still fits INSIDE THE BOX, wrapping allowed. Wrapping has to be
# allowed: 'Mailbox / Onedrive Perms' has always been two lines, and forcing it
# onto one would set the width for all forty-odd buttons and make the window
# considerably wider.
#
# Every tool in Tools.json is measured, including ones a capability check will
# later hide. Which tools are hidden is not known yet -- the probe has not run
# -- and measuring a button that turns out to be invisible costs a few pixels,
# where skipping it would clip a caption the first time a capability was
# switched on.
#
# ButtonWidth in Sections.json overrides the measurement, and is deliberately
# NOT in the Sections dialog. It is there for the case where the fit is correct
# but not to taste; a layout setting that is easy to reach is one that gets
# changed by accident and then reported as a bug.
$cfgW = 0
if ($script:SectionCfg.PSObject.Properties['ButtonWidth']) {
    [int]::TryParse("$($script:SectionCfg.ButtonWidth)", [ref]$cfgW) | Out-Null
}
if ($cfgW -ge 60) {
    $script:BtnW = $cfgW
} else {
    $script:BtnW = Get-ConsoleButtonWidth `
        -Captions @(@($script:AllTools) | ForEach-Object { "$($_.Name)" }) `
        -Font $form.Font `
        -ButtonHeight $script:BtnH
}

$script:BtnCell = $script:BtnW + ($script:BtnPad * 2)

# Start NARROW. Update-SectionPanelWidth widens later if a scrollbar actually
# turns up.
#
# Scrollbar width used to be reserved unconditionally, which cost ~27px down the
# right-hand edge of every panel permanently -- a FlowLayoutPanel is
# left-aligned, so width nothing uses collects on the right and reads as the
# layout being lopsided.
#
# Unconditional was itself a fix, and the bug behind it is still real: a panel
# that grows a scrollbar loses that width from its client area and wraps its
# last button, so the fullest section looked special-cased when it was only the
# fullest. Predicting which sections will overflow gets that right most of the
# time and wrong exactly where it matters -- capability-hidden tools are not
# known at this point, so a count taken here can be higher than what is drawn.
#
# So this no longer predicts. It lays out narrow, asks the panels whether they
# grew scrollbars, and widens every panel together if any of them did. See
# Update-SectionPanelWidth.
$script:PanelW = Get-ConsolePanelWidth -ButtonWidth   $script:BtnW `
                                       -ButtonPad     $script:BtnPad `
                                       -ButtonsPerRow $script:PerRow

# Widened HERE rather than where they were created, because PanelW is not known
# until Sections.json has been read and Get-SectionDefinitions is defined below
# the controls. Referencing it at creation time gave a $null that WinForms
# silently converts to a width of zero -- the input and output boxes would have
# vanished rather than erred.
$txtInput.Width  = $script:PanelW
$lblResult.Width = $script:PanelW
$txtOut.Width    = $script:PanelW

# The find row is laid out from the RIGHT edge of the panel so it stays put as
# ButtonsPerRow changes the console width -- fixed left offsets would strand it
# mid-form at 7 columns and run it over the Close button at 4.
#
# TWO LAYOUTS, because at four or five columns there is genuinely not enough
# room on the button row. The first version of this "handled" that by falling
# back to a fixed 150px box, which at 4 columns put the textbox at 404..554
# inside a panel ending at 540 -- off the edge, on top of the count label. A
# fallback that produces an overlapping layout is not a fallback, it is the
# same bug with an if in front of it.
#
# So: share the button row when it fits, take a row of its own when it does
# not, and move everything below down by exactly that much.
$findRight = 12 + $script:PanelW
$findLeft  = 336
$boxW      = $findRight - 214 - 8 - $findLeft - 64

if ($boxW -ge 90) {
    $lblFind.Location      = New-Object System.Drawing.Point($findLeft, 103)
    $txtFind.Location      = New-Object System.Drawing.Point(($findLeft + 64), 99)
    $txtFind.Size          = New-Object System.Drawing.Size($boxW, 24)
    $btnFindClear.Location = New-Object System.Drawing.Point(($findRight - 214), 98)
    $lblFindCount.Location = New-Object System.Drawing.Point(($findRight - 118), 103)
    $lblFindCount.Size     = New-Object System.Drawing.Size(118, 18)
} else {
    $rowY = 132
    $lblFind.Location      = New-Object System.Drawing.Point(12, ($rowY + 5))
    $txtFind.Location      = New-Object System.Drawing.Point(76, ($rowY + 1))
    $txtFind.Size          = New-Object System.Drawing.Size([Math]::Max(120, ($script:PanelW - 300)), 24)
    $btnFindClear.Location = New-Object System.Drawing.Point(($findRight - 214), $rowY)
    $lblFindCount.Location = New-Object System.Drawing.Point(($findRight - 118), ($rowY + 5))
    $lblFindCount.Size     = New-Object System.Drawing.Size(118, 18)

    # Everything below the find row moves down together. SectionStartY is read
    # by Update-ConsoleSearch when it reflows, so shifting it here is enough --
    # the filter does not need to know which layout is in use.
    $shift = 34
    $lblResult.Location    = New-Object System.Drawing.Point(12, (136 + $shift))
    $txtOut.Location       = New-Object System.Drawing.Point(12, (160 + $shift))
    $script:SectionStartY += $shift
}

$y = $script:SectionStartY
foreach ($sec in @($script:SectionCfg.Sections)) {
    $rows = [int]$sec.Rows
    if ($rows -lt 1) { $rows = 1 }
    $script:SectionRows["$($sec.Name)"] = $rows
    $panelH = ($rows * ($script:BtnH + ($script:BtnPad * 2))) + 12

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "$($sec.Name)"
    $lbl.Location = New-Object System.Drawing.Point(12, $y)
    $lbl.Size = New-Object System.Drawing.Size($script:PanelW, 20)
    $form.Controls.Add($lbl)

    $flow = New-Object System.Windows.Forms.FlowLayoutPanel
    $flow.Location = New-Object System.Drawing.Point(12, ($y + 22))
    $flow.Size = New-Object System.Drawing.Size($script:PanelW, $panelH)

    # Half the slack goes on the left so the spare width reads as a margin on
    # both edges rather than a gap on one. A FlowLayoutPanel lays out from its
    # left padding edge, so this shifts the grid right by half and leaves the
    # other half past the last column.
    $flow.Padding = New-Object System.Windows.Forms.Padding(
                        [int]($script:ConsolePanelSlack / 2), 0, 0, 0)

    $flow.AutoScroll = $true
    $flow.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $form.Controls.Add($flow)

    $script:SectionPanels["$($sec.Name)"] = $flow
    $script:SectionLabels["$($sec.Name)"] = $lbl
    $y += 22 + $panelH + 8
}

# Grow the window to whatever the sections actually needed.
#
# The tooltip provider is created here rather than further down because the
# version footer below needs it. It was originally below, which meant the
# footer called SetToolTip on a null and threw -- and the catch turned that
# ordering mistake into "version unavailable" on every launch. A catch that
# reports the wrong cause is worse than no catch.
$toolTip = New-Object System.Windows.Forms.ToolTip

# Version footer, below the last section panel. Added before ClientSize is
# fixed so the form grows to include it rather than clipping it.
#
# Shows the version AND whether the files still match it -- because the number
# alone is a claim. Green when everything matches, amber when files differ or
# the manifest is stale, grey when the copy was never stamped.
# Width is the form's, less both margins, so the line has the whole footer to
# work with rather than the 600px it started with. AutoEllipsis makes any
# remaining overflow end in "..." -- a visible signal that there is more to
# read, instead of a sentence that silently stops mid-word.
$lblVersion              = New-Object System.Windows.Forms.Label
$lblVersion.Tag          = 'NoTheme'
$lblVersion.Location     = New-Object System.Drawing.Point(12, ($y + 4))
$lblVersion.Size         = New-Object System.Drawing.Size($script:PanelW, 18)
$lblVersion.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif', 8)
$lblVersion.AutoEllipsis = $true
$form.Controls.Add($lblVersion)
$y += 26

try {
    $vv = Test-ToolkitIntegrity
    if (-not $vv.HasManifest) {
        $lblVersion.Text      = 'Admin Toolkit -- unversioned copy (no Version.json)'
        $lblVersion.ForeColor = [System.Drawing.Color]::DimGray
    } elseif ($vv.Modified.Count -eq 0 -and $vv.Missing.Count -eq 0 -and $vv.Extra.Count -eq 0) {
        $lblVersion.Text      = "Admin Toolkit v$($vv.Version)  --  $($vv.Matched) files verified  --  released $($vv.Released)"
        $lblVersion.ForeColor = [System.Drawing.Color]::FromArgb(80, 220, 130)
    } else {
        # Built from the parts rather than reusing $vv.Summary. Summary already
        # opens with "v1.03 --", so pasting it after the name printed the
        # version twice -- and it is worded for the Tester report, which has
        # room to spare. This line does not.
        #
        # Only non-zero categories are named. "1 modified" says the same thing
        # as "1 modified, 0 missing, 0 extra" in a third of the width, and the
        # zeroes were the padding that pushed this off the edge.
        $bits = @()
        if ($vv.Modified.Count -gt 0) { $bits += "$($vv.Modified.Count) modified" }
        if ($vv.Missing.Count  -gt 0) { $bits += "$($vv.Missing.Count) missing" }
        if ($vv.Extra.Count    -gt 0) { $bits += "$($vv.Extra.Count) extra" }
        $detail = if ($bits.Count -gt 0) { $bits -join ', ' } else { 'differs from the manifest' }
        if ($vv.Stale) { $detail += ' (manifest older -- re-stamp)' }

        $lblVersion.Text      = "Admin Toolkit v$($vv.Version)  --  $detail  --  run Tester for the list"
        $lblVersion.ForeColor = [System.Drawing.Color]::FromArgb(230, 170, 60)
    }
} catch {
    $lblVersion.Text      = 'Admin Toolkit -- version unavailable'
    $lblVersion.ForeColor = [System.Drawing.Color]::DimGray
}

# Outside the try deliberately: a tooltip that fails to attach must not be able
# to overwrite a version line that was worked out correctly.
$toolTip.SetToolTip($lblVersion, "Version and per-file SHA-256 come from Version.json.`r`nRun Tester to see exactly which files differ from the master.")

# Width follows PanelW plus the 12px margin on each side, rather than the 824
# it was hardcoded to. PanelW is now derived from ButtonsPerRow, so changing
# that setting to 4 or 6 resizes the window instead of clipping the panels.
$form.ClientSize = New-Object System.Drawing.Size(($script:PanelW + 24), ($y + 8))

# Never taller than the screen it opens on.
#
# The height is the sum of whatever Sections.json declares, so it has always
# been unbounded -- add a section or a row and the window grows, with nothing
# stopping it growing past the bottom of the display. Raising the button height
# to 38 for two-line captions added six pixels per row, eleven rows, sixty-six
# pixels, which is enough to put the version footer and the last section under
# the taskbar on a 1080p screen.
#
# AutoScroll rather than a smaller layout: the buttons stay legible and the
# window scrolls. WorkingArea already excludes the taskbar.
$form.AutoScroll = $true
$work = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea
if ($form.Height -gt $work.Height) {
    $form.Height = $work.Height
    $form.Top    = $work.Top
}
if ($form.Width -gt $work.Width) { $form.Width = $work.Width }
#endregion

#region State + lookup ----------------------------------------------------------
$script:LastSmtp = $null
$script:LastType = $null
$script:LastADServer = $null
$script:LastEnvironment = $null
# $script:AllTools is loaded ABOVE, in the section geometry region, and is
# deliberately NOT re-read here.
#
# It had to move: button width is measured from the tool captions now, so the
# list has to exist before the first Size() is computed. Leaving a second load
# here would re-read the file on every launch and, on the failure path, blank a
# list the layout had already been built from.

# Find-TargetEnvironment comes from Common.ps1.
#
# A local copy lived here and shadowed it. It had gone stale: it still
# pinned $script:DefaultDC where Common uses (Get-TenantDirectoryServer)
# -- the pinning removed after the write-scope bug (14f). The console was
# quietly running the pre-fix version of a function it dot-sources.
# See ARCHITECTURE 14n.
function Get-ToolParameterSet {
    <#
        Which of the console's context parameters a tool actually declares.

        Passing a parameter a script does not declare is a BINDING ERROR, and
        the failure mode is asymmetric in a way that hides it: a script with no
        param() block at all absorbs anything into $args and works, so most of
        the toolkit tolerated being handed all three. A script that declares
        SOME of them rejects the rest outright -- which is how a newly added
        tool can fail on its very first click while every older one is fine.

        Same approach Invoke-CasQuery uses against cmdlet surfaces: ask what is
        declared, pass only that.
    #>
    param([Parameter(Mandatory)][string]$ScriptPath)

    # EMPTY IS NOT A VALUE. Context the console does not have is left OUT of the
    # splat rather than passed as ''.
    #
    # Passing it defeats the script's own default: an explicitly supplied empty
    # string binds, and binding happens BEFORE the default can apply. A
    # parameter with [ValidateSet] then fails outright -- 2-Export-Usage died
    # with 'the argument "" does not belong to the set' -- naming the
    # directories it did allow -- when the
    # operator simply had not looked anything up yet.
    #
    # "I do not know which directory" and "the directory is the empty string"
    # are different statements, and only one of them is true here.
    $ctx = [ordered]@{}
    if ("$($script:LastSmtp)".Trim())        { $ctx['SmtpAddress']      = $script:LastSmtp }
    if ("$($script:LastEnvironment)".Trim()) { $ctx['Environment']      = $script:LastEnvironment }
    if ("$($script:LastADServer)".Trim())    { $ctx['DomainController'] = $script:LastADServer }
    try {
        $cmd = Get-Command -Name $ScriptPath -CommandType ExternalScript -ErrorAction Stop
        # No param() block means no declared parameters -- and no rejection
        # either, so pass the lot and let $args have them.
        if ($cmd.Parameters.Count -eq 0) { return $ctx }

        $out = @{}
        foreach ($k in $ctx.Keys) {
            if ($cmd.Parameters.ContainsKey($k)) { $out[$k] = $ctx[$k] }
        }
        return $out
    } catch {
        # Unreadable script: fall back to the old behaviour rather than
        # refusing to launch. A tool that might work beats one that cannot.
        return $ctx
    }
}

function Test-ToolDeclaresOwnProcess {
    <#
        Does the script itself insist on running in its own process?

        Tools.json has a NewProcess checkbox, and on 2026-08-06 an on-premises
        tool ran with it unticked. It loaded into this console's process, where
        Exchange Online was already connected. The two share cmdlet names, so
        every write went to the cloud and failed with a message that read like
        a permissions problem.

        The operator did not do anything wrong. A checkbox that must be
        remembered for the tool to be CORRECT -- as opposed to merely
        convenient -- is a trap, and the same trap as the hidden DC selector
        that went on applying its default after it stopped being visible.

        So the script declares the requirement and the console obeys it. The
        checkbox still works, and still means "I would like this in its own
        window", but it can no longer be the only thing standing between an
        on-premises tool and the wrong directory.

        Declared by putting this on any line in the first 40 lines:

            # TOOLKIT-REQUIRES: OwnProcess

        Read from the file rather than from Tools.json on purpose: the file is
        the thing that knows, and it stays true when the tool is copied into
        another tenant's toolkit with a fresh Tools.json.
    #>
    param([Parameter(Mandatory)][string]$ScriptPath)

    if ($null -eq $script:OwnProcessCache) { $script:OwnProcessCache = @{} }
    if ($script:OwnProcessCache.ContainsKey($ScriptPath)) { return $script:OwnProcessCache[$ScriptPath] }

    $result = $false
    try {
        # Header only. This runs on every click, and reading whole files to
        # answer a header question is how the cold-launch scan got slow (14c).
        $head = Get-Content -LiteralPath $ScriptPath -TotalCount 40 -ErrorAction Stop
        $result = [bool]($head -match 'TOOLKIT-REQUIRES:\s*OwnProcess')
    } catch { }

    $script:OwnProcessCache[$ScriptPath] = $result
    return $result
}

function Start-ToolScriptInNewProcess {
    <#
        Launches a tool in its OWN PowerShell process, connecting to
        on-premises Exchange first.

        Why this is safe here when it was not for the connect buttons: a
        connect script has to leave its session in the console's process, so a
        child process was useless. A TOOL only needs the session itself -- it
        connects, does the work, and exits. Nothing has to cross the boundary.

        Two things fall out for free. The child has its own thread and message
        loop, so sign-in and credential dialogs behave normally (section 14d).
        And on-premises cmdlets are loaded only there, so Get-Mailbox in the
        child unambiguously means on-prem while the console's stays cloud.

        The cost is real and paid every launch: a fresh Import-PSSession is
        tens of seconds before the tool's window appears.
    #>
    param([Parameter(Mandatory)]$Tool, [Parameter(Mandatory)][string]$ScriptPath)

    # Launch the same PowerShell that is running us, not whatever
    # 'powershell.exe' resolves to on PATH.
    $psExe = $null
    try { $psExe = (Get-Process -Id $PID).Path } catch { }
    if (-not $psExe) { $psExe = 'powershell.exe' }

    # Every path and value is single-quoted with internal quotes doubled, so
    # the whole -Command payload can be wrapped in double quotes without
    # anything inside fighting it. Start-Process does NOT quote for you --
    # the same trap that once sent Edge to a profile called 'Profile'.
    $esc = { param([string]$s) "'" + ("$s" -replace "'", "''") + "'" }

    $parts = New-Object System.Collections.ArrayList

    # Shrink the console window before anything prints.
    #
    # It only ever shows connect output and whatever the tool writes, and at the
    # default size it covers most of the screen -- including the tool's own
    # window, which is the thing you actually wanted.
    #
    # Added FIRST so it runs before the connect banner prints, not after.
    [void]$parts.Add((Get-ConsoleResizeCommand))

    $connect = Join-Path $PSScriptRoot 'Connect-ExSE.ps1'
    if (Test-Path -LiteralPath $connect) {
        [void]$parts.Add("& $(& $esc $connect)")
    }

    $call = "& $(& $esc $ScriptPath)"
    foreach ($kv in (Get-ToolParameterSet -ScriptPath $ScriptPath).GetEnumerator()) {
        if ("$($kv.Value)".Trim()) { $call += " -$($kv.Key) $(& $esc $kv.Value)" }
    }
    [void]$parts.Add($call)

    # -NoExit so the window survives the tool closing: the output is the record
    # of what happened, and a window that vanishes takes it with it.
    $argLine = '-NoExit -ExecutionPolicy Bypass -Command "' + ($parts -join '; ') + '"'
    Start-Process -FilePath $psExe -ArgumentList $argLine
}

# Resolve-PairedToolScript and the whole ScriptOnPrem "AD twin" feature were
# REMOVED on 2026-08-13. See ARCHITECTURE section 30.
#
# The short version: a tool could name an on-premises twin and become one
# button that asked which directory to act on. It worked, but ToolManager's
# Move Up / Move Down swapped a hand-maintained list of columns that did not
# include ScriptOnPrem -- so reordering left the twin behind on the old row
# position and it silently reattached to whatever tool landed there. Message
# Trace, DL Send To/As and Connect-ExSE all acquired twins they never had, and
# Save wrote the corruption back.
#
# The pairing was also invisible: nothing on the console said a button had two
# scripts behind it, so the wrong one running looked like the tool
# misbehaving rather than like config.
#
# Grouping (the Group column) does the same job out loud -- several tools
# behind one button, with a chooser that names them -- so pairs become groups.

function Start-ToolScript {
    param($Tool)

    $scriptPath = Join-Path $PSScriptRoot $Tool.Script
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "'$($Tool.Name)' points to '$(Split-Path -Leaf $scriptPath)', which does not exist under this console's folder.`n`nThis is expected if this is a reduced-tier distribution that doesn't include every tool -- otherwise, check Tools.json (ToolManager.ps1) for a stale/typo'd path.",
            'Tool Script Not Found', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    # The script's own declaration wins over the checkbox. See
    # Test-ToolDeclaresOwnProcess for why this is not merely a convenience.
    # THE SCRIPT DECIDES. There is no longer a Tools.json checkbox in the or.
    #
    # NewProcess started as the only control and was demoted to a hint when an
    # unticked on-premises tool wrote to the cloud (see Test-ToolDeclaresOwnProcess).
    # From then on every tool that needed isolation declared it, and the
    # checkbox never decided anything again -- all 55 rows sat unticked for
    # months with no symptom, because the declarations were doing the work.
    #
    # Leaving it meant one requirement expressed in two places, with the
    # weaker copy editable and visible. That is a question waiting to be asked
    # by whoever opens Tool Manager next, and the honest answer was "ignore
    # that column".
    $needsOwnProcess = Test-ToolDeclaresOwnProcess -ScriptPath $scriptPath

    try {
        if ($needsOwnProcess) {
            Start-ToolScriptInNewProcess -Tool $Tool -ScriptPath $scriptPath
            $txtOut.Text = "'$($Tool.Name)' opened in its own PowerShell window." +
                [Environment]::NewLine +
                'It connects to on-premises Exchange first, so the tool takes a while to appear.' +
                [Environment]::NewLine +
                'That window keeps its own connections -- this console is unaffected.'
            return
        }
        # @toolParams, not @(...). Splatting needs a VARIABLE holding the
        # dictionary; @(...) is an array subexpression and would hand the whole
        # thing over as one positional argument.
        $toolParams = Get-ToolParameterSet -ScriptPath $scriptPath
        & $scriptPath @toolParams
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context "Launching $($Tool.Name)"
    }
}

$script:KnownObjectTypes = @('Mailbox', 'UnifiedGroup', 'OnpremGroup')

function Get-ToolSection {
    param($Tool)
    $s = "$($Tool.Section)".Trim()
    foreach ($n in $script:SectionNames) {
        if ($s -eq $n) { return $n }
    }
    # Unknown or blank section -- fall back to the first declared section so a
    # tool can never become invisible just because its section was renamed.
    return $script:SectionNames[0]
}

function Get-ToolGroupName {
    param($Tool)
    if (-not $Tool) { return '' }
    if (-not $Tool.PSObject.Properties['Group']) { return '' }
    return "$($Tool.Group)".Trim()
}

function Show-ToolGroupChooser {
    <#
        The tools behind a group button.

        A group exists to buy back console real estate: five room tools are
        five buttons' worth of grid for something touched twice a month. One
        button, one extra click, and the front page stays readable.

        The chooser repeats each tool's availability rather than just listing
        names. A tool that cannot run is disabled HERE too, with the same
        reason -- otherwise the group button becomes a place where greying goes
        to die, and the operator learns the reason one dialog too late.
    #>
    param([Parameter(Mandatory)][string]$GroupName)

    $tools = @($script:ToolGroups[$GroupName])
    if ($tools.Count -eq 0) { return }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $GroupName
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false
    $dlg.StartPosition = 'CenterParent'
    $dlg.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $pad = 12
    $bw  = 260
    $bh  = 34
    $y   = $pad

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "$($tools.Count) tools"
    $lbl.Location = New-Object System.Drawing.Point($pad, $y)
    $lbl.Size = New-Object System.Drawing.Size($bw, 18)
    $dlg.Controls.Add($lbl)
    $y += 24

    $script:GroupPick = $null
    foreach ($t in $tools) {
        $b = New-Object System.Windows.Forms.Button
        $b.Text = "$($t.Name)"
        $b.Location = New-Object System.Drawing.Point($pad, $y)
        $b.Size = New-Object System.Drawing.Size($bw, $bh)
        $b.Tag = $t

        $block = Get-ToolCapabilityBlock -Tool $t
        if ($block) {
            $b.Enabled = $false
            $tip = New-Object System.Windows.Forms.ToolTip
            $tip.SetToolTip($b, $block.Reason)
        }
        $b.Add_Click({
            $script:GroupPick = $this.Tag
            $this.FindForm().DialogResult = [System.Windows.Forms.DialogResult]::OK
            $this.FindForm().Close()
        })
        $dlg.Controls.Add($b)
        $y += $bh + 6
    }

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'
    $cancel.Location = New-Object System.Drawing.Point($pad, ($y + 6))
    $cancel.Size = New-Object System.Drawing.Size($bw, 28)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($cancel); $dlg.CancelButton = $cancel

    $dlg.ClientSize = New-Object System.Drawing.Size(($bw + ($pad * 2)), ($y + 46))
    Set-ConsoleTheme -Form $dlg
    [void]$dlg.ShowDialog()

    if ($script:GroupPick) {
        $picked = $script:GroupPick
        $script:GroupPick = $null
        Start-ToolScript -Tool $picked
        Update-ConnectionButtonColors
    }
}

function Add-ToolButtonsToPanel {
    param($Panel, $Tools, [string]$EmptyText)
    if (@($Tools).Count -eq 0) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $EmptyText
        $lbl.AutoSize = $true
        $Panel.Controls.Add($lbl)
        return
    }

    # Group buttons are inserted AT THE POSITION OF THE FIRST MEMBER, and the
    # members are still created -- just hidden. Two reasons for building them
    # rather than skipping them: the search filter reveals a matching member
    # directly, which is how you reach a grouped tool by name without going
    # through the group; and the grey-out pass can compute the group's own
    # state from the members it already has in front of it.
    $seen = @{}
    foreach ($tool in $Tools) {
        $grp = Get-ToolGroupName -Tool $tool
        if ($grp -and -not $seen.ContainsKey($grp)) {
            $seen[$grp] = $true
            $gb = New-Object System.Windows.Forms.Button
            $gb.Text   = $grp
            $gb.Size   = New-Object System.Drawing.Size($script:BtnW, $script:BtnH)
            $gb.Margin = New-Object System.Windows.Forms.Padding($script:BtnPad)
            $gb.Tag    = [PSCustomObject]@{ IsToolGroup = $true; GroupName = $grp }
            $gb.Add_Click({ Show-ToolGroupChooser -GroupName $this.Tag.GroupName })
            $Panel.Controls.Add($gb)
        }

        $btn = New-Object System.Windows.Forms.Button

        # Sized from the script constants, not from literals.
        #
        # This said Size(150, 32) and Padding(4) while $script:BtnW/BtnH/BtnPad
        # sat three hundred lines above holding the same numbers -- so the
        # constants were decorative, the panel height was computed from one set
        # of values and the buttons drawn from another, and raising BtnH changed
        # the panels without changing a single button.
        #
        # The height went 32 -> 38 because a two-line caption did not fit.
        # 'Mailbox / Onedrive Perms' wraps, and at 32px the first line was
        # clipped along its top edge.
        $btn.Text = $tool.Name
        $btn.Size = New-Object System.Drawing.Size($script:BtnW, $script:BtnH)
        $btn.Margin = New-Object System.Windows.Forms.Padding($script:BtnPad)
        $btn.Tag = $tool
        $btn.Add_Click({
            $tool = $this.Tag

            # A connect script CANNOT run from here.
            #
            # ShowDialog() holds this thread in a modal message loop, and a
            # modal form disables every other window on its thread. The
            # Exchange Online sign-in prompt is created on this thread, so it
            # draws and then never accepts input -- confirmed 2026-08-03 by
            # running the same script from a bare prompt, where it works.
            # -DisableWAM was not enough: it moves sign-in off the WAM broker
            # onto an EMBEDDED browser, which is still a window on this thread.
            #
            # So the button records what to run and closes the form. The work
            # happens after ShowDialog() returns, where there is no modal loop
            # -- and in this SAME process, which is the whole point: a
            # connection made anywhere else would not be ours.
            if (Get-ConnectButtonService -Tool $tool) {
                $script:PendingConnect = $tool
                $form.Close()
                return
            }

            Start-ToolScript -Tool $tool
            # Tools run in THIS process via '&', so a tool that connected to
            # something has changed what every other tool can do. The moment a
            # tool returns is the reliable re-probe point. Full probe including
            # SharePoint -- a tool finishing is infrequent and user-initiated,
            # so the cost is fine here in a way it would not be on every Look Up.
            Update-ConnectionButtonColors
        })
        # Hidden, not skipped -- see the note above the group button.
        if ($grp) { $btn.Visible = $false }
        $Panel.Controls.Add($btn)
    }
}

function Build-ToolButtons {
    # Membership is worked out ONCE, from the whole tool list, before any panel
    # is built. Deriving it per panel would scatter a group across sections
    # silently if two members disagreed about which section they belong to.
    $script:ToolGroups = @{}
    foreach ($t in @($script:AllTools)) {
        $g = Get-ToolGroupName -Tool $t
        if (-not $g) { continue }
        if (-not $script:ToolGroups.ContainsKey($g)) { $script:ToolGroups[$g] = @() }
        $script:ToolGroups[$g] += $t
    }

    foreach ($name in $script:SectionNames) {
        $panel = $script:SectionPanels[$name]
        if (-not $panel) { continue }
        $panel.Controls.Clear()
        $tools = @($script:AllTools | Where-Object { (Get-ToolSection $_) -eq $name })
        Add-ToolButtonsToPanel -Panel $panel -Tools $tools `
            -EmptyText "No tools in '$name'. Use ToolManager.ps1 to add some."
    }
}

function Update-ConsoleSearch {
    <#
        Filter the tool buttons by caption.

        HIDES rather than greys. Greying is already taken -- it means "this tool
        cannot run right now", with a tooltip explaining which of three reasons.
        A second meaning on the same signal would make both useless, and a grid
        of fifty dimmed buttons is not a search result. A tool that does not
        match simply is not there until the box is cleared.

        Visible is safe to own here because nothing else touches it: capability
        blocking and object-type mismatch both work through Enabled. That
        separation is why search and grey-out can coexist without a combined
        state machine -- checked before this was written, not assumed.

        SECTIONS WITH NO MATCH COLLAPSE ENTIRELY, label included, and everything
        below moves up. A heading above an empty box reads as "the tools in here
        vanished"; no heading reads as "nothing in here matched", which is what
        actually happened.

        Matching is a plain case-insensitive substring on the caption, not a
        regex. An operator typing 'perms' is searching, not writing a pattern,
        and a stray bracket should not throw.
    #>
    param([string]$Text)

    $q = "$Text".Trim()
    $filtering = [bool]$q

    $shown = 0
    $total = 0
    $y = $script:SectionStartY

    foreach ($name in $script:SectionNames) {
        $panel = $script:SectionPanels[$name]
        $lbl   = $script:SectionLabels[$name]
        if (-not $panel) { continue }

        $secVisible = 0
        foreach ($ctrl in @($panel.Controls)) {
            # Labels inside a panel are the "no tools in this section" notice,
            # not buttons. They follow the section, not the filter.
            if (-not $ctrl.Tag) { continue }

            # GROUPS INVERT WHILE FILTERING.
            #
            # Unfiltered, the group button is shown and its members are hidden
            # -- that is the whole point of grouping. Filtered, the group button
            # goes and the matching members appear in its place, so a search for
            # 'roomplace' reaches the tool directly instead of stopping at a
            # button called 'Rooms' that gives no hint of what is inside.
            $isGroup = [bool]$ctrl.Tag.PSObject.Properties['IsToolGroup']
            $grouped = (-not $isGroup) -and [bool](Get-ToolGroupName -Tool $ctrl.Tag)

            if (-not $filtering) {
                $ctrl.Visible = -not $grouped
                if ($ctrl.Visible) { $secVisible++ }
                if (-not $isGroup) { $total++ }
                if ($ctrl.Visible -and -not $isGroup) { $shown++ }
                continue
            }

            if ($isGroup) { $ctrl.Visible = $false; continue }
            $total++

            # Script name as well as caption, because the two are often the only
            # thing an operator remembers -- 'not' finds Returning User through
            # NotDeadYet.ps1, which no amount of searching the caption would.
            $hay = "$($ctrl.Text)"
            if ($ctrl.Tag.PSObject.Properties['Script']) { $hay += " $($ctrl.Tag.Script)" }
            $match = ($hay.IndexOf($q, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
            $ctrl.Visible = $match
            if ($match) { $secVisible++; $shown++ }
        }

        $rows = [int]$script:SectionRows[$name]
        if ($rows -lt 1) { $rows = 1 }

        # While filtering, a section only needs as many rows as it is showing.
        # Keeping the declared height would leave a filtered console mostly
        # empty space and push the last section off the bottom.
        if ($filtering) {
            $need = [int][Math]::Ceiling($secVisible / [double][Math]::Max(1, $script:PerRow))
            if ($need -lt 1) { $need = 1 }
            if ($need -lt $rows) { $rows = $need }
        }
        $panelH = ($rows * ($script:BtnH + ($script:BtnPad * 2))) + 12

        $empty = ($secVisible -eq 0)
        if ($lbl) { $lbl.Visible = -not $empty }
        $panel.Visible = -not $empty
        if ($empty) { continue }

        if ($lbl) { $lbl.Location = New-Object System.Drawing.Point(12, $y) }
        $panel.Location = New-Object System.Drawing.Point(12, ($y + 22))
        $panel.Height   = $panelH
        $y += 22 + $panelH + 8
    }

    if ($lblVersion) {
        $lblVersion.Location = New-Object System.Drawing.Point(12, ($y + 4))
        $y += 26
    }

    $form.ClientSize = New-Object System.Drawing.Size(($script:PanelW + 24), ($y + 8))
    $work = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea
    if ($form.Height -gt $work.Height) { $form.Height = $work.Height }
    if ($form.Width  -gt $work.Width)  { $form.Width  = $work.Width }

    if ($filtering) {
        $lblFindCount.Text = "$shown of $total"
    } else {
        $lblFindCount.Text = ''
    }
}

function Update-SectionPanelWidth {
    <#
        Widen every panel if any of them actually grew a scrollbar.

        This ASKS rather than predicts, which is the whole point. Counting tools
        against declared capacity is a good guess and wrong in the one case that
        matters: capability-hidden tools are removed after the count would have
        been taken, so a predicted overflow can be an overflow that never
        happens -- and the console would carry a scrollbar's width of dead space
        for a scrollbar that was never drawn.

        Widening is GLOBAL. One panel a scrollbar's width narrower than its
        neighbours is the original layout bug: it wraps its last button and
        looks special-cased when it is only the fullest.

        NO OSCILLATION. The scrollbar exists because there are more ROWS than
        the section declared, and adding width does not remove a row: the extra
        is one scrollbar wide, always far less than the width of a button, so it
        can never admit another column and never reduce the row count. The
        widened layout therefore still needs the scrollbar, which is what makes
        one pass enough.

        Called twice on purpose -- see the call sites.
    #>
    if (@($script:SectionPanels.Values).Count -eq 0) { return }

    $need = $false
    foreach ($flow in $script:SectionPanels.Values) {
        if (-not $flow) { continue }
        try {
            # Force the handle and a layout pass. Scrollbar visibility is
            # decided during layout, and before the form is shown neither has
            # necessarily happened -- reading the property first would report
            # the state of a panel that has not been arranged yet.
            $null = $flow.Handle
            $flow.PerformLayout()

            # Two ways of asking, because AutoScroll reports through both and
            # they do not always agree before the form is visible.
            if ($flow.VerticalScroll.Visible) { $need = $true }
            elseif ($flow.DisplayRectangle.Height -gt $flow.ClientSize.Height) { $need = $true }
        } catch { }
    }

    $want = Get-ConsolePanelWidth -ButtonWidth      $script:BtnW `
                                  -ButtonPad        $script:BtnPad `
                                  -ButtonsPerRow    $script:PerRow `
                                  -ReserveScrollBar:$need
    if ($want -eq $script:PanelW) { return }

    $script:PanelW = $want
    $form.SuspendLayout()
    try {
        foreach ($flow in $script:SectionPanels.Values) {
            if ($flow) { $flow.Width = $script:PanelW }
        }
        # Everything else that was sized from PanelW has to follow, or the input
        # and output boxes stop lining up with the button grid -- which is the
        # same lopsided look this was written to remove, just moved.
        $txtInput.Width   = $script:PanelW
        $lblResult.Width  = $script:PanelW
        $txtOut.Width     = $script:PanelW
        $lblVersion.Width = $script:PanelW

        $form.ClientSize = New-Object System.Drawing.Size(
                                ($script:PanelW + 24), $form.ClientSize.Height)

        # Re-apply the screen clamp. Growing the form here can push it past the
        # working area on a small display, and the cap set at build time was
        # calculated against the narrower width.
        $work = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea
        if ($form.Width -gt $work.Width) { $form.Width = $work.Width }
    } finally {
        $form.ResumeLayout()
    }
}

#region NEW 2026-07-24 -- colour the Connect Scripts buttons by live state ----
# ORDER MATTERS: first match wins, and the on-premises pattern MUST be tested
# before the cloud one. 'Connect Exchange SE' contains the word 'exchange', so
# with ExO first the on-prem button would be painted with Exchange ONLINE's
# status -- green while the on-prem session was down, and no error anywhere.
$script:ConnectButtonServices = @(
    [PSCustomObject]@{ Service = 'EMS'; Pattern = 'exse|\bse\b|on-?prem|\bems\b' }
    [PSCustomObject]@{ Service = 'ExO'; Pattern = 'exo|exchange' }
    [PSCustomObject]@{ Service = 'SPO'; Pattern = 'spo|sharepoint' }
)

$script:ConnClrOk   = [System.Drawing.Color]::FromArgb(198, 239, 206)
$script:ConnClrBad  = [System.Drawing.Color]::FromArgb(255, 199, 206)
$script:ConnClrNone = $null   # $null = leave the themed colour alone
$script:ConnTextClr = [System.Drawing.Color]::FromArgb(20, 20, 20)

function Get-ConnectButtonService {
    <#
        Which service a CONNECT button is for, or $null if it is an ordinary
        tool.

        The name must contain the word 'connect' before any service pattern is
        considered. Without that gate the service patterns match far too much:
        'on-prem' alone caught 1-Onprem-Export and 2-Onprem-Decomm, which are
        ordinary DL Conversions tools -- and since this function now also
        decides which buttons SKIP capability greying, they would have quietly
        stopped greying out when the on-premises session was missing.

        'Connect' is the honest test anyway. This asks whether a button's job
        is to establish a connection, and that is what the word means.
    #>
    param($Tool)

    # Identified by SCRIPT FILE first, name second.
    #
    # The filename is a convention this toolkit controls -- every connect
    # script is Connect-<service>.ps1 -- whereas the button NAME is free text
    # the operator types into ToolManager. Keying only on the name made the
    # exemption depend on somebody happening to include the word 'connect',
    # which is a rule nobody was told and which fails silently: the button
    # greys out and nothing explains that its label was the problem.
    $leaf = ''
    try { $leaf = Split-Path -Leaf "$($Tool.Script)" } catch { }
    $name = "$($Tool.Name)"

    $isConnect = ($leaf -imatch '^connect[-_]') -or ($name -imatch '\bconnect\b')
    if (-not $isConnect) { return $null }

    # Service still comes from whichever of the two says something. Filename
    # first for the same reason.
    foreach ($svc in $script:ConnectButtonServices) {
        if ($leaf -imatch $svc.Pattern) { return $svc.Service }
    }
    foreach ($svc in $script:ConnectButtonServices) {
        if ($name -imatch $svc.Pattern) { return $svc.Service }
    }
    return $null
}

function Update-ConnectionButtonColors {
    param([switch]$SkipSpo)

    $prevCursor = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        Clear-ConnectedTenantDomainCache
        Update-TenantLabel

        # Read the capability map rather than probing again. These labels and the
        # tool grey-out have to agree: "ExO: Connected" beside a button greyed
        # for a missing Exchange session would be the console contradicting
        # itself. One probe, one answer, both consumers.
        $map = Get-CapabilityStateMap -Refresh -SkipSpo:$SkipSpo
        $exo = $map['ExO']
        $ems = $map['EMS']
        $spo = if ($map['SPO'].Unknown) { $null } else { $map['SPO'] }

        # Only the real Connect Scripts panel gets status colouring. The old
        # fallback to the first section painted Admin Tools buttons green/red if
        # a tool's NAME happened to match 'exchange' or 'sharepoint' -- harmless
        # with today's names, a mystery bug the first time one is renamed.
        $connPanel = $script:SectionPanels['Connect Scripts']
        foreach ($ctrl in @(if ($connPanel) { $connPanel.Controls } else { @() })) {
            if (-not $ctrl.Tag) { continue }
            $svc = Get-ConnectButtonService -Tool $ctrl.Tag
            if (-not $svc) {
                if ($script:ConnClrNone) { $ctrl.BackColor = $script:ConnClrNone }
                continue
            }
            $state = switch ($svc) { 'ExO' { $exo } 'EMS' { $ems } 'SPO' { $spo } }
            if (-not $state) { continue }

            $ctrl.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
            $ctrl.BackColor = if ($state.Available) { $script:ConnClrOk } else { $script:ConnClrBad }
            # Pale status colours need dark text. Without this the button keeps
            # the theme's white ForeColor and reads as disabled.
            $ctrl.ForeColor = $script:ConnTextClr
            $toolTip.SetToolTip($ctrl, "$svc -- $($state.Detail)")
        }

        # Contrast, not decoration. These sit on the theme's dark navy, where
        # #008000 green and #C80000 red are both dark-on-dark: legible if you
        # already know what they say, which is not what a status light is for.
        # Lifted to a mid-bright green and a warm red that hold up on dark
        # without glowing on the light themes.
        $green = [System.Drawing.Color]::FromArgb(80, 220, 130)
        $red   = [System.Drawing.Color]::FromArgb(255, 105, 97)

        # One label writer, so the four status labels cannot disagree about what
        # a state means. Three states, not two:
        #
        #   Connected      green
        #   Not Connected  red    -- present in this tenant, unreachable now
        #   n/a            grey   -- this tenant does not have it at all
        #
        # The third used to render as red "Not Connected", which in a cloud-only
        # tenant is a permanent alarm about a server that was never supposed to
        # exist. Red is for things worth fixing.
        $setState = {
            param($Lbl, $State, $Prefix)
            if (-not $State) {
                $Lbl.Text = "${Prefix}: not checked"
                $Lbl.ForeColor = [System.Drawing.Color]::DimGray
                return
            }
            if ($State.PSObject.Properties['NotConfigured'] -and $State.NotConfigured) {
                # Not greyed -- gone. See Update-StatusRowLayout.
                $Lbl.Visible = $false
                return
            }
            $Lbl.Visible = $true
            if ($State.Available) {
                $Lbl.Text = "${Prefix}: Connected"
                $Lbl.ForeColor = $green
            } else {
                $Lbl.Text = "${Prefix}: Not Connected"
                $Lbl.ForeColor = $red
            }
        }

        & $setState $lblExo $exo 'ExO'
        $toolTip.SetToolTip($lblExo, "$($exo.Detail)")

        if (-not $spo) {
            $lblSpo.Text = 'SPO: not checked'
            $lblSpo.ForeColor = [System.Drawing.Color]::DimGray
            $toolTip.SetToolTip($lblSpo, 'Not checked yet -- runs after any tool closes.')
        } else {
            & $setState $lblSpo $spo 'SPO'
            $toolTip.SetToolTip($lblSpo, "$($spo.Detail)")
        }
        # EMS and AD come from the same map, so these labels cannot disagree
        # with the grey-out (which is the whole reason the map exists).
        $ems = $map['EMS']
        $ad  = $map['AD']
        & $setState $lblEms $ems 'EMS'
        $toolTip.SetToolTip($lblEms, "$($ems.Detail)")

        & $setState $lblAd $ad 'AD'
        $toolTip.SetToolTip($lblAd, "$($ad.Detail)")

        # Only counts things this tenant actually has. Otherwise the hint stays
        # on for ever in a cloud-only tenant, telling you to connect to a
        # server that does not exist -- advice that cannot be followed.
        Update-StatusRowLayout -Row @($lblExo, $lblSpo, $lblTenant) -Y 8
        Update-StatusRowLayout -Row @($lblEms, $lblAd, $lblPim, $lblWarn) -Y 30

        $relevant = @($exo, $spo, $ems, $ad) | Where-Object {
            $_ -and -not ($_.PSObject.Properties['NotConfigured'] -and $_.NotConfigured)
        }
        $lblWarn.Visible = @($relevant | Where-Object { -not $_.Available }).Count -gt 0

        # Same probe, same answer: repaint the tool buttons from the map that
        # just produced these labels.
        Update-CapabilityState
    } catch {
        $lblExo.Text = 'ExO: unavailable'
        $lblExo.ForeColor = [System.Drawing.Color]::DimGray
        $lblSpo.Text = 'SPO: unavailable'
        $lblSpo.ForeColor = [System.Drawing.Color]::DimGray
        $lblEms.Text = 'EMS: unavailable'
        $lblEms.ForeColor = [System.Drawing.Color]::DimGray
        $lblAd.Text = 'AD: unavailable'
        $lblAd.ForeColor = [System.Drawing.Color]::DimGray
        # Probe failed -- show the prompt, since we cannot confirm either one.
        $lblWarn.Visible = $true
    } finally {
        $form.Cursor = $prevCursor
    }
}
#endregion

#region NEW 2026-07-29 -- grey out tools this window cannot run ---------------
<#
    Two independent reasons a button can be disabled, and they must be decided
    together in ONE place.

    The object-type grey-out below used to end with an unconditional
    $ctrl.Enabled = $true in its else branch. Adding capability greying as a
    separate second pass would mean whichever ran last won, and a Look Up would
    silently re-enable a button whose Exchange session does not exist. So both
    reasons are evaluated here and capability is reported first, because it is
    the more fundamental block: the wrong object type means "not for this
    target", a missing connection means "not at all".
#>
$script:ToolAvail = @{}

function Update-CapabilityState {
    <#
        Recomputes per-tool availability from the CURRENT capability map and
        repaints the buttons. Does not probe -- Update-ConnectionButtonColors
        owns the probing, so the labels and the grey-out can never be computed
        from two different answers.
    #>
    try {
        $map = Get-CapabilityStateMap
        $script:ToolAvail = @{}
        foreach ($tool in @($script:AllTools)) {
            $path = Join-Path $PSScriptRoot $tool.Script
            $script:ToolAvail["$($tool.Name)"] = Get-ToolAvailability -ScriptPath $path -StateMap $map
        }
    } catch {
        # A failed classification must not disable anything. An empty map means
        # every tool falls back to enabled -- the same fail-open rule that keeps
        # an unreadable script from being blocked.
        $script:ToolAvail = @{}
    }
    Update-ToolButtonState -DetectedType $script:LastType
}

function Get-ToolCapabilityBlock {
    param($Tool)

    # A tool whose job is to CREATE a connection must never be gated on that
    # connection. Connect-ExO.ps1 calls Connect-ExchangeOnline, so the
    # capability scan correctly reports it as needing Exchange Online -- and
    # acting on that would grey the button out at exactly the moment it is the
    # only useful control on the form. The greying is right for every other
    # tool and exactly inverted for these.
    #
    # Keyed on Get-ConnectButtonService rather than the section name, because
    # sections are user-editable in Sections.json and renaming one must not
    # quietly re-arm this.
    if (Get-ConnectButtonService -Tool $Tool) { return $null }

    # A NewProcess tool provisions its OWN session in its OWN process. This
    # console's probes describe this console's process, which is not the one
    # the tool will use -- so greying it here would block it for a condition
    # that has nothing to do with whether it can run. Same inverted logic as
    # the connect buttons, arrived at from the other direction.
    $tp = Join-Path $PSScriptRoot $Tool.Script
    if ((Test-Path -LiteralPath $tp) -and (Test-ToolDeclaresOwnProcess -ScriptPath $tp)) { return $null }

    $a = $script:ToolAvail["$($Tool.Name)"]
    if (-not $a) { return $null }        # never classified: fail open
    if ($a.Enabled) { return $null }
    return $a
}
#endregion

function Update-ToolButtonState {
    param([string]$DetectedType)
    $showDimming = ($DetectedType -and ($script:KnownObjectTypes -contains $DetectedType))
    $allCtrls = @()
    foreach ($n in $script:SectionNames) {
        $p = $script:SectionPanels[$n]
        if ($p) { $allCtrls += @($p.Controls) }
    }
    foreach ($ctrl in $allCtrls) {
        if (-not $ctrl.Tag) { continue }
        $tool = $ctrl.Tag

        # A group button is not a tool and must be handled before anything asks
        # it for a Name or ObjectTypes.
        #
        # Its state is the OR of its members: enabled while at least one thing
        # behind it can run. Disabling it whenever any member is blocked would
        # hide four working tools behind one that is not, and enabling it
        # unconditionally would send the operator through a click to reach a
        # dialog of dead buttons.
        if ($tool.PSObject.Properties['IsToolGroup']) {
            $members = @($script:ToolGroups["$($tool.GroupName)"])
            $usable  = 0
            foreach ($m in $members) {
                if (Get-ToolCapabilityBlock -Tool $m) { continue }
                $md = @($m.ObjectTypes | Where-Object { "$_".Trim() })
                if ($showDimming -and $md.Count -gt 0 -and ($md -notcontains $DetectedType)) { continue }
                $usable++
            }
            $ctrl.Enabled = ($usable -gt 0)
            if ($usable -eq 0) {
                $toolTip.SetToolTip($ctrl,
                    "None of the $($members.Count) tools in '$($tool.GroupName)' can run right now. Open it to see why each one is blocked.")
            } else {
                $toolTip.SetToolTip($ctrl,
                    "$($members.Count) tools. $usable available for the current lookup.")
            }
            continue
        }

        $capBlock = Get-ToolCapabilityBlock -Tool $tool

        # An EMPTY ObjectTypes means "applies to anything", not "applies to
        # nothing".
        #
        # The old expression read @() as a list that contains no match, so any
        # tool with no declared types went grey the moment a lookup succeeded --
        # Mailbox / OneDrive Perms, ResourceDuplicator and MB w No activity were
        # permanently disabled after any Look Up, with a tooltip claiming they
        # "apply to: " and then nothing.
        #
        # Declaring no types is how you say a tool is not object-scoped, which
        # is the opposite of what it was being read as. Fail open, same rule the
        # capability classifier follows: an unclassified tool stays enabled.
        $declared = @($tool.ObjectTypes | Where-Object { "$_".Trim() })
        $typeOk   = $true
        if ($showDimming -and $declared.Count -gt 0) {
            $typeOk = ($declared -contains $DetectedType)
        }

        if ($capBlock) {
            $ctrl.Enabled = $false
            $tip = "'$($tool.Name)' cannot run in this window." + "`r`n`r`n" + $capBlock.Reason
            if (-not $typeOk) {
                $tip += "`r`n`r`nIt also applies only to: $(@($tool.ObjectTypes) -join ', ') -- the last lookup resolved to '$DetectedType'."
            }
            $toolTip.SetToolTip($ctrl, $tip)
        } elseif (-not $typeOk) {
            $ctrl.Enabled = $false
            $toolTip.SetToolTip($ctrl, "'$($tool.Name)' applies to: $(@($tool.ObjectTypes) -join ', '). The last lookup resolved to '$DetectedType', so this is disabled for it. Look up a matching object to enable it -- or clear the input box and re-run Look Up with nothing entered to release the grey-out on every tool.")
        } else {
            $ctrl.Enabled = $true
            # Asks the script, like everything else now. This branch used to
            # read the checkbox, so a tool that DECLARED its own process got the
            # generic tooltip instead of this one -- the message describing the
            # behaviour was shown only to tools that were not getting it.
            $tp = Join-Path $PSScriptRoot $tool.Script
            if ((Test-Path -LiteralPath $tp) -and (Test-ToolDeclaresOwnProcess -ScriptPath $tp)) {
                $toolTip.SetToolTip($ctrl,
                    "Opens in its own PowerShell window and connects to on-premises Exchange itself." + "`r`n" +
                    "This console's connections do not apply to it, so it is never greyed out." + "`r`n" +
                    "Expect a delay while that window imports the Exchange session.")
            } else {
                $a = $script:ToolAvail["$($tool.Name)"]
                $toolTip.SetToolTip($ctrl, $(if ($a) { $a.Reason } else { '' }))
            }
        }
    }
}

$btnRun.Add_Click({
    $btnRun.Enabled = $false
    $txtOut.Text = ''
    $lblResult.Text = 'Looking up...'
    try {
        $smtp = Get-TargetSmtp -Text $txtInput.Text
        if (-not $smtp) {
            $script:LastSmtp = $null
            $script:LastADServer = $null
            $script:LastEnvironment = $null
            $script:LastType = $null
            $lblResult.Text = 'Enter an email address first. (Tools grey-out has been cleared.)'
            Update-ToolButtonState -DetectedType $null
            return
        }
        $lblResult.Text = 'Looking up...'

        # THE DIRECTORY LOOKUP IS OPTIONAL. THE OBJECT IS NOT.
        #
        # This required an on-premises hit before it would look at anything, so
        # on a cloud-only tenant every address came back "not found" and every
        # tool stayed greyed -- for a mailbox that exists and that Exchange
        # Online can see perfectly well. The address was never the problem; the
        # question was.
        #
        # Resolve-Target already tries Exchange Online FIRST and only falls back
        # to AD, so it needs no directory at all to find a cloud mailbox. All
        # that was missing was letting it run.
        $hasAd = $true
        try { $hasAd = Test-TenantCapability -Name 'OnPremAD' } catch { $hasAd = $true }

        $envName = $null
        $adServer = $null
        if ($hasAd) {
            $envResult = Find-TargetEnvironment -Smtp $smtp
            if (-not $envResult.Environment) {
                $script:LastSmtp = $null
                $script:LastADServer = $null
                $script:LastEnvironment = $null
                $script:LastType = $null
                $lblResult.Text = "Not found in the on-premises directory -- $smtp"
                if ($envResult.Errors.Count -gt 0) {
                    $txtOut.Text = ($envResult.Errors -join "`r`n`r`n")
                }
                Update-ToolButtonState -DetectedType $null
                return
            }
            $envName  = $envResult.Environment
            $adServer = $envResult.ADServer
        }

        $target = Resolve-Target -Smtp $smtp -ADServer $adServer
        if ($target.Type -eq 'NotFound' -and -not $hasAd) {
            $script:LastSmtp = $null
            $script:LastADServer = $null
            $script:LastEnvironment = $null
            $script:LastType = $null
            $lblResult.Text = "Not found in Exchange Online -- $smtp"
            if ($target.ResolutionErrors.Count -gt 0) {
                $txtOut.Text = ($target.ResolutionErrors -join "`r`n`r`n")
            }
            Update-ToolButtonState -DetectedType $null
            return
        }
        $script:LastSmtp = $smtp
        $script:LastADServer = $adServer
        $script:LastEnvironment = $envName
        $script:LastType = $target.Type

        # Name the directory only when there IS one. On a cloud-only tenant
        # "Found: Mailbox () -- name" reads as a bug in the console rather
        # than as an absent directory.
        $where = if ($envName) { " ($envName)" } else { '' }

        switch ($target.Type) {
            'Mailbox'      { $lblResult.Text = "Found: Mailbox$where -- $smtp" }
            'UnifiedGroup' { $lblResult.Text = "Found: Microsoft 365 Group$where -- $smtp" }
            'OnpremGroup'  { $lblResult.Text = "Found: On-Prem Group$where -- $smtp" }
            default {
                $lblResult.Text = "Object found$where but not confirmed as a Mailbox, Microsoft 365 Group, or on-prem AD group -- $smtp"
                if ($target.ResolutionErrors.Count -gt 0) {
                    $txtOut.Text = ($target.ResolutionErrors -join "`r`n`r`n")
                }
            }
        }
        Update-ToolButtonState -DetectedType $target.Type
    } catch {
        $lblResult.Text = 'Lookup failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Admin Console lookup'
        Update-ToolButtonState -DetectedType $null
    } finally {
        $btnRun.Enabled = $true
    }
})
$btnClear.Add_Click({
    $txtInput.Text = ''
    $lblResult.Text = 'Not looked up yet.'
    $txtOut.Text = ''
    $script:LastSmtp = $null
    $script:LastADServer = $null
    $script:LastEnvironment = $null
    $script:LastType = $null
    Update-ToolButtonState -DetectedType $null
})
$btnClose.Add_Click({ $form.Close() })

# Filter as you type. SuspendLayout around it because the handler moves every
# panel and the version label on each keystroke, and WinForms would otherwise
# repaint the whole form per character.
$txtFind.Add_TextChanged({
    $form.SuspendLayout()
    try { Update-ConsoleSearch -Text $txtFind.Text }
    finally { $form.ResumeLayout() }
})

# Escape clears the filter rather than closing the console.
#
# $form.CancelButton is Close, so without this Escape would shut the window --
# which is exactly what a person does when they want to abandon a search, and
# losing the console for it would be infuriating. Only swallowed when there is
# something to clear, so Escape still closes an unfiltered console.
$txtFind.Add_KeyDown({
    if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape -and $txtFind.Text) {
        $_.Handled = $true
        $_.SuppressKeyPress = $true
        $txtFind.Text = ''
    }
})

$btnFindClear.Add_Click({
    $txtFind.Text = ''
    $txtFind.Focus()
})

Build-ToolButtons
Update-ToolButtonState -DetectedType $null

$script:Theme = Get-ConsoleTheme
if ($script:Theme) {
    Set-ConsoleTheme -Form $form -Theme $script:Theme
    # lblWarn is NoTheme-tagged, so it is coloured explicitly here: the theme's
    # own text colour on a themed console, grey on the default one.
    $wc = ConvertTo-ThemeColor $script:Theme.FormText
    if ($wc) { $lblWarn.ForeColor = $wc }
}
# Margin 10, not 6. At 6 the logo sat hard against the window edge and looked
# clipped; the corner now has room to breathe and the '?' tucks in beside it.
Add-FormLogo -Form $form -Corner 'TopRight' -MaxHeight 64 -Margin 10 | Out-Null

Update-ConnectionButtonColors

# Panel width is settled HERE, not after Build-ToolButtons.
#
# Update-ConnectionButtonColors ends in Update-CapabilityState, which is what
# hides tools this tenant does not have. Measuring before that runs would ask
# the panels about a button set that is about to change -- and the answer would
# be a scrollbar's width too wide in exactly the tenants that hide the most.
Update-SectionPanelWidth

# Second pass once the window is really on screen.
#
# Scrollbar visibility is decided during layout, and a forced PerformLayout on
# an unshown form is close to the truth but not guaranteed to BE it -- an
# off-by-one in the scrollbar's own width, or a theme that changed a control's
# size after the first pass, and the pre-show answer is stale. This one costs
# nothing when the answer is unchanged: the function returns early when the
# width it computes matches the width already applied.
# GUARDED, because an exception in a Shown handler kills the window.
#
# Shown runs inside the modal loop. A terminating error here does not print and
# does not stop the script -- WinForms tears the form down and ShowDialog
# returns as though the operator had closed it. The console appears for a
# fraction of a second and vanishes, with nothing written anywhere.
#
# That matches a rare, unreproducible "it closed immediately, second launch was
# fine" (2026-08-13) -- unproven as the cause, but this is measurement code
# running at first paint, so it depends on scrollbar state, DPI and theme
# timing, which is exactly the shape of a fault that only sometimes happens.
#
# The catch does not hide anything: the width stays at its pre-show value,
# which is the value that was already good enough to draw with, and the reason
# lands in $txtOut where it can be read and sent on.
$form.Add_Shown({
    Write-StartupBreadcrumb -Text 'form Shown -- laying out panels'
    try {
        Update-SectionPanelWidth
        Write-StartupBreadcrumb -Text 'panels laid out'
    }
    catch {
        # BOTH, and the file first. The output box is nicer to read but only
        # exists while the window does, and a window that dies takes the
        # message with it -- which is exactly how the first version of this
        # guard turned out to be worthless.
        Write-StartupFailure -Stage 'Update-SectionPanelWidth (form Shown)' -ErrorRecord $_
        try {
            $txtOut.Text = ("The console opened, but laying out the button panels failed." +
                [Environment]::NewLine + [Environment]::NewLine +
                "$($_.Exception.Message)" + [Environment]::NewLine + [Environment]::NewLine +
                "Everything still works -- the panels may be slightly the wrong width. " +
                "This is also in Logs\ConsoleStartup.log.")
        } catch { }
    }
})

# One write, after every button has been classified, rather than one per tool.
# The scan itself is what used to make a cold launch slow; persisting it means
# only the FIRST launch on this machine pays, and only again for scripts that
# have actually changed since. Failing to save costs a rescan, never an error.
Save-ScriptCapabilityCache

# Surfaced here, not at migration time -- the move happens before $txtOut
# exists, and a file move the operator never sees is the kind of thing that
# gets blamed for something else three weeks later.
if ($script:ConfigMigrationNotes -and $script:ConfigMigrationNotes.Count -gt 0) {
    # Guarded for the same reason as the handler above: a Shown handler that
    # throws takes the window with it, and this one runs only on the FIRST
    # launch after a config migration -- the launch least likely to be
    # reproducible and most likely to be blamed on something else.
    $form.Add_Shown({
        try {
            $txtOut.Text = ("Config folder created -- existing settings moved in:" +
                [Environment]::NewLine +
                (($script:ConfigMigrationNotes | ForEach-Object { "  $_" }) -join [Environment]::NewLine) +
                [Environment]::NewLine +
                "Update rule from here: replace the .ps1 files, leave Config\ alone.")
        } catch { }
    })
}

Set-FormScreenCorner -Form $form -Corner 'TopLeft'

$script:PendingConnect = $null
# The console's own '?', in the same corner as every tool's. x = ClientSize-40
# lands at 784..812, which is clear -- lblTenant ends at 660.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Admin Console' | Out-Null

Update-PimLabel -Lbl $lblPim -Tip $toolTip

# One tick a minute, doing arithmetic on two timestamps and setting a label.
# Deliberately not shorter, and deliberately no DoEvents anywhere near it --
# see 14c on why extra message-loop activity in this console is treated with
# suspicion.
#
# THE CONTROLS RIDE ON THE TIMER'S .Tag, and the handler must not name them.
#
# A tick scriptblock is not a closure: PowerShell resolves its variables
# against whatever scope is current WHEN THE TICK FIRES. The console launches
# tools with '& $scriptPath', so while a tool's ShowDialog() pumps messages on
# this thread the current scope is the TOOL's -- and a tool that happens to use
# the name $lblPim shadows the console's label for the duration.
#
# That is not hypothetical: it is why this console threw
# "The property 'Visible' cannot be found on this object" once a minute while
# Dashboard was open in a tenant without PIM. See Update-PimLabel.
#
# Tag is the same fix Add-ReportFileControls and Add-ToolHelpButton use, and
# for the same underlying reason: the scope a scriptblock RUNS in is not the
# scope it was WRITTEN in. GetNewClosure is not the answer here -- it would
# bind this scriptblock to a dynamic module and Update-PimLabel, which is
# dot-sourced into script scope, would stop resolving.
$pimTimer = New-Object System.Windows.Forms.Timer
$pimTimer.Interval = 60000
$pimTimer.Tag      = [PSCustomObject]@{ Lbl = $lblPim; Tip = $toolTip }
$pimTimer.Add_Tick({
    $c = $this.Tag
    if (-not $c) { return }
    Update-PimLabel -Lbl $c.Lbl -Tip $c.Tip
})
$pimTimer.Start()
$form.Add_FormClosed({ $pimTimer.Stop(); $pimTimer.Dispose() })

# The two breadcrumbs that localise the fault.
#
# "about to ShowDialog" with nothing after it = the window never opened, and
# everything above this line is suspect. "ShowDialog returned" a fraction of a
# second later = the window DID open and something inside the message loop tore
# it down. Those are different investigations, and without these two lines
# there is no way to tell them apart after the fact.
Write-StartupBreadcrumb -Text 'about to ShowDialog'
[void]$form.ShowDialog()
Write-StartupBreadcrumb -Text 'ShowDialog returned (window closed)'

# Past this line the modal loop is gone, so a sign-in window can take input.
if ($script:PendingConnect) {
    $pending = $script:PendingConnect
    $script:PendingConnect = $null

    $path = Join-Path $PSScriptRoot $pending.Script
    if (Test-Path -LiteralPath $path) {
        Write-Host ''
        Write-Host "Closing the console to sign in -- it will reopen when you are done." -ForegroundColor Cyan

        # ShowDialog installs a WindowsFormsSynchronizationContext on this
        # thread and leaves it there after the form is gone. An interactive
        # sign-in draws its own window, and MSAL posts that window onto whatever
        # context it finds -- so it lands on a pump that closed with the form
        # and fails with "the current thread is not in a single-threaded
        # apartment".
        #
        # Cleared HERE as well as inside the connect script. The script defends
        # itself because it can be run directly; this line means the console
        # never hands it the problem in the first place. Nothing below needs the
        # context -- the console is rebuilt from scratch when it reopens.
        [System.Threading.SynchronizationContext]::SetSynchronizationContext($null)

        try {
            & $path
        } catch {
            Write-Host (Format-ErrorDetail -ErrorRecord $_ -Context "Running $($pending.Name)") -ForegroundColor Red
        }
    } else {
        Write-Host "'$($pending.Name)' points to '$($pending.Script)', which does not exist." -ForegroundColor Red
    }

    # ASK TO BE RE-OPENED. Do not re-open here.
    #
    # This used to be '& $PSCommandPath' -- the script calling itself, which
    # left every connect one stack frame deeper and needed a depth guard that
    # eventually told the operator to relaunch by hand. The trampoline at the
    # top of this file owns the loop now, so setting a flag and falling off the
    # end of the script is the whole mechanism.
    #
    # Falling through rather than returning MATTERS: the session-marker cleanup
    # below still has to run on the way out. The trampoline sets it again on
    # the next pass, and nothing executes in the gap between.
    $global:ToolkitConsoleRelaunch = $true
}

# THE MARKER DIES WITH THE CONSOLE.
#
# Set at the top, cleared here, and this line is the whole reason the guard is
# worth anything on a shared machine.
#
# The console runs INSIDE the shell that started it -- there is no separate
# console process -- so an environment variable it sets survives it closing.
# Without this line, launching the console once left that window permanently
# marked: every script typed at that prompt afterwards passed the guard, for as
# long as the window stayed open. Confirmed live on 2026-08-13, which is how it
# was found.
#
# CLEARED HERE, NOT IMMEDIATELY AFTER ShowDialog, because the connect path runs
# in between: the console closes, a guarded Connect-*.ps1 runs, and then the
# console relaunches itself. Clearing any earlier blocks the connect script and
# breaks sign-in.
#
# The relaunch loop is safe. Each pass sets the marker on entry and clears it
# here on the way out, and the trampoline starts the next pass immediately --
# so the marker is only empty in the gap between two passes, where no script
# runs. The last pass leaves it empty, which is the state that matters.
#
# HONEST LIMIT: if the console dies on an unhandled error this line never runs
# and the marker persists until that window closes. That is a speed bump losing
# its bump in a rare case, not a hole in something load-bearing -- see
# Test-ConsoleLaunch on what this is and is not.
$env:ToolkitConsoleSession = $null
Write-StartupBreadcrumb -Text 'console session marker cleared'
#endregion
```


<br>

==============================================================================

# ◆ FILE 4 OF 16 — `ToolManager.ps1`

> **948 lines.** Config round-trip. Note how it carries through fields it does not itself edit -- silently dropping them is a documented past bug that once deleted a live setting on first save.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```powershell
#Requires -Version 5.1
<#
    ToolManager.ps1
    Editor for Tools.json -- add, remove, rename, reorder, and repoint the
    tools AdminConsole.ps1 displays, with no code changes required.

    Author   : Glenn Armbruster
    Revised  : 2026-08-10
#>
# TOOLKIT-HELP-BEGIN
# Decides which buttons the console shows, and what each one runs.
#
# Every row is one button. Edit cells directly, then Save. AdminConsole.ps1
# re-reads this file each time it starts, so changes appear on the next launch.
#
# THE COLUMNS
#   Name         The button caption.
#   Script       The .ps1 it runs, relative to the toolkit folder.
#   ObjectTypes  Which looked-up object the tool applies to -- Mailbox,
#                UnifiedGroup, OnpremGroup, comma-separated. LEAVE BLANK for a
#                tool that is not about one object; blank means "always
#                applies", not "never".
#   Section      Which panel it appears in. Sections... defines them.
#   Group        Optional. Type the same word on several tools and they collapse
#                behind one button that opens a chooser -- 'Rooms' on the room
#                tools, for instance. Blank means the tool stands alone.
#   Tier         Free text for your own notes. The console ignores it.
#   Own window   Runs the tool in its own PowerShell process. REQUIRED for
#                anything touching on-premises Exchange -- see below.
#
# OWN WINDOW IS NOT A PREFERENCE. On-premises Exchange and Exchange Online
# export the same cmdlet names, and in one process the last connection wins. A
# tool without it can act on the wrong directory with no error at all. If in
# doubt, tick it.
#
# Order is row position -- use Move Up / Move Down, not the Order column.
#
# The other buttons: Sections... edits the panels and how many buttons wide
# they are. Theme... picks colours. Settings... holds YOUR paths and email,
# which are per-engineer and not shipped.
# TOOLKIT-HELP-END

param(
    [string]$ToolsConfigPath = ''
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

# Config now lives in Config\. Resolved after Common.ps1 loads, because the
# resolver is defined there -- a param default cannot call it.
if (-not $ToolsConfigPath) { $ToolsConfigPath = Get-ConfigFilePath -Name 'Tools.json' }

$validObjectTypes = @('Mailbox', 'UnifiedGroup', 'OnpremGroup')

# Sections are user-defined now and live in Sections.json alongside Tools.json.
# Each section is ButtonsPerRow wide and as many rows as it declares, so its
# capacity is simply rows x width -- nothing here is hardcoded to a name.
$SectionsConfigPath = Get-ConfigFilePath -Name 'Sections.json'

function Read-SectionsJson {
    $default = [PSCustomObject]@{
        ButtonsPerRow = 5
        Sections = @(
            [PSCustomObject]@{ Name = 'Connect Scripts'; Rows = 1 }
            [PSCustomObject]@{ Name = 'Admin Tools';     Rows = 4 }
            [PSCustomObject]@{ Name = 'Daily Scripts';   Rows = 1 }
        )
    }
    if (-not (Test-Path -LiteralPath $SectionsConfigPath)) { return $default }
    try {
        $cfg = Get-Content -LiteralPath $SectionsConfigPath -Raw | ConvertFrom-Json
        if (-not $cfg.Sections -or @($cfg.Sections).Count -eq 0) { return $default }
        if (-not $cfg.ButtonsPerRow) {
            Add-Member -InputObject $cfg -NotePropertyName ButtonsPerRow -NotePropertyValue 5 -Force
        }
        return $cfg
    } catch { return $default }
}

function Write-SectionsJson {
    param($Config)
    $Config | ConvertTo-Json -Depth 6 |
        Set-Content -LiteralPath $SectionsConfigPath -Encoding UTF8
}

$script:SectionCfg = Read-SectionsJson

function Get-ValidSections {
    return @(@($script:SectionCfg.Sections) | ForEach-Object { "$($_.Name)" })
}

function Get-SectionCapacityTable {
    $w = [int]$script:SectionCfg.ButtonsPerRow
    if ($w -lt 1) { $w = 5 }
    $h = @{}
    foreach ($s in @($script:SectionCfg.Sections)) {
        $r = [int]$s.Rows; if ($r -lt 1) { $r = 1 }
        $h["$($s.Name)"] = $r * $w
    }
    return $h
}

$validSections   = Get-ValidSections
$script:sectionCapacity = Get-SectionCapacityTable

function Get-NormalizedSection {
    param([string]$Value)
    $s = "$Value".Trim()
    foreach ($n in (Get-ValidSections)) { if ($s -eq $n) { return $n } }
    return (Get-ValidSections)[0]
}

function Show-SectionEditor {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Sections'
    $dlg.Size = New-Object System.Drawing.Size(460, 450)
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.StartPosition = 'CenterParent'
    $dlg.Font = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $lbl = New-Object System.Windows.Forms.Label
    # The button count used to be baked into this sentence. It was written once
    # at dialog creation and never updated, so the moment the spinner below was
    # touched there were two different numbers on screen claiming to be the
    # same setting. The live one is in the preview at the bottom now.
    $lbl.Text = "Sections appear in the console top to bottom, in this order. Every section is the same number of buttons across."
    $lbl.Location = New-Object System.Drawing.Point(12, 10)
    $lbl.Size = New-Object System.Drawing.Size(420, 34)
    $dlg.Controls.Add($lbl)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(12, 50)
    $grid.Size = New-Object System.Drawing.Size(320, 250)
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.AutoSizeColumnsMode = 'Fill'

    $cN = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $cN.Name = 'Name'; $cN.HeaderText = 'Section name'; $cN.FillWeight = 70
    [void]$grid.Columns.Add($cN)
    $cR = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $cR.Name = 'Rows'; $cR.HeaderText = 'Rows'; $cR.FillWeight = 30
    [void]$grid.Columns.Add($cR)
    $dlg.Controls.Add($grid)

    foreach ($s in @($script:SectionCfg.Sections)) {
        [void]$grid.Rows.Add("$($s.Name)", "$([int]$s.Rows)")
    }

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = 'Add'; $btnAdd.Location = New-Object System.Drawing.Point(342, 50)
    $btnAdd.Size = New-Object System.Drawing.Size(90, 26)
    $btnAdd.Add_Click({ [void]$grid.Rows.Add('New Section', '1') })
    $dlg.Controls.Add($btnAdd)

    $btnDel = New-Object System.Windows.Forms.Button
    $btnDel.Text = 'Remove'; $btnDel.Location = New-Object System.Drawing.Point(342, 82)
    $btnDel.Size = New-Object System.Drawing.Size(90, 26)
    $btnDel.Add_Click({
        if ($grid.Rows.Count -le 1) {
            [void][System.Windows.Forms.MessageBox]::Show(
                'At least one section is required.', 'Sections', 'OK', 'Warning'); return
        }
        if ($grid.SelectedRows.Count -gt 0) { $grid.Rows.Remove($grid.SelectedRows[0]) }
    })
    $dlg.Controls.Add($btnDel)

    $btnUp = New-Object System.Windows.Forms.Button
    $btnUp.Text = 'Move Up'; $btnUp.Location = New-Object System.Drawing.Point(342, 120)
    $btnUp.Size = New-Object System.Drawing.Size(90, 26)
    $btnUp.Add_Click({
        if ($grid.SelectedRows.Count -eq 0) { return }
        $r = $grid.SelectedRows[0]; $i = $r.Index
        if ($i -le 0) { return }
        $grid.Rows.Remove($r); $grid.Rows.Insert($i - 1, $r); $r.Selected = $true
    })
    $dlg.Controls.Add($btnUp)

    $btnDown = New-Object System.Windows.Forms.Button
    $btnDown.Text = 'Move Down'; $btnDown.Location = New-Object System.Drawing.Point(342, 152)
    $btnDown.Size = New-Object System.Drawing.Size(90, 26)
    $btnDown.Add_Click({
        if ($grid.SelectedRows.Count -eq 0) { return }
        $r = $grid.SelectedRows[0]; $i = $r.Index
        if ($i -ge $grid.Rows.Count - 1) { return }
        $grid.Rows.Remove($r); $grid.Rows.Insert($i + 1, $r); $r.Selected = $true
    })
    $dlg.Controls.Add($btnDown)

    $lblW = New-Object System.Windows.Forms.Label
    $lblW.Text = 'Buttons per row:'
    $lblW.Location = New-Object System.Drawing.Point(12, 310)
    $lblW.Size = New-Object System.Drawing.Size(100, 20)
    $dlg.Controls.Add($lblW)

    $numW = New-Object System.Windows.Forms.NumericUpDown
    $numW.Location = New-Object System.Drawing.Point(116, 308)
    $numW.Size = New-Object System.Drawing.Size(50, 22)
    $numW.Minimum = 1; $numW.Maximum = 8
    $numW.Value = [int]$script:SectionCfg.ButtonsPerRow
    $dlg.Controls.Add($numW)

    # Live size preview.
    #
    # Buttons per row used to be a number you changed, saved, relaunched the
    # console, and then judged -- and if it was wrong you did the whole loop
    # again. The console sizes its buttons by measuring the captions now, so
    # the resulting width is knowable HERE, before saving.
    #
    # It is measured with $dlg.Font, which is the same Microsoft Sans Serif 9
    # the console uses. If those two ever diverge this preview becomes fiction,
    # so they are both left alone deliberately.
    $lblPreview = New-Object System.Windows.Forms.Label
    $lblPreview.Location = New-Object System.Drawing.Point(12, 342)
    $lblPreview.Size = New-Object System.Drawing.Size(420, 48)
    $dlg.Controls.Add($lblPreview)

    # Read straight from Tools.json rather than from whatever the grid behind
    # this dialog holds. The console measures the FILE, so previewing anything
    # else would predict a width the console is not going to use.
    $script:PreviewTools = @()
    try {
        $script:PreviewTools = @(Get-Content -LiteralPath $ToolsConfigPath -Raw | ConvertFrom-Json)
    } catch {
        $script:PreviewTools = @()
    }

    $updatePreview = {
        try {
            $per  = [int]$numW.Value
            $caps = @(@($script:PreviewTools) | ForEach-Object { "$($_.Name)" } |
                        Where-Object { "$_".Trim() })
            if ($caps.Count -eq 0) {
                $lblPreview.Text = 'Tools.json could not be read, so the size cannot be previewed.'
                return
            }

            # Sections come from the GRID, not from the saved file. The preview
            # has to describe what Save is about to write -- previewing what is
            # already on disk would mean the number stops moving the moment you
            # change a row count, which reads as the preview being broken.
            $secs = @()
            foreach ($row in $grid.Rows) {
                $n = "$($row.Cells['Name'].Value)".Trim()
                if (-not $n) { continue }
                $r = 1
                [int]::TryParse("$($row.Cells['Rows'].Value)", [ref]$r) | Out-Null
                if ($r -lt 1) { $r = 1 }
                $secs += [PSCustomObject]@{ Name = $n; Rows = $r }
            }

            # An ESTIMATE of the scrollbar case, and knowingly a pessimistic
            # one. The console decides by asking the panels after tenant
            # capabilities have hidden what they are going to hide; this counts
            # every tool in the file, so it can predict an overflow the console
            # will not actually have. Wrong in the safe direction: it warns
            # about a wider window than you may get, never a narrower one.
            $over = Test-ConsoleSectionOverflow -Sections      $secs `
                                                -Tools         $script:PreviewTools `
                                                -ButtonsPerRow $per

            $bw = Get-ConsoleButtonWidth -Captions     $caps `
                                         -Font         $dlg.Font `
                                         -ButtonHeight $script:ConsoleBtnHeight
            $pw = Get-ConsolePanelWidth -ButtonWidth      $bw `
                                        -ButtonPad        $script:ConsoleBtnPad `
                                        -ButtonsPerRow    $per `
                                        -ReserveScrollBar:$over
            $formW   = $pw + 24
            $screenW = [System.Windows.Forms.Screen]::FromControl($dlg).WorkingArea.Width

            $t = "$per across, buttons $bw px wide. Console window: $formW px." +
                 " This screen fits $screenW px."
            if ($over) {
                $t += "  A section holds more tools than its rows allow, so every" +
                      " panel carries scrollbar width."
            }
            if ($formW -gt $screenW) {
                $t += "  Too wide for this screen -- the console will scroll sideways."
            }
            $lblPreview.Text = $t
        } catch {
            $lblPreview.Text = ''
        }
    }
    $numW.Add_ValueChanged($updatePreview)
    $grid.Add_CellValueChanged($updatePreview)
    & $updatePreview

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Save'; $btnOk.Location = New-Object System.Drawing.Point(232, 306)
    $btnOk.Size = New-Object System.Drawing.Size(90, 28)
    $dlg.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'; $btnCancel.Location = New-Object System.Drawing.Point(330, 306)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 28)
    $btnCancel.Add_Click({ $dlg.Close() })
    $dlg.Controls.Add($btnCancel)

    $btnOk.Add_Click({
        $list = @()
        $seen = @{}
        foreach ($row in $grid.Rows) {
            $n = "$($row.Cells['Name'].Value)".Trim()
            if (-not $n) { continue }
            if ($seen.ContainsKey($n)) {
                [void][System.Windows.Forms.MessageBox]::Show(
                    "Duplicate section name: '$n'. Names must be unique.",
                    'Sections', 'OK', 'Warning'); return
            }
            $seen[$n] = $true
            $r = 1
            [int]::TryParse("$($row.Cells['Rows'].Value)", [ref]$r) | Out-Null
            if ($r -lt 1) { $r = 1 }
            $list += [PSCustomObject]@{ Name = $n; Rows = $r }
        }
        if ($list.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show(
                'At least one section is required.', 'Sections', 'OK', 'Warning'); return
        }
        # ButtonWidth is carried through rather than rebuilt.
        #
        # It is not editable in this dialog, so a save that dropped it would
        # wipe a deliberate override the next time anyone merely reordered a
        # section -- and the symptom would be "the console resized itself for
        # no reason", which is not a clue that leads anywhere.
        $new = [ordered]@{ ButtonsPerRow = [int]$numW.Value }
        if ($script:SectionCfg.PSObject.Properties['ButtonWidth']) {
            $new['ButtonWidth'] = $script:SectionCfg.ButtonWidth
        }
        $new['Sections'] = $list
        $script:SectionCfg = [PSCustomObject]$new
        try {
            Write-SectionsJson -Config $script:SectionCfg
            $dlg.Tag = 'saved'
            $dlg.Close()
        } catch {
            [void][System.Windows.Forms.MessageBox]::Show(
                "Could not save Sections.json:`n$($_.Exception.Message)",
                'Sections', 'OK', 'Error')
        }
    })

    # This dialog was never themed -- it opened in default Windows colours while
    # its parent was dark. It is safe to theme now that DataGridView cell styles
    # are handled properly; before that fix it would have inherited the same
    # dark-text-on-dark-background problem the main grid had.
    foreach ($col in @($grid.Columns)) {
        $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
    }
    $grid.RowTemplate.Height = 26
    Set-ConsoleTheme -Form $dlg

    [void]$dlg.ShowDialog()
    return ($dlg.Tag -eq 'saved')
}

function Read-ToolsJson {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    if ("$raw".Trim() -eq '') { return @() }
    $parsed = $raw | ConvertFrom-Json -ErrorAction Stop
    return @($parsed)
}

function Write-ToolsJson {
    param([string]$Path, $ToolObjects)
    $tools = @($ToolObjects)
    $json = $tools | ConvertTo-Json -Depth 4
    if ($tools.Count -le 1) { $json = "[`r`n$json`r`n]" }
    Set-Content -LiteralPath $Path -Value $json -Encoding UTF8
}

#region UI ---------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Tool Manager -- $ToolsConfigPath"
# 1180 wide, up from 980.
#
# The grid gained a Group column and the status line reports a full path --
# 'E:\scripts\m365\adminconsole\Config\Tools.json' is 44 characters before the
# count in front of it. At 980 the columns were squeezed to ellipses and the
# path ran out of label.
$form.Size = New-Object System.Drawing.Size(1180, 820)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

# Resizable, because 32 tools in a fixed 380px viewport is the actual problem --
# no amount of contrast fixes having to scroll a list you are trying to compare
# against itself. Anchors below keep the grid growing and the buttons pinned.
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
$form.MinimumSize     = New-Object System.Drawing.Size(900, 600)
$form.MaximizeBox     = $true

$lblNote = New-Object System.Windows.Forms.Label
$lblNote.Text = "Use Add Tool / Delete Selected to add or remove rows, edit cells directly to rename or repoint a tool, then click Save. AdminConsole.ps1 reads this file every time it starts. Order (leftmost column) is read-only -- it reflects row position; use Move Up / Move Down to change it."
$lblNote.Location = New-Object System.Drawing.Point(12, 10)
$lblNote.Size = New-Object System.Drawing.Size(860, 34)
$form.Controls.Add($lblNote)

$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Location = New-Object System.Drawing.Point(12, 48)
$dgv.Size = New-Object System.Drawing.Size(860, 380)
$dgv.AllowUserToAddRows = $true
$dgv.AllowUserToDeleteRows = $true
$dgv.AutoGenerateColumns = $false
$dgv.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dgv.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dgv.Font = New-Object System.Drawing.Font("Consolas", 10)
$dgv.RowTemplate.Height = 26
$dgv.AllowUserToResizeRows = $false
$dgv.ColumnHeadersHeight = 28
$dgv.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$dgv.RowHeadersWidth = 24
$dgv.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
$form.Controls.Add($dgv)

$colOrder = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colOrder.Name = 'Order'
$colOrder.HeaderText = 'Order'
$colOrder.ReadOnly = $true
$colOrder.FillWeight = 7
[void]$dgv.Columns.Add($colOrder)
$colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colName.Name = 'Name'
$colName.HeaderText = 'Name'
$colName.FillWeight = 22
[void]$dgv.Columns.Add($colName)
$colScript = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colScript.Name = 'Script'
$colScript.HeaderText = 'Script'
$colScript.FillWeight = 26
[void]$dgv.Columns.Add($colScript)
$colObjectTypes = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colObjectTypes.Name = 'ObjectTypes'
$colObjectTypes.HeaderText = 'ObjectTypes'
# Was 22, which truncated 'UnifiedGroup, OnpremGroup' to
# 'UnifiedGroup, Onpre...' -- the longest legitimate value in the column, so
# the one case the width had to accommodate. Taken from Order and Tier, which
# hold two digits and one short word.
$colObjectTypes.FillWeight = 25
[void]$dgv.Columns.Add($colObjectTypes)
$colSection = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$colSection.Name = 'Section'
$colSection.HeaderText = 'Section'
$colSection.FillWeight = 14
$colSection.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$validSections | ForEach-Object { [void]$colSection.Items.Add($_) }
[void]$dgv.Columns.Add($colSection)
# Group folds several tools behind ONE console button.
#
# Type the same name on every tool that belongs together -- 'Rooms' on the five
# room tools -- and the console shows a single Rooms button that opens a chooser
# instead of five buttons on the front page. Blank means the tool stands alone,
# which is the default and what almost everything should stay as.
#
# The button takes the position of the FIRST member in Order, so the group lands
# where you would expect rather than wherever the alphabet puts it.
$colGroup = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colGroup.Name = 'Group'
$colGroup.HeaderText = 'Group'
$colGroup.FillWeight = 12
[void]$dgv.Columns.Add($colGroup)

# TIER IS GONE. It was free text nobody read -- ToolManager was the only file
# that mentioned it, and the help text said outright that the console ignores
# it. A column that exists but does nothing invites people to fill it in and
# then wonder why nothing happened.
#
# Existing Tier values in Tools.json are NOT carried through, so they disappear
# on the next Save. That is the intended outcome: a field with no reader is
# noise in a file that is otherwise all meaning. If a per-tool label is wanted
# later it should come back with something actually consuming it.

# THERE IS NO 'NEW WINDOW' COLUMN, DELIBERATELY.
#
# It was the original control for running a tool in its own PowerShell process.
# On 2026-08-06 an on-premises tool ran with it unticked, loaded into a console
# already connected to Exchange Online, and every write went to the cloud --
# so the requirement was moved INTO the scripts as TOOLKIT-REQUIRES: OwnProcess
# and the console obeys that.
#
# The checkbox survived as "I would like this in its own window" and was never
# used again. All 55 rows sat unticked for months with no symptom, because the
# declarations were doing the work. It looked like the control for something it
# no longer controlled.
#
# A tool that needs its own process says so in its own header. Tools.json now
# describes only what the console SHOWS; the script describes what it NEEDS.

<#
    Sorting is disabled on every column, and this is a data-safety fix rather
    than a cosmetic one.

    Text columns default to SortMode = Automatic, so clicking a header sorted
    the grid. Order is not stored data -- Set-OrderColumn renumbers it from ROW
    POSITION, and Save writes those numbers. So one click on 'Name' followed by
    Save would silently renumber all 32 tools into alphabetical order and
    rearrange every button in the console. Nothing would have looked wrong at
    the time.

    Move Up / Move Down remain the only way to reorder, which is what the note
    at the top of the window already tells you.
#>
foreach ($col in @($dgv.Columns)) {
    $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
}

function Set-OrderColumn {
    $n = 0
    foreach ($row in $dgv.Rows) {
        if ($row.IsNewRow) { continue }
        $n++
        $row.Cells['Order'].Value = "$n"
    }
}

function Update-ToolsGrid {
    param($ToolObjects)
    $dgv.Rows.Clear()
    foreach ($t in @($ToolObjects)) {
        $typesVal = (@($t.ObjectTypes) -join ', ')
        $sectionVal = Get-NormalizedSection -Value "$($t.Section)"
        # BY NAME, NOT BY POSITION.
        #
        # This was a positional Rows.Add with one value per column, in column
        # order. Adding the Group column between Section and Tier then shifted
        # every value after it one cell left: Group received Tier, Tier received
        # the NewProcess boolean and rendered it as 'False', and NewProcess --
        # the Own-window flag -- received nothing and came back UNTICKED.
        #
        # Nothing errored. Saving wrote the shifted values back, so a second
        # open compounded it, and the console grew a group button called
        # 'False'. The dangerous part was invisible: Own window silently cleared
        # on every on-premises tool is how a write lands in the wrong directory.
        #
        # Exactly the failure documented on LogViewer's grid. Naming the cells
        # costs four lines and makes column order irrelevant.
        $i = $dgv.Rows.Add()
        $row = $dgv.Rows[$i]
        $row.Cells['Name'].Value         = "$($t.Name)"
        $row.Cells['Script'].Value       = "$($t.Script)"
        $row.Cells['ObjectTypes'].Value  = $typesVal
        $row.Cells['Section'].Value      = $sectionVal
        $row.Cells['Group'].Value        = "$($t.Group)".Trim()
    }
    Set-OrderColumn
}

$dgv.Add_DefaultValuesNeeded({
    param($sender, $e)
    $e.Row.Cells['Section'].Value = (Get-ValidSections)[0]
})

function Get-ToolsFromGrid {
    $result = @()
    $n = 0
    foreach ($row in $dgv.Rows) {
        if ($row.IsNewRow) { continue }
        $name = "$($row.Cells['Name'].Value)".Trim()
        $script = "$($row.Cells['Script'].Value)".Trim()
        if ($name -eq '' -and $script -eq '') { continue }
        $n++
        $types = @(("$($row.Cells['ObjectTypes'].Value)" -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
        $section = Get-NormalizedSection -Value "$($row.Cells['Section'].Value)"
        # Ordered, so the written file keeps a predictable field order and
        # optional fields can be OMITTED rather than written as "". Emitting an
        # empty value on every tool would add a line to all fifty-odd entries
        # the first time anyone opened ToolManager and hit Save -- a whole-file
        # diff for nothing. Config\ is the operator's, and a config editor
        # should change what you changed.
        #
        # ScriptOnPrem used to be written here. Removed 2026-08-13 with the
        # rest of the AD twin feature; any value still in Tools.json is now
        # ignored on read and dropped on the next Save.
        $fields = [ordered]@{
            Order       = $n
            Name        = $name
            Script      = $script
        }
        $fields['ObjectTypes'] = $types
        $fields['Section']     = $section
        # Written only when set. An empty Group on every tool would add a dead
        # key to sixty entries and make a diff of Tools.json unreadable.
        $grp = "$($row.Cells['Group'].Value)".Trim()
        if ($grp) { $fields['Group'] = $grp }

        $result += [PSCustomObject]$fields
    }
    return $result
}

$script:CurrentTools = @(Read-ToolsJson -Path $ToolsConfigPath)
Update-ToolsGrid -ToolObjects $script:CurrentTools

# THERE IS NO HINT LABEL, DELIBERATELY.
#
# A 700-character paragraph used to sit under the grid explaining every column.
# It was the only documentation this tool had, so it had to carry everything --
# and it did that badly, clipped mid-word in a strip too short for it, and
# silently went stale when the Group column was added without anyone updating
# it.
#
# All of it now lives in the ? button, which is also where every other tool
# keeps its explanation. Leaving a summary on screen as well would defeat that:
# permanent on-screen text is the reason nobody opens help, and the copy that
# gets read is then the copy that nobody maintains.
#
# The space it occupied goes to the grid.

$btnAddRow = New-Object System.Windows.Forms.Button
$btnAddRow.Text = "Add Tool"
$btnAddRow.Location = New-Object System.Drawing.Point(12, 460)
$btnAddRow.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnAddRow)
$btnDeleteRow = New-Object System.Windows.Forms.Button
$btnDeleteRow.Text = "Delete Selected"
$btnDeleteRow.Location = New-Object System.Drawing.Point(120, 460)
$btnDeleteRow.Size = New-Object System.Drawing.Size(120, 28)
$form.Controls.Add($btnDeleteRow)

$btnMoveUp = New-Object System.Windows.Forms.Button
$btnMoveUp.Text = "Move Up"
$btnMoveUp.Location = New-Object System.Drawing.Point(12, 496)
$btnMoveUp.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnMoveUp)
$btnMoveDown = New-Object System.Windows.Forms.Button
$btnMoveDown.Text = "Move Down"
$btnMoveDown.Location = New-Object System.Drawing.Point(120, 496)
$btnMoveDown.Size = New-Object System.Drawing.Size(100, 28)
$form.Controls.Add($btnMoveDown)

$btnReload = New-Object System.Windows.Forms.Button
$btnReload.Text = "Reload from Disk"
$btnReload.Location = New-Object System.Drawing.Point(228, 496)
$btnReload.Size = New-Object System.Drawing.Size(120, 28)
$form.Controls.Add($btnReload)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "$($script:CurrentTools.Count) tool(s) loaded from $ToolsConfigPath"
$lblStatus.Location = New-Object System.Drawing.Point(360, 502)
$lblStatus.Size = New-Object System.Drawing.Size(300, 20)
# Width is recomputed in the layout block below to fill the gap up to Save.
# AutoEllipsis so that if the path is still too long for the window the line
# ends in '...' -- a visible signal there is more, rather than a sentence that
# stops mid-path and reads as complete.
$lblStatus.AutoEllipsis = $true
$form.Controls.Add($lblStatus)

$btnSections = New-Object System.Windows.Forms.Button
$btnSections.Text = "Sections..."
$btnSections.Location = New-Object System.Drawing.Point(12, 532)
$btnSections.Size = New-Object System.Drawing.Size(100, 28)
$btnSections.Add_Click({
    if (Show-SectionEditor) {
        # Rebuild the Section dropdown so newly-added sections are selectable
        # immediately, and warn about tools stranded on a deleted section.
        $valid = Get-ValidSections
        $colSection.Items.Clear()
        $valid | ForEach-Object { [void]$colSection.Items.Add($_) }
        $script:sectionCapacity = Get-SectionCapacityTable
        $stranded = @()
        foreach ($row in $dgv.Rows) {
            if ($row.IsNewRow) { continue }
            $cur = "$($row.Cells['Section'].Value)".Trim()
            if ($valid -notcontains $cur) {
                $stranded += "$($row.Cells['Name'].Value) (was '$cur')"
                $row.Cells['Section'].Value = $valid[0]
            }
        }
        $lblStatus.Text = "Sections saved. Reopen the Admin Console to see the new layout."
        if ($stranded.Count -gt 0) {
            [void][System.Windows.Forms.MessageBox]::Show(
                "These tools were in a section that no longer exists and have been moved to '$($valid[0])'. Click Save to keep that change:`n`n$($stranded -join "`n")",
                'Sections Changed', [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }
})
$form.Controls.Add($btnSections)

$btnTheme = New-Object System.Windows.Forms.Button
$btnTheme.Text = "Theme..."
$btnTheme.Location = New-Object System.Drawing.Point(120, 532)
$btnTheme.Size = New-Object System.Drawing.Size(100, 28)
$btnTheme.Add_Click({
    $p = Join-Path $PSScriptRoot 'ThemeManager.ps1'
    if (Test-Path -LiteralPath $p) {
        & $p
    } else {
        [System.Windows.Forms.MessageBox]::Show(
            "ThemeManager.ps1 was not found next to this script.",
            'Tool Manager', 'OK', 'Warning') | Out-Null
    }
})
$form.Controls.Add($btnTheme)

#region Settings dialog -------------------------------------------------------
# Show-SettingsEditor MOVED TO Common.ps1, 2026-08-20.
#
# It was only ever reachable from this window, which made "change my email
# address" a trip through the tool registry. The dashboard now has its own
# Settings button and Settings.ps1 opens the same dialog -- so the dialog had to
# stop belonging to one caller.
#
# COPIED WOULD HAVE BEEN WRONG. Two settings editors drift, and the one you are
# not looking at is the one that writes the file. Common.ps1 already held every
# helper it calls (Get-ToolkitSettings, Save-ToolkitSettings, Get-EdgeProfiles,
# Get-SpoAdminUrl, Get-ExchangeServerUri), so it moved without changing a line.

$btnSettings          = New-Object System.Windows.Forms.Button
$btnSettings.Text     = 'Settings...'
$btnSettings.Location = New-Object System.Drawing.Point(228, 532)
$btnSettings.Size     = New-Object System.Drawing.Size(100, 28)
$btnSettings.Add_Click({
    if (Show-SettingsEditor) {
        $lblStatus.Text = "Settings saved. Reports -> $(Get-ReportPath)"
    }
})
$form.Controls.Add($btnSettings)
#endregion


$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "Save"
$btnSave.Location = New-Object System.Drawing.Point(672, 496)
$btnSave.Size = New-Object System.Drawing.Size(90, 28)
$btnSave.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnSave)
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(770, 496)
$btnClose.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnClose)

#region Layout and resize behaviour -------------------------------------------
<#
    Positions are recomputed from the CURRENT client size, then anchors are
    applied -- in that order, and the order matters.

    An anchor preserves the distance between a control and the edge it is
    anchored to AS AT THE MOMENT THE ANCHOR IS SET. The buttons were laid out
    for a 640-tall form; anchoring them to Bottom while the form is already 700
    tall would have frozen that old spacing and left a 60px dead strip along
    the bottom. So everything is moved into place first.

    The grid takes all the growth. That is the actual fix for this window: with
    32 tools in a fixed 380px viewport, nine rows were visible at a time, and no
    amount of contrast helps when you are scrolling a list you are trying to
    compare against itself.
#>
$cw = $form.ClientSize.Width
$ch = $form.ClientSize.Height

$rowC = $ch - 40          # Sections... / Theme...  and  Save / Close
$rowB = $rowC - 36        # Move Up / Move Down / Reload
$rowA = $rowB - 36        # Add Tool / Delete Selected
# 44, not 24. The hint wraps to two lines at this width and the old figure gave
# The grid now runs all the way down to the button row -- the hint label that
# used to sit between them is gone, so there is nothing to reserve space for.
$lblNote.Width = $cw - 24
$dgv.Width     = $cw - 24
$dgv.Height    = $rowA - 8 - $dgv.Top

$btnAddRow.Top    = $rowA
$btnDeleteRow.Top = $rowA
$btnMoveUp.Top    = $rowB
$btnMoveDown.Top  = $rowB
$btnReload.Top    = $rowB
$lblStatus.Top    = $rowB + 6
$btnSections.Top  = $rowC
$btnTheme.Top     = $rowC
$btnSettings.Top  = $rowC

# Save and Close stay on the Move row but pinned right, so they keep their
# distance from the edge as the window widens.
$btnSave.Top  = $rowB
$btnClose.Top = $rowB
$btnClose.Left = $cw - 12 - $btnClose.Width
$btnSave.Left  = $btnClose.Left - 8 - $btnSave.Width

# The status line fills whatever is left between Reload and Save, rather than a
# fixed 300px that was already too short for a full path and got shorter every
# time the window grew.
$lblStatus.Width = [Math]::Max(200, ($btnSave.Left - 12 - $lblStatus.Left))

$dgv.Anchor       = 'Top,Left,Bottom,Right'
$lblNote.Anchor   = 'Top,Left,Right'
foreach ($b in @($btnAddRow, $btnDeleteRow, $btnMoveUp, $btnMoveDown, $btnReload,
                 $btnSections, $btnTheme, $btnSettings)) {
    $b.Anchor = 'Bottom,Left'
}
$lblStatus.Anchor = 'Bottom,Left'
$btnSave.Anchor   = 'Bottom,Right'
$btnClose.Anchor  = 'Bottom,Right'
#endregion

function Move-SelectedRow {
    param([int]$Delta)
    if ($dgv.SelectedRows.Count -eq 0) { return }
    $idx = $dgv.SelectedRows[0].Index
    if ($dgv.Rows[$idx].IsNewRow) { return }
    $targetIdx = $idx + $Delta
    if ($targetIdx -lt 0 -or $targetIdx -ge $dgv.Rows.Count -or $dgv.Rows[$targetIdx].IsNewRow) { return }
    # EVERY DATA COLUMN, DERIVED -- NOT A HAND-MAINTAINED LIST.
    #
    # This was a literal list: Name, Script, ObjectTypes, Section, Group. When
    # the ScriptOnPrem column was added nobody added it here, so Move Up/Down
    # carried a tool's identity to the new row and left its AD twin sitting on
    # the OLD row position, where it silently reattached to whatever landed
    # there. Message Trace, DL Send To/As and Connect-ExSE all acquired twins
    # they never had, Save wrote it back, and it compounded every session.
    #
    # The twin feature is gone now (ARCHITECTURE section 30), but the trap was
    # never really about twins: it was a list of column names that had to be
    # updated by hand whenever a column was added, with nothing to enforce it
    # and silent data corruption as the failure mode. The same omission also
    # produced the Group/Tier/NewProcess shift documented in Update-ToolsGrid.
    #
    # Asking the grid for its own columns cannot fall out of date. Order is
    # excluded because Set-OrderColumn recomputes it from row position
    # immediately below -- swapping it would be undone anyway.
    foreach ($col in @($dgv.Columns)) {
        if ($col.Name -eq 'Order') { continue }
        $a = $dgv.Rows[$idx].Cells[$col.Name].Value
        $b = $dgv.Rows[$targetIdx].Cells[$col.Name].Value
        $dgv.Rows[$idx].Cells[$col.Name].Value = $b
        $dgv.Rows[$targetIdx].Cells[$col.Name].Value = $a
    }
    Set-OrderColumn
    $dgv.ClearSelection()
    $dgv.Rows[$targetIdx].Selected = $true
    $dgv.CurrentCell = $dgv.Rows[$targetIdx].Cells['Name']
}
$btnMoveUp.Add_Click({ Move-SelectedRow -Delta -1 })
$btnMoveDown.Add_Click({ Move-SelectedRow -Delta 1 })

$btnAddRow.Add_Click({
    if ($dgv.IsCurrentCellInEditMode) { $dgv.EndEdit() }
    $newRowIndex = $dgv.Rows.Count - 1
    $dgv.CurrentCell = $dgv.Rows[$newRowIndex].Cells['Name']
    $dgv.FirstDisplayedScrollingRowIndex = $newRowIndex
    $dgv.BeginEdit($true)
})

$btnDeleteRow.Add_Click({
    if ($dgv.IsCurrentCellInEditMode) { $dgv.EndEdit() }
    if ($dgv.SelectedRows.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Select a row first.", 'Delete Tool', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $idx = $dgv.SelectedRows[0].Index
    if ($dgv.Rows[$idx].IsNewRow) { return }
    $name = "$($dgv.Rows[$idx].Cells['Name'].Value)"
    $label = if ($name -ne '') { "'$name'" } else { 'this row' }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Delete $label from the grid? This isn't written to Tools.json until you click Save.",
        'Confirm Delete', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    $dgv.Rows.RemoveAt($idx)
    Set-OrderColumn
})

$btnReload.Add_Click({
    if ($dgv.IsCurrentCellInEditMode) { $dgv.EndEdit() }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Reload from '$ToolsConfigPath'? Any unsaved edits in the grid will be discarded.",
        'Confirm Reload', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    try {
        $script:CurrentTools = @(Read-ToolsJson -Path $ToolsConfigPath)
        Update-ToolsGrid -ToolObjects $script:CurrentTools
        $lblStatus.Text = "Reloaded $($script:CurrentTools.Count) tool(s) from $ToolsConfigPath at $(Get-Date -Format 'HH:mm:ss')."
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Reload failed: $($_.Exception.Message)",
            'Reload Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$btnSave.Add_Click({
    if ($dgv.IsCurrentCellInEditMode) { $dgv.EndEdit() }
    try {
        $tools = Get-ToolsFromGrid
        $badTypes = @($tools | ForEach-Object { $_.ObjectTypes } | Where-Object { $validObjectTypes -notcontains $_ } | Select-Object -Unique)
        if ($badTypes.Count -gt 0) {
            [void][System.Windows.Forms.MessageBox]::Show(
                "Unrecognized ObjectTypes value(s): $($badTypes -join ', ')`n`nValid values: $($validObjectTypes -join ', ')",
                'Validation Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $blankNames = @($tools | Where-Object { $_.Name -eq '' -or $_.Script -eq '' })
        if ($blankNames.Count -gt 0) {
            [void][System.Windows.Forms.MessageBox]::Show(
                "Every tool needs both a Name and a Script path.",
                'Validation Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        $overCapacity = @($script:sectionCapacity.Keys | ForEach-Object {
            $section = $_
            $count = @($tools | Where-Object { $_.Section -eq $section }).Count
            if ($count -gt $script:sectionCapacity[$section]) {
                "$section`: $count tool(s), grid holds $($script:sectionCapacity[$section])"
            }
        } | Where-Object { $_ })
        if ($overCapacity.Count -gt 0) {
            [void][System.Windows.Forms.MessageBox]::Show(
                "This will still save -- AdminConsole's button panels scroll, so nothing is hidden -- but the following section(s) now exceed the grid size they're sized for:`n`n$($overCapacity -join "`n")",
                'Section Over Capacity', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        Write-ToolsJson -Path $ToolsConfigPath -ToolObjects $tools
        $script:CurrentTools = $tools
        Set-OrderColumn
        $lblStatus.Text = "Saved $($tools.Count) tool(s) to $ToolsConfigPath at $(Get-Date -Format 'HH:mm:ss')."
    } catch {
        [void][System.Windows.Forms.MessageBox]::Show(
            "Save failed: $($_.Exception.Message)",
            'Save Error', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})
$btnClose.Add_Click({ $form.Close() })

# Wrapped, like TenantManager's. A missing or unreadable help block must not
# stop the editor opening -- the tool is how you fix a broken Tools.json, so it
# has to survive a toolkit that is not entirely well.
try { Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Tool Manager' | Out-Null } catch { }

Set-ConsoleTheme -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
[void]$form.ShowDialog()
#endregion
```


<br>

==============================================================================

# ◆ FILE 5 OF 16 — `Connect-ExO.ps1`

> **251 lines.** Cloud connection. Encodes B2 (a modal loop blocks any window a tool raises) and C17 (connections are per-process and cannot be inherited by a child).
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```powershell
#Requires -Version 5.1
<#
    Connect-ExO.ps1
    Connects this PowerShell process to Exchange Online.

    Reasoning: ScriptNotes/Connect-Scripts.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-03
#>
# TOOLKIT-HELP-BEGIN
# Connects this console to Exchange Online.
#
# The console closes, the sign-in happens, and the console reopens with the
# connection live. That is deliberate -- a sign-in window cannot accept typing
# while the console's own window is open in front of it, so the console gets out
# of the way.
#
# Complete the sign-in and MFA prompt when it appears; the console comes back on
# its own.
# TOOLKIT-HELP-END

param(
    [string]$UserPrincipalName,
    [switch]$Force
)

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
#
# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost: the window flashes up, work may have started, and
# the operator has seen a tool they were not meant to be in.
if (-not (Test-ConsoleLaunch)) { return }

# No WinForms. This script must NOT open a window of its own -- the sign-in
# prompt is the window, and anything of ours competing with it is the problem
# this script exists to avoid.

function Write-Line {
    param([string]$Text, [string]$Colour)
    if ($Colour) { Write-Host $Text -ForegroundColor $Colour } else { Write-Host $Text }
}

Write-Line ''
Write-Line '=== Connect to Exchange Online ===' 'Cyan'

$before = Test-ExoConnectionState
if ($before.Connected -and -not $Force) {
    Write-Line "Already connected. $($before.Detail)" 'Green'
    Write-Line 'Nothing to do. Re-run with -Force to reconnect.'
    Write-Line ''
    return
}
if (-not $before.Connected) { Write-Line $before.Detail 'DarkGray' }

# MODULE FIRST, SIGN-IN SECOND.
#
# This used to print the Install-Module line and stop, which is correct and
# useless: the operators these guardrails exist for read "not installed" and
# stop too. Three states, three different answers -- and only one of them
# involves downloading anything.
$mod = Test-ToolkitModule -Id 'ExO'

if ($mod.State -eq 'Installed') {
    # On disk, just not in this process. Free to fix and not worth asking about.
    Write-Line $mod.Detail 'DarkGray'
    $imp = Import-ToolkitModule -Id 'ExO'
    Write-Line $imp.Detail $(if ($imp.Ok) { 'Green' } else { 'Red' })
    if (-not $imp.Ok) {
        Write-Line ''
        Write-CompletedLog -Action 'Connect' -Target 'Exchange Online' -Result 'Failed' -Detail $imp.Detail
        return
    }
    $mod = Test-ToolkitModule -Id 'ExO'
}

if ($mod.State -ne 'Imported') {
    Write-Line "$($mod.Name) is not installed for your account." 'Yellow'

    # ASKED, NEVER ASSUMED. A tool that quietly installs software is a
    # different thing from a tool that offers to.
    if (-not (Request-ToolkitModuleInstall -Id 'ExO')) {
        Write-Line 'Not installed -- you said no. Nothing was changed.' 'DarkGray'
        Write-Line "To do it yourself later:  Install-Module $($mod.Name) -Scope CurrentUser"
        Write-Line ''
        Write-CompletedLog -Action 'Connect' -Target 'Exchange Online' -Result 'Failed' `
            -Detail "$($mod.Name) not installed; operator declined the install."
        return
    }

    Write-Line 'Installing. This can take a minute or two the first time.' 'Cyan'
    $ins = Install-ToolkitModule -Id 'ExO'
    foreach ($s in @($ins.Steps)) { Write-Line "  $s" 'DarkGray' }

    if (-not $ins.Ok) {
        Write-Line $ins.Detail 'Red'
        Write-Line ''
        Write-Line 'If this machine cannot reach the PowerShell Gallery, your desktop team' 'Yellow'
        Write-Line 'will need to install it for you. The command is:' 'Yellow'
        Write-Line "    Install-Module $($mod.Name) -Scope CurrentUser"
        Write-Line ''
        Write-CompletedLog -Action 'Connect' -Target 'Exchange Online' -Result 'Failed' -Detail $ins.Detail
        return
    }

    Write-Line $ins.Detail 'Green'
    # The dashboard's cached view is now wrong -- it was built before the
    # install. Clearing it is cheaper than letting it report Missing for a
    # module that is now loaded.
    $global:ToolkitModuleStateMap = $null
} else {
    Write-Line $mod.Detail 'DarkGray'
}

$params = @{ ShowBanner = $false; ErrorAction = 'Stop' }

# -DisableWAM is the point of this script.
#
# From version 3.7.0 the module signs in through Web Account Manager, which
# raises its own broker window. When the Admin Console launched a connect
# script from a button, that window appeared while $form.ShowDialog() held the
# thread in a modal message loop -- and a modal form disables other windows on
# the same thread. The prompt drew and then never accepted input, which reads
# as "the MFA page never finished loading".
#
# Microsoft document the switch for exactly this class of problem: "If you
# encounter WAM-related issues during sign in, you can use this switch to
# disable WAM." Without WAM the sign-in goes to a browser instead, which is a
# separate process and unaffected by our message loop.
#
# Version-guarded because the parameter only exists from 3.7.2-Preview1. Passing
# it to an older module is a binding error, and "your connect script is broken"
# is a worse outcome than a prompt that needs a second click.
if ((Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')) {
    $params['DisableWAM'] = $true
    Write-Line 'Using -DisableWAM (sign-in goes to the browser, not the WAM broker).' 'DarkGray'
} else {
    Write-Line 'This module version has no -DisableWAM; using the default sign-in path.' 'DarkGray'
}

# Skips the username page. Not cosmetic on a shared jump box, where the broker
# may remember somebody else's account.
if ($UserPrincipalName) { $params['UserPrincipalName'] = $UserPrincipalName }

# The sign-in window is an ActiveX WebBrowser control, and it can only be
# created on an STA thread that is not carrying somebody else's synchronization
# context.
#
# This script is started from AdminConsole immediately after ShowDialog()
# returns. ShowDialog installs a WindowsFormsSynchronizationContext on the
# thread and does NOT remove it when the form closes, so MSAL captures a context
# whose message pump has already gone. That is the reported failure:
#
#   ActiveX control '8856f961-340a-11d0-a96b-00c04fd705a2' cannot be
#   instantiated because the current thread is not in a single-threaded
#   apartment
#
# 8856f961- is the IE WebBrowser control, and the trace ends in
# WinFormsLegacyWebUi.OnAuthenticate -- the embedded browser that -DisableWAM
# selects.
#
# Clearing the context is cheap and safe: nothing after this point needs it, and
# the console is rebuilt from scratch when it reopens.
$apartment = [System.Threading.Thread]::CurrentThread.GetApartmentState()
$staleCtx  = [System.Threading.SynchronizationContext]::Current
if ($staleCtx) {
    Write-Line "Clearing a leftover $($staleCtx.GetType().Name) before sign-in." 'DarkGray'
    [System.Threading.SynchronizationContext]::SetSynchronizationContext($null)
}
if ($apartment -ne [System.Threading.ApartmentState]::STA) {
    Write-Line ''
    Write-Line "This PowerShell thread is $apartment, not STA. The sign-in window cannot open here." 'Red'
    Write-Line 'Start the console from powershell.exe (STA by default) rather than with -MTA, or from pwsh with -STA.' 'Yellow'
    Write-CompletedLog -Action 'Connect' -Target 'Exchange Online' -Result 'Failed' `
        -Detail "Thread apartment is $apartment; interactive sign-in requires STA."
    Write-Line ''
    return
}

Write-Line ''
Write-Line 'A sign-in window will open. Complete MFA there, then return here.' 'Yellow'
Write-Line ''

# ErrorAction is ALREADY in $params (line 60). Do not add it here as well --
# splatting a key and passing the same parameter explicitly is a binding error,
# "parameter 'ErrorAction' is specified more than once", and it fails before the
# module is even called. That broke every connect attempt until it was spotted.
#
# 2>$null suppresses the error STREAM, not the exception. The module writes
# forty lines of MSAL async plumbing to it before the terminating error
# surfaces, and none of those lines names a cause -- only the last one does, and
# the catch below prints it.
try {
    Connect-ExchangeOnline @params 2>$null
} catch {
    $msg = "$($_.Exception.Message)"

    # One targeted retry, and ONLY for this error. The context is already
    # cleared above, so a second attempt starts clean. Retrying anything else
    # would just hide it -- which is why this checks the message rather than
    # wrapping every failure in a loop.
    if ($msg -match 'single-threaded apartment|8856f961') {
        Write-Line ''
        Write-Line 'The sign-in window could not be created on this thread. Retrying once with a clean context.' 'Yellow'
        [System.Threading.SynchronizationContext]::SetSynchronizationContext($null)
        try {
            Connect-ExchangeOnline @params 2>$null
            $msg = $null
        } catch {
            $msg = "$($_.Exception.Message)"
        }
    }

    if ($msg) {
        Write-Line ''
        Write-Line "Connection failed: $msg" 'Red'
        Write-CompletedLog -Action 'Connect' -Target 'Exchange Online' -Result 'Failed' -Detail $msg
        Write-Line ''
        return
    }
}

# Prove it rather than assume it. Connect-ExchangeOnline returning without an
# error is not the same as a usable session, and the console's buttons are
# about to be re-coloured from this same probe -- so the message here and the
# button state cannot disagree.
$after = Test-ExoConnectionState
Write-Line ''
if ($after.Connected) {
    Write-Line "Connected. $($after.Detail)" 'Green'
    Write-CompletedLog -Action 'Connect' -Target 'Exchange Online' -Result 'Success' `
        -Detail "$($after.Detail)"
} else {
    Write-Line "Sign-in completed but no session is present: $($after.Detail)" 'Red'
    Write-CompletedLog -Action 'Connect' -Target 'Exchange Online' -Result 'Failed' `
        -Detail "$($after.Detail)"
}

# The tenant is cached from the first thing that asked. Connecting is exactly
# the event that can change the answer, so the cache is dropped here rather
# than left to go stale behind a correct-looking label.
Clear-ConnectedTenantDomainCache
$mismatch = Test-TenantDomainMismatch
if ($mismatch) {
    Write-Line ''
    Write-Line $mismatch 'Yellow'
}
Write-Line ''
```


<br>

==============================================================================

# ◆ FILE 6 OF 16 — `Connect-ExSE.ps1`

> **203 lines.** On-premises connection. Session setup, -ViewEntireForest asserted rather than assumed (C6), and why it avoids Get-Credential entirely.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```powershell
#Requires -Version 5.1
<#
    Connect-ExSE.ps1
    Connects this PowerShell process to on-premises Exchange (Server SE)
    by importing a remote PowerShell session from the hybrid server.

    Reasoning: ScriptNotes/Connect-Scripts.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-03
#>
# TOOLKIT-HELP-BEGIN
# Connects this console to on-premises Exchange.
#
# Uses your current Windows sign-in -- no password prompt -- by importing a
# session from the hybrid server.
#
# Needed by anything that touches the on-premises directory: distribution list
# management, new starters, terminations and the on-premises reports.
# TOOLKIT-HELP-END

param(
    [string]$ConnectionUri,
    [switch]$PromptForCredential,
    [switch]$Force
)

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
#
# FIRST THING AFTER Common LOADS. A guard that runs after the form is built is
# a guard that already lost: the window flashes up, work may have started, and
# the operator has seen a tool they were not meant to be in.
if (-not (Test-ConsoleLaunch)) { return }

function Write-Line {
    param([string]$Text, [string]$Colour)
    if ($Colour) { Write-Host $Text -ForegroundColor $Colour } else { Write-Host $Text }
}

Write-Line ''
Write-Line '=== Connect to on-premises Exchange (Server SE) ===' 'Cyan'

$before = Test-EmsAvailable
if ($before.Available -and -not $Force) {
    Write-Line "Already connected. $($before.Detail)" 'Green'
    Write-Line 'Nothing to do. Re-run with -Force to reconnect.'
    Write-Line ''
    return
}

# Importing on-prem Exchange into a window that already holds Exchange Online
# is the collision the whole toolkit is careful about (ARCHITECTURE section 2).
# Get-Mailbox, Set-Mailbox and Get-DistributionGroup exist in BOTH, and after
# -AllowClobber the on-prem copies win by load order -- silently, at every call
# site, with nothing on screen to say which world answered.
#
# Not blocked, because a hybrid admin legitimately needs both and knows it. But
# it is said out loud, once, at the moment it becomes true.
$exo = Test-ExoConnectionState
if ($exo.Connected) {
    Write-Line ''
    Write-Line 'WARNING: Exchange Online is already connected in this window.' 'Yellow'
    Write-Line '  Get-Mailbox, Set-Mailbox and Get-DistributionGroup exist in both.' 'Yellow'
    Write-Line '  After this import the ON-PREMISES versions will answer.' 'Yellow'
    Write-Line '  For cloud work, use a separate PowerShell window.' 'Yellow'
    Write-Line ''
}

if (-not $ConnectionUri) { $ConnectionUri = Get-ExchangeServerUri }
if (-not $ConnectionUri) {
    Write-Line 'No Exchange server configured. Set it in ToolManager -> Settings..., or pass -ConnectionUri.' 'Red'
    Write-Line ''
    return
}
Write-Line "Endpoint: $ConnectionUri" 'DarkGray'

$sessionParams = @{
    ConfigurationName = 'Microsoft.Exchange'
    ConnectionUri     = $ConnectionUri
    Authentication    = 'Kerberos'
    ErrorAction       = 'Stop'
}

# No Get-Credential by default, and that is the important difference from the
# script this replaces.
#
# Two reasons. First, -Credential is not mandatory on New-PSSession -- the
# documented default is "Current user" -- so with Kerberos the logged-on
# ticket is used, which is the normal single-sign-on case on a jump box and
# needs no prompt at all.
#
# Second, a prompt is the thing that breaks here. Get-Credential raises a
# Windows credential dialog, and when this runs from a console button the
# console is inside $form.ShowDialog() -- a modal loop that disables other
# windows on the same thread. That is exactly what made the Exchange Online
# and SharePoint sign-in prompts hang. Not prompting sidesteps it entirely.
#
# -PromptForCredential is there for the case that actually needs it: signing in
# as a different admin account. If the dialog does not appear, run this script
# directly from the prompt rather than from the console button.
if ($PromptForCredential) {
    Write-Line ''
    Write-Line 'A credential prompt will open. If it does not appear, close the console' 'Yellow'
    Write-Line 'and run .\Connect-ExSE.ps1 -PromptForCredential from the prompt instead.' 'Yellow'
    $cred = Get-Credential -Message "Exchange on-premises ($ConnectionUri)"
    if (-not $cred) {
        Write-Line 'Cancelled -- no credentials entered.' 'DarkGray'
        Write-Line ''
        return
    }
    $sessionParams['Credential'] = $cred
} else {
    Write-Line "Using the logged-on account ($env:USERDOMAIN\$env:USERNAME) via Kerberos." 'DarkGray'
    Write-Line 'Use -PromptForCredential to sign in as a different account.' 'DarkGray'
}

Write-Line ''
Write-Line 'Opening the remote session...' 'DarkGray'

$session = $null
try {
    $session = New-PSSession @sessionParams
} catch {
    Write-Line ''
    Write-Line "Could not open the session: $($_.Exception.Message)" 'Red'
    Write-Line ''
    Write-Line 'Common causes:' 'DarkGray'
    Write-Line '  - The logged-on account has no Exchange role. Retry with -PromptForCredential.' 'DarkGray'
    Write-Line '  - The server name is wrong. Check ToolManager -> Settings...' 'DarkGray'
    Write-Line '  - Kerberos needs the FQDN. A short host name will not authenticate.' 'DarkGray'
    Write-CompletedLog -Action 'Connect' -Target 'Exchange on-premises' -Result 'Failed' `
        -Detail "$ConnectionUri -- $($_.Exception.Message)" -Environment 'On-Prem'
    Write-Line ''
    return
}

try {
    # THIS is where the button version of the old script would have failed
    # silently.
    #
    # Import-PSSession has no -Global parameter and no documented scope
    # behaviour. It builds a temporary module and returns the PSModuleInfo. Run
    # interactively, its commands land where you want them. Run inside a .ps1
    # launched with '&' -- which is how every console button works -- there is
    # no guarantee they outlive the script, and in practice they do not. The
    # script would report success and leave the window with no Exchange
    # cmdlets in it.
    #
    # Import-Module -Global is the documented fix: "imports modules into the
    # global session state so they're available to all commands in the
    # session." Capture the module, then re-import it globally.
    #
    # -AllowClobber because Exchange Online may already own Get-Mailbox and
    # friends; without it those commands are skipped and the import looks
    # partial for no visible reason. See the warning above -- this is the
    # deliberate half of that collision.
    $module = Import-PSSession -Session $session -DisableNameChecking -AllowClobber -ErrorAction Stop
    Import-Module -ModuleInfo $module -Global -DisableNameChecking -ErrorAction Stop
} catch {
    Write-Line ''
    Write-Line "Session opened but the import failed: $($_.Exception.Message)" 'Red'
    try { Remove-PSSession -Session $session -ErrorAction SilentlyContinue } catch { }
    Write-CompletedLog -Action 'Connect' -Target 'Exchange on-premises' -Result 'Failed' `
        -Detail "$ConnectionUri -- import failed: $($_.Exception.Message)" -Environment 'On-Prem'
    Write-Line ''
    return
}

# Held globally so it survives this script's scope and can be found later.
# Microsoft: "you must keep the connection to the other session open. If you
# close it, the imported commands are no longer available." A session nobody
# holds a reference to is a session nobody can deliberately close.
$global:ExchangeOnPremSession = $session

# The forest-wide view is the reason this console can look up objects in any
# directory without being told which. Without it, recipient lookups silently return only
# the local domain -- an empty result that looks like "no such mailbox".
try {
    Set-ADServerSettings -ViewEntireForest $true -ErrorAction Stop
    Write-Line 'View entire forest: on.' 'DarkGray'
} catch {
    Write-Line "Connected, but Set-ADServerSettings failed: $($_.Exception.Message)" 'Yellow'
    Write-Line 'Lookups may only see the local domain.' 'Yellow'
}

$after = Test-EmsAvailable
Write-Line ''
if ($after.Available) {
    Write-Line "Connected. $($after.Detail)" 'Green'
    Write-CompletedLog -Action 'Connect' -Target 'Exchange on-premises' -Result 'Success' `
        -Detail $ConnectionUri -Environment 'On-Prem'
} else {
    # The import reported success and the probe disagrees. Almost always the
    # scope problem above, so name it rather than leaving a puzzle.
    Write-Line "Import completed but the cmdlets are not present: $($after.Detail)" 'Red'
    Write-Line 'If this persists, run this script directly rather than from the console button.' 'Yellow'
    Write-CompletedLog -Action 'Connect' -Target 'Exchange on-premises' -Result 'Failed' `
        -Detail "$ConnectionUri -- imported but cmdlets absent" -Environment 'On-Prem'
}
Write-Line ''
```


<br>

==============================================================================

# ◆ FILE 7 OF 16 — `MailboxHealth.ps1`

> **373 lines.** THE CANONICAL SMALL TOOL and the first one to build. The whole required shape at a glance: guard, work, form, wrapped handler, audit.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```powershell
#Requires -Version 5.1
<#
    MailboxHealth.ps1
    Standalone Mailbox Health tool -- Layer 2 reference template for the
    launcher/tool-script architecture refactor.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-BEGIN
# A read-only snapshot of one mailbox -- the first place to look when someone
# says "my email is doing something strange".
#
# Shows size and quota, the mailbox type, forwarding, delegates and
# permissions, auto-reply state, and the litigation hold and archive status
# in one view, so you are not running six commands to answer one question.
#
# Read-only. Nothing here changes anything -- it is for finding out what is
# true before you decide what to do.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
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

# NO PROMPT BEFORE THE WINDOW. The address box is ON the form -- see
# Add-ToolTargetBar in Common.ps1. Resolve-ToolTarget is not used at all now:
# its job was to decide which DIRECTORY an address belongs to, and this tool
# reads Exchange Online, which does not need to be told.
#
# An empty address is fine. The window opens and asks.
$script:Target = Get-TargetSmtp -Text $SmtpAddress

#region Actions - mailbox health (moved unchanged from Mailbox-Group-Inspector.ps1 lines 1556-1851)
function Get-MailboxHealthSnapshot {
    param([string]$Smtp)
    $health = [ordered]@{
        MailboxType                = $null
        RecipientTypeDetails       = $null
        LicenseState               = $null
        LastLogonTime              = $null
        ItemCount                  = $null
        MailboxSize                = $null
        ProhibitSendQuota          = $null
        ProhibitSendReceiveQuota   = $null
        ArchiveEnabled             = $null
        ArchiveSize                = $null
        AccountDisabled            = $null
        ForwardingSmtpAddress      = $null
        ForwardingAddress          = $null
        DeliverToMailboxAndForward = $null
        LitigationHoldEnabled      = $null
        RetentionHoldEnabled       = $null
        InPlaceHoldsCount          = $null
        FullAccessDelegates        = @()
        SendAsDelegates            = @()
        RuleCount                  = $null
        SuspiciousRules            = @()
        OWAEnabled                 = $null
        ActiveSyncEnabled          = $null
        PopEnabled                 = $null
        ImapEnabled                = $null
        MAPIEnabled                = $null
        EwsEnabled                 = $null
        Errors                     = @()
    }
    try {
        $stats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $health.LastLogonTime = $stats.LastLogonTime
        $health.ItemCount     = $stats.ItemCount
        $health.MailboxSize   = $stats.TotalItemSize
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics (health)')
    }
    try {
        $mbx = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }

        # FREE -- this came back on the Get-Mailbox call that was already being
        # made. RecipientTypeDetails is kept alongside the simplified label
        # because the raw value distinguishes RoomMailbox from EquipmentMailbox
        # and SchedulingMailbox, and an operator chasing a booking problem needs
        # to know which one it is.
        $health.RecipientTypeDetails = "$($mbx.RecipientTypeDetails)"
        $health.MailboxType          = Get-SimpleMailboxType -RecipientTypeDetails $health.RecipientTypeDetails

        $health.ArchiveEnabled             = ($mbx.ArchiveGuid -and $mbx.ArchiveGuid -ne [Guid]::Empty)
        $health.AccountDisabled            = $mbx.AccountDisabled
        $health.ProhibitSendQuota          = $mbx.ProhibitSendQuota
        $health.ProhibitSendReceiveQuota   = $mbx.ProhibitSendReceiveQuota
        $health.ForwardingSmtpAddress      = $mbx.ForwardingSmtpAddress
        $health.ForwardingAddress          = $mbx.ForwardingAddress
        $health.DeliverToMailboxAndForward = $mbx.DeliverToMailboxAndForward
        $health.LitigationHoldEnabled      = $mbx.LitigationHoldEnabled
        $health.RetentionHoldEnabled       = $mbx.RetentionHoldEnabled
        $health.InPlaceHoldsCount          = if ($mbx.InPlaceHolds) { @($mbx.InPlaceHolds).Count } else { 0 }
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-Mailbox (quota/forwarding/hold/account status)')
    }
    if ($health.ArchiveEnabled) {
        try {
            $astats = Invoke-EXOCommand -Name 'Get-MailboxStatistics' -Parameters @{ Identity = $Smtp; Archive = $true; ErrorAction = 'Stop' }
            $health.ArchiveSize = $astats.TotalItemSize
        } catch {
            $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxStatistics -Archive (health)')
        }
    }
    try {
        $fa = @(Invoke-EXOCommand -Name 'Get-MailboxPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object { $_.User -notlike 'NT AUTHORITY\SELF' -and $_.IsInherited -eq $false -and $_.Deny -eq $false -and ($_.AccessRights -contains 'FullAccess') })
        $health.FullAccessDelegates = @($fa | ForEach-Object { $_.User.ToString() })
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MailboxPermission (Full Access delegates)')
    }
    try {
        $sa = @(Invoke-EXOCommand -Name 'Get-RecipientPermission' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' } |
            Where-Object { $_.Trustee -notlike 'NT AUTHORITY\SELF' -and $_.AccessControlType -eq 'Allow' })
        $health.SendAsDelegates = @($sa | ForEach-Object { $_.Trustee.ToString() })
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-RecipientPermission (Send As delegates)')
    }
    try {
        $rules = Get-InboxRulesForMailbox -Smtp $Smtp
        if ($rules.Error) {
            $health.Errors += $rules.Error
        } else {
            $health.RuleCount = $rules.Rules.Count
            $health.SuspiciousRules = @(
                $rules.Rules | Where-Object {
                    $_.Enabled -and (
                        $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo -or
                        $_.DeleteMessage -or $_.PermanentDelete -or $_.SoftDeleteMessage
                    )
                } | ForEach-Object {
                    $actions = @()
                    if ($_.ForwardTo -or $_.ForwardAsAttachmentTo) { $actions += 'Forward' }
                    if ($_.RedirectTo) { $actions += 'Redirect' }
                    if ($_.DeleteMessage -or $_.PermanentDelete -or $_.SoftDeleteMessage) { $actions += 'Delete' }
                    "$($_.Name)  [$($actions -join ', ')]"
                }
            )
        }
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-InboxRule (rule count / suspicious rule scan)')
    }
    # LAST, AND IN ITS OWN TRY. This is the only field on the form that asks a
    # DIRECTORY rather than Exchange Online, so it is the only one that can fail
    # for a completely unrelated reason. Putting it at the end means an
    # unreachable domain controller costs the licence line and nothing else --
    # the whole Exchange-side snapshot is already built by the time it runs.
    try {
        $health.LicenseState = Get-UserLicenseState -Smtp $Smtp
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-UserLicenseState (licence group membership)')
    }
    try {
        $cas = Invoke-EXOCommand -Name 'Get-CASMailbox' -Parameters @{ Identity = $Smtp; ErrorAction = 'Stop' }
        $health.OWAEnabled        = $cas.OWAEnabled
        $health.ActiveSyncEnabled = $cas.ActiveSyncEnabled
        $health.PopEnabled        = $cas.PopEnabled
        $health.ImapEnabled       = $cas.ImapEnabled
        $health.MAPIEnabled       = $cas.MAPIEnabled
        $health.EwsEnabled        = $cas.EwsEnabled
    } catch {
        $health.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-CASMailbox (protocol access)')
    }
    return [PSCustomObject]$health
}
function Show-MailboxHealthForm {
    # NO -Smtp PARAMETER ANY MORE. The address can change while the window is
    # open, so it lives in $script:Target rather than in a parameter that was
    # fixed the moment the form was built.
    $healthForm = New-Object System.Windows.Forms.Form
    $healthForm.Text = 'Mailbox Health'
    $healthForm.ClientSize = New-Object System.Drawing.Size(624, 770)
    $healthForm.StartPosition = 'Manual'
    $healthForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $txtHealth = New-Object System.Windows.Forms.TextBox
    $txtHealth.Multiline = $true
    $txtHealth.ReadOnly = $true
    $txtHealth.ScrollBars = "Vertical"
    $txtHealth.Location = New-Object System.Drawing.Point(12, 12)
    $txtHealth.Size = New-Object System.Drawing.Size(600, 710)
    $txtHealth.Anchor = 'Top,Left,Right,Bottom'
    $txtHealth.Font = New-Object System.Drawing.Font("Consolas", 9)
    $healthForm.Controls.Add($txtHealth)
    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = "Refresh"
    $btnRefresh.Location = New-Object System.Drawing.Point(12, 730)
    $btnRefresh.Anchor = 'Left,Bottom'
    $btnRefresh.Size = New-Object System.Drawing.Size(100, 28)
    $healthForm.Controls.Add($btnRefresh)
    $btnCopyHealth = New-Object System.Windows.Forms.Button
    $btnCopyHealth.Text = "Copy"
    $btnCopyHealth.Location = New-Object System.Drawing.Point(120, 730)
    $btnCopyHealth.Anchor = 'Left,Bottom'
    $btnCopyHealth.Size = New-Object System.Drawing.Size(100, 28)
    $healthForm.Controls.Add($btnCopyHealth)
    $btnHealthClose = New-Object System.Windows.Forms.Button
    $btnHealthClose.Text = "Close"
    $btnHealthClose.Location = New-Object System.Drawing.Point(228, 730)
    $btnHealthClose.Anchor = 'Left,Bottom'
    $btnHealthClose.Size = New-Object System.Drawing.Size(100, 28)
    $healthForm.Controls.Add($btnHealthClose)
    $copyResetTimer = New-Object System.Windows.Forms.Timer
    $copyResetTimer.Interval = 1200
    $copyResetTimer.Add_Tick({
        $btnCopyHealth.Text = "Copy"
        $copyResetTimer.Stop()
    })
    function Update-HealthFromCurrent {
        if (-not "$script:Target".Trim()) {
            $txtHealth.Text = 'Enter a mailbox address above and press Look Up.'
            $healthForm.Text = 'Mailbox Health'
            return
        }
        $healthForm.Text = "Mailbox Health: $script:Target"
        $btnRefresh.Enabled = $false
        $txtHealth.Text = "Loading mailbox health..."
        $healthForm.Refresh()
        try {
            $h = Get-MailboxHealthSnapshot -Smtp $script:Target
            $sb = New-Object System.Text.StringBuilder
            $issues = @()
            if ($h.AccountDisabled -eq $true) { $issues += "Account sign-in is disabled (AccountDisabled=True) -- unverified as a Get-Mailbox field, confirm before relying on it." }
            if ($h.ForwardingSmtpAddress -or $h.ForwardingAddress) {
                $fwdTarget = if ($h.ForwardingSmtpAddress) { $h.ForwardingSmtpAddress } else { $h.ForwardingAddress }
                $fwdMode = if ($h.DeliverToMailboxAndForward) { "copy kept in mailbox AND forwarded" } else { "forwarded ONLY -- nothing kept in this mailbox" }
                $issues += "Mail forwarding is configured to '$fwdTarget' ($fwdMode)."
            }
            # Only raised when the licence question was actually ANSWERED.
            # Get-UserLicenseState returns Known = $false for a cloud-only
            # tenant or an unreachable DC, and "unknown" must not be reported
            # as "wrong" -- an issue list that cries wolf stops being read.
            if ($h.LicenseState -and $h.LicenseState.Known) {
                if ($h.MailboxType -eq 'Shared' -and $h.LicenseState.Licensed) {
                    $issues += ("Shared mailbox IS licensed ($($h.LicenseState.Groups -join ', ')). " +
                                "Shared mailboxes under 50 GB with no archive do not need one -- check before removing, " +
                                "since a licence here is sometimes deliberate.")
                }
                if ($h.MailboxType -eq 'User' -and -not $h.LicenseState.Licensed) {
                    $issues += ("User mailbox is NOT in any licence group. If the account is still active this will " +
                                "stop mail flow when the grace period ends.")
                }
            }
            if ($h.LitigationHoldEnabled -eq $true) { $issues += "Litigation Hold is enabled." }
            if ($h.RetentionHoldEnabled -eq $true) { $issues += "Retention Hold is enabled." }
            if ($h.InPlaceHoldsCount -gt 0) { $issues += "$($h.InPlaceHoldsCount) In-Place/eDiscovery hold(s) applied." }
            if ($h.SuspiciousRules.Count -gt 0) { $issues += "$($h.SuspiciousRules.Count) enabled Inbox rule(s) forward, redirect, or delete mail -- see 'Suspicious Inbox Rules' below." }
            [void]$sb.AppendLine("Mailbox                : $script:Target")
            [void]$sb.AppendLine("=== Potential Issues (review before escalating) ===")
            if ($issues.Count -eq 0) { [void]$sb.AppendLine("  None of the checks below flagged anything.") }
            else { $issues | ForEach-Object { [void]$sb.AppendLine("  - $_") } }
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Basics ---")
            # FIRST IN THE SECTION, above the quotas and the logon date.
            #
            # Almost every question that reaches this form is answered
            # differently for a shared or resource mailbox than for a user's --
            # "why can't they sign in", "why is there no licence", "why did the
            # quota not go up". Reading that at the bottom is reading it after
            # the wrong assumption has already been made.
            [void]$sb.AppendLine("Mailbox Type            : $($h.MailboxType)  [$($h.RecipientTypeDetails)]")
            if ($h.LicenseState) {
                [void]$sb.AppendLine("Licensed                : $(Format-UserLicenseState -State $h.LicenseState)")
            } else {
                [void]$sb.AppendLine("Licensed                : Not checked -- see field errors below.")
            }
            [void]$sb.AppendLine("Last Logon             : $($h.LastLogonTime)")
            [void]$sb.AppendLine("Item Count              : $($h.ItemCount)")
            [void]$sb.AppendLine("Mailbox Size            : $($h.MailboxSize)")
            [void]$sb.AppendLine("ProhibitSendQuota       : $($h.ProhibitSendQuota)  (compare to Mailbox Size above)")
            [void]$sb.AppendLine("ProhibitSendReceiveQuota: $($h.ProhibitSendReceiveQuota)")
            [void]$sb.AppendLine("Archive Enabled         : $($h.ArchiveEnabled)")
            if ($h.ArchiveEnabled) { [void]$sb.AppendLine("Archive Size            : $($h.ArchiveSize)") }
            [void]$sb.AppendLine("Account Disabled        : $($h.AccountDisabled)  (unverified Get-Mailbox field -- see script header)")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Forwarding & Holds ---")
            [void]$sb.AppendLine("Forwarding SMTP Address : $($h.ForwardingSmtpAddress)")
            [void]$sb.AppendLine("Forwarding Address      : $($h.ForwardingAddress)")
            [void]$sb.AppendLine("Deliver And Forward     : $($h.DeliverToMailboxAndForward)")
            [void]$sb.AppendLine("Litigation Hold         : $($h.LitigationHoldEnabled)")
            [void]$sb.AppendLine("Retention Hold          : $($h.RetentionHoldEnabled)")
            [void]$sb.AppendLine("In-Place Holds          : $($h.InPlaceHoldsCount)")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Delegates ---")
            [void]$sb.AppendLine("Full Access grantees ($($h.FullAccessDelegates.Count)):")
            if ($h.FullAccessDelegates.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
            else { $h.FullAccessDelegates | ForEach-Object { [void]$sb.AppendLine("  $_") } }
            [void]$sb.AppendLine("Send As grantees ($($h.SendAsDelegates.Count)):")
            if ($h.SendAsDelegates.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
            else { $h.SendAsDelegates | ForEach-Object { [void]$sb.AppendLine("  $_") } }
            [void]$sb.AppendLine("  (Distribution List memberships are not scanned here -- use the DL Memberships tool, since a full-tenant scan is too slow to run on every Health click.)")
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Inbox Rules ---")
            [void]$sb.AppendLine("Inbox Rule Count        : $($h.RuleCount)")
            [void]$sb.AppendLine("Suspicious Inbox Rules (enabled, forward/redirect/delete):")
            if ($h.SuspiciousRules.Count -eq 0) { [void]$sb.AppendLine("  (none)") }
            else { $h.SuspiciousRules | ForEach-Object { [void]$sb.AppendLine("  $_") } }
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("--- Protocol Access (enabled/disabled -- not traffic volume) ---")
            [void]$sb.AppendLine("  OWA        : $($h.OWAEnabled)")
            [void]$sb.AppendLine("  ActiveSync : $($h.ActiveSyncEnabled)")
            [void]$sb.AppendLine("  POP        : $($h.PopEnabled)")
            [void]$sb.AppendLine("  IMAP       : $($h.ImapEnabled)")
            [void]$sb.AppendLine("  MAPI       : $($h.MAPIEnabled)")
            [void]$sb.AppendLine("  EWS        : $($h.EwsEnabled)")
            if ($h.Errors.Count -gt 0) {
                [void]$sb.AppendLine("")
                [void]$sb.AppendLine("--- Field errors (partial data above is still valid) ---")
                $h.Errors | ForEach-Object { [void]$sb.AppendLine($_) }
            }
            $txtHealth.Text = $sb.ToString()
        } catch {
            $txtHealth.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Mailbox Health'
        } finally {
            $btnRefresh.Enabled = $true
        }
    }
    Update-HealthFromCurrent
    $btnRefresh.Add_Click({ Update-HealthFromCurrent })
    $btnCopyHealth.Add_Click({
        try {
            [System.Windows.Forms.Clipboard]::SetText($txtHealth.Text)
            $btnCopyHealth.Text = "Copied!"
            $copyResetTimer.Stop()
            $copyResetTimer.Start()
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Copy to clipboard failed: $($_.Exception.Message)",
                "Copy", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        }
    })
    $btnHealthClose.Add_Click({ $healthForm.Close() })
    # BEFORE Add-ToolHelpButton and before the theme: it shifts every existing
    # control down, so anything added after it would be left at the old offset.
    $bar = Add-ToolTargetBar -Form $healthForm -InitialValue $script:Target

    # WIRED HERE, NOT PASSED INTO THE HELPER. Update-HealthFromCurrent is nested
    # in this function, so the handler has to be created in this scope to see it
    # -- exactly like $btnRefresh above.
    $bar.Button.Add_Click({
        $script:Target = Get-TargetSmtp -Text $bar.TextBox.Text
        Update-HealthFromCurrent
    })

    $healthForm.StartPosition = 'Manual'
    Set-ConsoleTheme -Form $healthForm
    Set-FormScreenCorner -Form $healthForm -Corner 'BottomLeft'
        # Consistent '?' in the top-right of every tool. Added after the theme so
    # it is not restyled into a normal button, and BringToFront'd inside the
    # helper because WinForms draws EARLIER-added controls on top.
    Add-ToolHelpButton -Form $healthForm -ScriptPath $PSCommandPath -Title 'Mailbox Health' | Out-Null

    [void]$healthForm.ShowDialog()
}
#endregion

Show-MailboxHealthForm
```


<br>

==============================================================================

# ◆ FILE 8 OF 16 — `MessageTrace.ps1`

> **688 lines.** A clean mid-size tool: grid, filter-without-re-running, export, and an explicit statement of what the tool is NOT for.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```powershell
#Requires -Version 5.1
<#
    MessageTrace.ps1
    Standalone Message Trace / Quarantine Release tool -- Layer 2 tool
    script.

    Author   : Glenn Armbruster
    Revised  : 2026-07-27
#>
# TOOLKIT-HELP-BEGIN
# Follow a message through Exchange Online -- where it went, whether it was
# delivered, and what happened if it was not.
#
# Search by sender, recipient, subject and date range, and filter by status so
# you can show only failures or only quarantined items. Results can be filtered
# further with the search box without re-running the trace.
#
# Quarantined messages can be released from here.
#
# Microsoft only keeps trace data for a limited window -- roughly the last ten
# days for the detailed view. For anything older you need a historical search,
# which Microsoft runs as a background job.
#
# WHEN THIS IS THE WRONG TOOL
#
# Message trace answers "what happened to this message" -- where it went, and
# why it did not arrive. It does not do threat hunting.
#
# For URL and attachment verdicts, campaign views, saved queries or bulk
# remediation, use THREAT EXPLORER instead. The button at the bottom opens it
# in your admin browser profile. It needs Defender for Office 365 Plan 2.
#
# Threat Explorer does not open carrying your filter. Select a row here, use
# Copy Sender or Copy Message-ID, and paste that into the portal's own filter.
#
# Read-only apart from releasing quarantine.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
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

#region Actions - message trace (moved unchanged from Mailbox-Group-Inspector.ps1 lines 3100-3651)
function Invoke-MessageTrace {
    param(
        [string]$SenderAddress,
        [string]$RecipientAddress,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate,
        [string]$Subject,
        [string]$Status
    )
    # 5000 is a deliberate cap, so Exchange warns "There are more results
    # available" whenever a busy window fills it. Silenced because that warning
    # goes to the PowerShell stream, where nobody using this window will see it
    # -- the honest signal is the count on screen, and exactly 5000 means there
    # was more.
    #
    # WarningAction sits NEXT TO ResultSize deliberately. Comment stripping
    # blanks lines rather than removing them, so an explanation in between puts
    # the two eight lines apart and the check that looks +/-3 lines never sees
    # it. A rule you satisfy has to be satisfied where it looks.
    $params = @{
        StartDate   = $StartDate
        EndDate     = $EndDate
        ResultSize    = 5000
        WarningAction = 'SilentlyContinue'
        ErrorAction   = 'Stop'
    }
    if ($SenderAddress)    { $params.SenderAddress    = $SenderAddress }
    if ($RecipientAddress) { $params.RecipientAddress = $RecipientAddress }
    if ($Subject) {
        $params.Subject           = $Subject
        $params.SubjectFilterType = 'Contains'
    }
    if ($Status) { $params.Status = $Status }
    Invoke-EXOCommand -Name 'Get-MessageTraceV2' -Parameters $params
}
function Invoke-MessageTraceDetail {
    param(
        [Parameter(Mandatory)][guid]$MessageTraceId,
        [Parameter(Mandatory)][string]$RecipientAddress,
        [Parameter(Mandatory)][datetime]$StartDate,
        [Parameter(Mandatory)][datetime]$EndDate
    )
    Invoke-EXOCommand -Name 'Get-MessageTraceDetailV2' -Parameters @{
        MessageTraceId   = $MessageTraceId
        RecipientAddress = $RecipientAddress
        StartDate        = $StartDate
        EndDate           = $EndDate
        ErrorAction       = 'Stop'
    }
}
function Find-QuarantineMatch {
    param(
        [string]$MessageId,
        [string]$RecipientAddress,
        [datetime]$StartReceivedDate,
        [datetime]$EndReceivedDate
    )
    if (-not $MessageId) {
        throw "This trace result has no Message-ID value to match against quarantine -- cannot look this message up automatically."
    }
    @(Invoke-EXOCommand -Name 'Get-QuarantineMessage' -Parameters @{
        MessageId         = $MessageId
        RecipientAddress  = $RecipientAddress
        StartReceivedDate = $StartReceivedDate
        EndReceivedDate   = $EndReceivedDate
        ErrorAction       = 'Stop'
    })
}
function Invoke-QuarantineRelease {
    param([Parameter(Mandatory)]$QuarantineIdentity)
    Invoke-EXOCommand -Name 'Release-QuarantineMessage' -Parameters @{
        Identity     = $QuarantineIdentity
        ReleaseToAll = $true
        Confirm      = $false
        ErrorAction  = 'Stop'
    } | Out-Null
}
$script:TraceTimeZoneIds = [ordered]@{
    'Eastern'  = 'Eastern Standard Time'
    'Central'  = 'Central Standard Time'
    'Mountain' = 'Mountain Standard Time'
    'Pacific'  = 'Pacific Standard Time'
    'UTC'      = 'UTC'
}
function Get-SelectedTraceTimeZone {
    param([string]$Label)
    $id = $script:TraceTimeZoneIds[$Label]
    if (-not $id) { $id = 'UTC' }
    try {
        return [System.TimeZoneInfo]::FindSystemTimeZoneById($id)
    } catch {
        return [System.TimeZoneInfo]::Utc
    }
}
function ConvertTo-TraceDisplayTime {
    param($UtcValue, [System.TimeZoneInfo]$TimeZone)
    if (-not $UtcValue) { return $UtcValue }
    try {
        $forced = [DateTime]::SpecifyKind([DateTime]$UtcValue, [DateTimeKind]::Utc)
        return [System.TimeZoneInfo]::ConvertTimeFromUtc($forced, $TimeZone)
    } catch {
        return $UtcValue
    }
}
function Format-TraceDetail {
    param($Details, [System.TimeZoneInfo]$TimeZone)
    $sb = New-Object System.Text.StringBuilder
    $rows = @($Details)
    if ($rows.Count -eq 0) {
        [void]$sb.AppendLine("(no detail records returned)")
        return $sb.ToString()
    }
    foreach ($d in $rows) {
        $displayDate = if ($TimeZone) { ConvertTo-TraceDisplayTime -UtcValue $d.Date -TimeZone $TimeZone } else { $d.Date }
        [void]$sb.AppendLine("$displayDate  [$($d.Event)]  $($d.Action)")
        if ($d.Detail) { [void]$sb.AppendLine("    $($d.Detail)") }
    }
    return $sb.ToString()
}
function Show-MessageTraceForm {
    param([string]$SmtpAddress, [string]$Environment)
    $traceForm = New-Object System.Windows.Forms.Form
    $traceForm.Text = if ($SmtpAddress) { "Message Trace (from: $SmtpAddress)" } else { "Message Trace" }
    $traceForm.Size = New-Object System.Drawing.Size(1014, 754)
    $traceForm.StartPosition = 'Manual'
    $traceForm.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)
    $lblFrom = New-Object System.Windows.Forms.Label
    $lblFrom.Text = "From:"
    $lblFrom.Location = New-Object System.Drawing.Point(12, 16)
    $lblFrom.AutoSize = $true
    $traceForm.Controls.Add($lblFrom)
    $txtFrom = New-Object System.Windows.Forms.TextBox
    $txtFrom.Location = New-Object System.Drawing.Point(55, 13)
    $txtFrom.Size = New-Object System.Drawing.Size(190, 20)
    $txtFrom.Font = New-Object System.Drawing.Font("Consolas", 9)
    if ($SmtpAddress) { $txtFrom.Text = $SmtpAddress }
    $traceForm.Controls.Add($txtFrom)
    $lblTo = New-Object System.Windows.Forms.Label
    $lblTo.Text = "To:"
    $lblTo.Location = New-Object System.Drawing.Point(12, 40)
    $lblTo.AutoSize = $true
    $traceForm.Controls.Add($lblTo)
    $txtTo = New-Object System.Windows.Forms.TextBox
    $txtTo.Location = New-Object System.Drawing.Point(55, 37)
    $txtTo.Size = New-Object System.Drawing.Size(190, 20)
    $txtTo.Font = New-Object System.Drawing.Font("Consolas", 9)
    $traceForm.Controls.Add($txtTo)
    $lblTraceNote = New-Object System.Windows.Forms.Label
    $lblTraceNote.Text = "Read-only -- no objects are modified. Leave From/To blank to trace all mail in the date range. Exchange Online only (Get-MessageTraceV2 / Get-MessageTraceDetailV2), routed through Invoke-EXOCommand."
    $lblTraceNote.Location = New-Object System.Drawing.Point(264, 12)
    $lblTraceNote.Size = New-Object System.Drawing.Size(488, 46)
    $traceForm.Controls.Add($lblTraceNote)
    $lblTZ = New-Object System.Windows.Forms.Label
    $lblTZ.Text = "Show time in:"
    $lblTZ.Location = New-Object System.Drawing.Point(766, 14)
    $lblTZ.AutoSize = $true
    $traceForm.Controls.Add($lblTZ)
    $cboTimeZone = New-Object System.Windows.Forms.ComboBox
    $cboTimeZone.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('Eastern', 'Central', 'Mountain', 'Pacific', 'UTC') | ForEach-Object { [void]$cboTimeZone.Items.Add($_) }
    $cboTimeZone.SelectedItem = 'Central'
    $cboTimeZone.Location = New-Object System.Drawing.Point(766, 34)
    $cboTimeZone.Size = New-Object System.Drawing.Size(200, 20)
    $traceForm.Controls.Add($cboTimeZone)
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = "Status:"
    $lblStatus.Location = New-Object System.Drawing.Point(510, 72)
    $lblStatus.AutoSize = $true
    $traceForm.Controls.Add($lblStatus)
    $cboStatus = New-Object System.Windows.Forms.ComboBox
    $cboStatus.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    @('All Messages', 'Delivered', 'Failed', 'Quarantined', 'FilteredAsSpam', 'Pending', 'Expanded', 'GettingStatus') |
        ForEach-Object { [void]$cboStatus.Items.Add($_) }
    $cboStatus.SelectedIndex = 0
    $cboStatus.Location = New-Object System.Drawing.Point(562, 69)
    $cboStatus.Size = New-Object System.Drawing.Size(180, 20)
    $traceForm.Controls.Add($cboStatus)
    $lblSubject = New-Object System.Windows.Forms.Label
    $lblSubject.Text = "Subject contains (optional):"
    $lblSubject.Location = New-Object System.Drawing.Point(12, 100)
    $lblSubject.AutoSize = $true
    $traceForm.Controls.Add($lblSubject)
    $txtSubject = New-Object System.Windows.Forms.TextBox
    $txtSubject.Location = New-Object System.Drawing.Point(270, 97)
    $txtSubject.Size = New-Object System.Drawing.Size(220, 20)
    $txtSubject.Font = New-Object System.Drawing.Font("Consolas", 9)
    $traceForm.Controls.Add($txtSubject)
    $lblStart = New-Object System.Windows.Forms.Label
    $lblStart.Text = "Search From:"
    $lblStart.Location = New-Object System.Drawing.Point(12, 129)
    $lblStart.AutoSize = $true
    $traceForm.Controls.Add($lblStart)
    $dtpStart = New-Object System.Windows.Forms.DateTimePicker
    $dtpStart.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $dtpStart.Location = New-Object System.Drawing.Point(100, 125)
    $dtpStart.Size = New-Object System.Drawing.Size(120, 20)
    $dtpStart.MinDate = (Get-Date).Date.AddDays(-30)
    $dtpStart.MaxDate = (Get-Date).Date
    $dtpStart.Value   = (Get-Date).Date.AddDays(-7)
    $traceForm.Controls.Add($dtpStart)
    $lblEnd = New-Object System.Windows.Forms.Label
    $lblEnd.Text = "Through:"
    $lblEnd.Location = New-Object System.Drawing.Point(240, 129)
    $lblEnd.AutoSize = $true
    $traceForm.Controls.Add($lblEnd)
    $dtpEnd = New-Object System.Windows.Forms.DateTimePicker
    $dtpEnd.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $dtpEnd.Location = New-Object System.Drawing.Point(300, 125)
    $dtpEnd.Size = New-Object System.Drawing.Size(120, 20)
    $traceForm.Controls.Add($dtpEnd)
    $dtpStart.Add_ValueChanged({
        $maxEnd = $dtpStart.Value.Date.AddDays(9)
        $today  = (Get-Date).Date
        if ($maxEnd -gt $today) { $maxEnd = $today }
        $dtpEnd.MinDate = $dtpStart.Value.Date
        $dtpEnd.MaxDate = $maxEnd
        if ($dtpEnd.Value -lt $dtpStart.Value.Date) { $dtpEnd.Value = $dtpStart.Value.Date }
        if ($dtpEnd.Value -gt $maxEnd) { $dtpEnd.Value = $maxEnd }
    })
    $dtpEnd.MinDate = $dtpStart.Value.Date
    $initialMaxEnd = $dtpStart.Value.Date.AddDays(9)
    if ($initialMaxEnd -gt (Get-Date).Date) { $initialMaxEnd = (Get-Date).Date }
    $dtpEnd.MaxDate = $initialMaxEnd
    $dtpEnd.Value   = $initialMaxEnd
    $btnTraceRun = New-Object System.Windows.Forms.Button
    $btnTraceRun.Text = "Trace"
    $btnTraceRun.Location = New-Object System.Drawing.Point(440, 123)
    $btnTraceRun.Size = New-Object System.Drawing.Size(90, 26)
    $traceForm.Controls.Add($btnTraceRun)
    $lblTraceStatus = New-Object System.Windows.Forms.Label
    $lblTraceStatus.Text = ""
    $lblTraceStatus.Location = New-Object System.Drawing.Point(12, 156)
    $lblTraceStatus.Size = New-Object System.Drawing.Size(974, 40)
    $traceForm.Controls.Add($lblTraceStatus)
    $lblSearchGrid = New-Object System.Windows.Forms.Label
    $lblSearchGrid.Text = "Filter loaded results:"
    $lblSearchGrid.Location = New-Object System.Drawing.Point(12, 204)
    $lblSearchGrid.AutoSize = $true
    $traceForm.Controls.Add($lblSearchGrid)
    $txtSearchGrid = New-Object System.Windows.Forms.TextBox
    $txtSearchGrid.Location = New-Object System.Drawing.Point(150, 200)
    $txtSearchGrid.Size = New-Object System.Drawing.Size(280, 22)
    $txtSearchGrid.Font = New-Object System.Drawing.Font("Consolas", 9)
    $traceForm.Controls.Add($txtSearchGrid)
    $btnSearchGrid = New-Object System.Windows.Forms.Button
    $btnSearchGrid.Text = "Search"
    $btnSearchGrid.Location = New-Object System.Drawing.Point(440, 198)
    $btnSearchGrid.Size = New-Object System.Drawing.Size(80, 26)
    $traceForm.Controls.Add($btnSearchGrid)
    $btnClearSearchGrid = New-Object System.Windows.Forms.Button
    $btnClearSearchGrid.Text = "Clear Filter"
    $btnClearSearchGrid.Location = New-Object System.Drawing.Point(526, 198)
    $btnClearSearchGrid.Size = New-Object System.Drawing.Size(90, 26)
    $traceForm.Controls.Add($btnClearSearchGrid)
    $grdTrace = New-Object System.Windows.Forms.DataGridView
    $grdTrace.Location = New-Object System.Drawing.Point(12, 234)
    $grdTrace.Size = New-Object System.Drawing.Size(974, 260)
    $grdTrace.ReadOnly = $true
    $grdTrace.AllowUserToAddRows = $false
    $grdTrace.AllowUserToDeleteRows = $false
    $grdTrace.AllowUserToResizeRows = $false
    $grdTrace.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grdTrace.MultiSelect = $false
    $grdTrace.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
    $grdTrace.Font = New-Object System.Drawing.Font("Consolas", 9)
    [void]$grdTrace.Columns.Add('Received', 'Received')
    [void]$grdTrace.Columns.Add('Sender', 'Sender Address')
    [void]$grdTrace.Columns.Add('Recipient', 'Recipient Address')
    [void]$grdTrace.Columns.Add('Subject', 'Subject')
    [void]$grdTrace.Columns.Add('Status', 'Status')
    [void]$grdTrace.Columns.Add('Key', 'Key')
    $grdTrace.Columns['Received'].FillWeight  = 15
    $grdTrace.Columns['Sender'].FillWeight    = 20
    $grdTrace.Columns['Recipient'].FillWeight = 20
    $grdTrace.Columns['Subject'].FillWeight   = 30
    $grdTrace.Columns['Status'].FillWeight    = 15
    $grdTrace.Columns['Received'].ValueType   = [DateTime]
    $grdTrace.Columns['Key'].Visible          = $false
    foreach ($col in $grdTrace.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic }
    $traceForm.Controls.Add($grdTrace)
    $lblDetail = New-Object System.Windows.Forms.Label
    $lblDetail.Text = "Detail (double-click a row above to load it here):"
    $lblDetail.Location = New-Object System.Drawing.Point(12, 502)
    $lblDetail.AutoSize = $true
    $traceForm.Controls.Add($lblDetail)
    $txtDetail = New-Object System.Windows.Forms.TextBox
    $txtDetail.Multiline = $true
    $txtDetail.ReadOnly = $true
    $txtDetail.ScrollBars = "Vertical"
    $txtDetail.BackColor = [System.Drawing.Color]::WhiteSmoke
    $txtDetail.Location = New-Object System.Drawing.Point(12, 522)
    $txtDetail.Size = New-Object System.Drawing.Size(974, 130)
    $txtDetail.Font = New-Object System.Drawing.Font("Consolas", 9)
    $traceForm.Controls.Add($txtDetail)
    $btnRelease = New-Object System.Windows.Forms.Button
    $btnRelease.Text = "Release from Quarantine"
    $btnRelease.Location = New-Object System.Drawing.Point(12, 662)
    $btnRelease.Size = New-Object System.Drawing.Size(160, 28)
    $btnRelease.Enabled = $false
    $traceForm.Controls.Add($btnRelease)
    $btnExportCsv = New-Object System.Windows.Forms.Button
    $btnExportCsv.Text = "Export to CSV"
    $btnExportCsv.Location = New-Object System.Drawing.Point(180, 662)
    $btnExportCsv.Size = New-Object System.Drawing.Size(110, 28)
    $btnExportCsv.Enabled = $false
    $traceForm.Controls.Add($btnExportCsv)
    # --- Threat Explorer hand-off -----------------------------------------
    #
    # Message trace answers "what happened to this message". It does not do
    # threat hunting -- URL and attachment verdicts, campaign views, saved
    # queries, bulk remediation. Those live in Threat Explorer, which needs
    # Defender for Office 365 Plan 2.
    #
    # Ported from the paused browser suite, which explained the distinction on
    # screen. Knowing which of the two tools answers your question is most of
    # the work, and it was the one useful thing that version had over this one.
    $btnExplorer = New-Object System.Windows.Forms.Button
    $btnExplorer.Text = "Threat Explorer"
    $btnExplorer.Location = New-Object System.Drawing.Point(300, 662)
    $btnExplorer.Size = New-Object System.Drawing.Size(150, 28)
    $traceForm.Controls.Add($btnExplorer)

    # Copy, rather than a pre-filled URL, and this is deliberate.
    #
    # The browser suite had a "Prepopulated" flag and a message reading "Threat
    # Explorer opened with your filter applied". It could never fire: the URL
    # template was the bare portal address with no substitution tokens, and the
    # function that would have replaced it was exported and never called. The
    # copy buttons were what actually worked there, so they are what is ported.
    #
    # Threat Explorer v3's filter parameters are not documented in a form worth
    # depending on. Copy and paste is unglamorous and correct.
    $btnCopySender = New-Object System.Windows.Forms.Button
    $btnCopySender.Text = "Copy Sender"
    $btnCopySender.Location = New-Object System.Drawing.Point(458, 662)
    $btnCopySender.Size = New-Object System.Drawing.Size(110, 28)
    $btnCopySender.Enabled = $false
    $traceForm.Controls.Add($btnCopySender)

    $btnCopyMsgId = New-Object System.Windows.Forms.Button
    $btnCopyMsgId.Text = "Copy Message-ID"
    $btnCopyMsgId.Location = New-Object System.Drawing.Point(576, 662)
    $btnCopyMsgId.Size = New-Object System.Drawing.Size(130, 28)
    $btnCopyMsgId.Enabled = $false
    $traceForm.Controls.Add($btnCopyMsgId)

    $lblCopyNote = New-Object System.Windows.Forms.Label
    $lblCopyNote.Tag = 'NoTheme'
    $lblCopyNote.Location = New-Object System.Drawing.Point(714, 666)
    $lblCopyNote.Size = New-Object System.Drawing.Size(166, 20)
    $lblCopyNote.AutoEllipsis = $true
    $lblCopyNote.ForeColor = [System.Drawing.Color]::DimGray
    $traceForm.Controls.Add($lblCopyNote)

    $tipTrace = New-Object System.Windows.Forms.ToolTip
    $tipTrace.SetToolTip($btnExplorer, @"
Opens Threat Explorer in the Microsoft 365 Defender portal.

Use it instead of message trace for: URL and attachment verdicts,
campaign views, saved queries, and bulk remediation.

Needs Defender for Office 365 Plan 2.

It does NOT open pre-filtered -- copy the sender or Message-ID
first and paste it into the portal's own filter.
"@)
    $tipTrace.SetToolTip($btnCopySender, 'Copy the selected row''s sender, to paste into a portal filter.')
    $tipTrace.SetToolTip($btnCopyMsgId,  'Copy the selected row''s Message-ID, to paste into a portal filter.')

    $btnTraceClose = New-Object System.Windows.Forms.Button
    $btnTraceClose.Text = "Close"
    $btnTraceClose.Location = New-Object System.Drawing.Point(886, 662)
    $btnTraceClose.Size = New-Object System.Drawing.Size(100, 28)
    $traceForm.Controls.Add($btnTraceClose)
    $script:CurrentTraceResults        = @()
    $script:CurrentTraceDisplayResults = @()
    $script:CurrentTraceDetailRaw       = $null
    $script:TraceSearchStart            = $null
    $script:TraceSearchEnd              = $null
    function Get-FilteredTraceResults {
        param([string]$SearchText)
        if ([string]::IsNullOrWhiteSpace($SearchText)) { return $script:CurrentTraceResults }
        $needle = "*$($SearchText.Trim())*"
        $script:CurrentTraceResults | Where-Object {
            ($_.SenderAddress    -and $_.SenderAddress    -like $needle) -or
            ($_.RecipientAddress -and $_.RecipientAddress -like $needle) -or
            ($_.Subject          -and $_.Subject          -like $needle) -or
            ($_.Status           -and $_.Status           -like $needle)
        }
    }
    function Update-TraceGridDisplay {
        $tz = Get-SelectedTraceTimeZone -Label $cboTimeZone.SelectedItem
        $grdTrace.Columns['Received'].HeaderText = "Received ($($cboTimeZone.SelectedItem))"
        $grdTrace.Rows.Clear()
        foreach ($r in $script:CurrentTraceDisplayResults) {
            $rowKey = "$($r.MessageTraceId)|$($r.RecipientAddress)"
            $displayReceived = ConvertTo-TraceDisplayTime -UtcValue $r.Received -TimeZone $tz
            [void]$grdTrace.Rows.Add($displayReceived, $r.SenderAddress, $r.RecipientAddress, $r.Subject, $r.Status, $rowKey)
        }
    }
    $cboTimeZone.Add_SelectedIndexChanged({
        Update-TraceGridDisplay
        if ($script:CurrentTraceDetailRaw) {
            $tz = Get-SelectedTraceTimeZone -Label $cboTimeZone.SelectedItem
            $txtDetail.Text = Format-TraceDetail -Details $script:CurrentTraceDetailRaw -TimeZone $tz
        }
    })
    $btnTraceRun.Add_Click({
        $btnTraceRun.Enabled = $false
        $btnRelease.Enabled = $false
        $btnExportCsv.Enabled = $false
        $txtDetail.Text = ""
        $txtSearchGrid.Text = ""
        $grdTrace.Rows.Clear()
        $script:CurrentTraceResults = @()
        $script:CurrentTraceDisplayResults = @()
        $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
        $lblTraceStatus.Text = "Running trace..."
        try {
            $fromAddr = $txtFrom.Text.Trim()
            $toAddr   = $txtTo.Text.Trim()
            $subject = $txtSubject.Text.Trim()
            $sd = $dtpStart.Value.Date
            $ed = $dtpEnd.Value.Date.AddDays(1).AddSeconds(-1)
            $traceParams = @{ StartDate = $sd; EndDate = $ed }
            if ($fromAddr) { $traceParams.SenderAddress    = $fromAddr }
            if ($toAddr)   { $traceParams.RecipientAddress = $toAddr }
            if ($subject) { $traceParams.Subject = $subject }
            if ($cboStatus.SelectedItem -and $cboStatus.SelectedItem -ne 'All Messages') {
                $traceParams.Status = $cboStatus.SelectedItem
            }
            $results = @(Invoke-MessageTrace @traceParams)
            $script:CurrentTraceResults = $results
            $script:CurrentTraceDisplayResults = $results
            $script:CurrentTraceDetailRaw = $null
            $script:TraceSearchStart = $sd
            $script:TraceSearchEnd   = $ed
            Update-TraceGridDisplay
            $btnExportCsv.Enabled = ($results.Count -gt 0)
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
            $statusNote = if ($traceParams.Status) { "  Status: $($traceParams.Status)." } else { "" }
            $lblTraceStatus.Text = "$($results.Count) result(s). $sd through $ed.$statusNote"
        } catch {
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Red
            $lblTraceStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MessageTraceV2'
        } finally {
            $btnTraceRun.Enabled = $true
        }
    })
    $btnSearchGrid.Add_Click({
        $script:CurrentTraceDisplayResults = @(Get-FilteredTraceResults -SearchText $txtSearchGrid.Text)
        Update-TraceGridDisplay
        $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
        if ([string]::IsNullOrWhiteSpace($txtSearchGrid.Text)) {
            $lblTraceStatus.Text = "$($script:CurrentTraceResults.Count) result(s)."
        } else {
            $lblTraceStatus.Text = "$($script:CurrentTraceDisplayResults.Count) of $($script:CurrentTraceResults.Count) result(s) match '$($txtSearchGrid.Text.Trim())'."
        }
    })
    $txtSearchGrid.Add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $btnSearchGrid.PerformClick()
            $e.SuppressKeyPress = $true
        }
    })
    $btnClearSearchGrid.Add_Click({
        $txtSearchGrid.Text = ''
        $script:CurrentTraceDisplayResults = $script:CurrentTraceResults
        Update-TraceGridDisplay
        $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
        $lblTraceStatus.Text = "$($script:CurrentTraceResults.Count) result(s)."
    })
    function Find-TraceResultByKey {
        param([string]$RowKey)
        $script:CurrentTraceResults | Where-Object { "$($_.MessageTraceId)|$($_.RecipientAddress)" -eq $RowKey } | Select-Object -First 1
    }
    $grdTrace.Add_SelectionChanged({
        $has = ($grdTrace.SelectedRows.Count -gt 0)
        $btnRelease.Enabled   = $has
        $btnCopySender.Enabled = $has
        # Message-ID is not always present on a trace row, and a button that
        # copies an empty string is worse than one that is greyed: the first
        # looks like it worked.
        $btnCopyMsgId.Enabled = $false
        if ($has) {
            $rec = Find-TraceResultByKey -RowKey $grdTrace.SelectedRows[0].Cells['Key'].Value
            if ($rec -and "$($rec.MessageId)".Trim()) { $btnCopyMsgId.Enabled = $true }
        }
        $lblCopyNote.Text = ''
    })

    function Copy-TraceValue {
        param([string]$Value, [string]$Label)
        if (-not "$Value".Trim()) {
            $lblCopyNote.Text = "No $Label on that row."
            return
        }
        try {
            [System.Windows.Forms.Clipboard]::SetText("$Value")
            $lblCopyNote.Text = "$Label copied."
        } catch {
            # Clipboard access fails if another process holds it open. Say so
            # rather than leaving the user to wonder why nothing pasted.
            $lblCopyNote.Text = "Copy failed -- clipboard busy."
        }
    }

    $btnCopySender.Add_Click({
        if ($grdTrace.SelectedRows.Count -eq 0) { return }
        $rec = Find-TraceResultByKey -RowKey $grdTrace.SelectedRows[0].Cells['Key'].Value
        if (-not $rec) { $lblCopyNote.Text = 'Row no longer matches -- re-run Trace.'; return }
        Copy-TraceValue -Value $rec.SenderAddress -Label 'Sender'
    })

    $btnCopyMsgId.Add_Click({
        if ($grdTrace.SelectedRows.Count -eq 0) { return }
        $rec = Find-TraceResultByKey -RowKey $grdTrace.SelectedRows[0].Cells['Key'].Value
        if (-not $rec) { $lblCopyNote.Text = 'Row no longer matches -- re-run Trace.'; return }
        Copy-TraceValue -Value $rec.MessageId -Label 'Message-ID'
    })

    $btnExplorer.Add_Click({
        # Admin browser profile, same as every other admin URL in the toolkit.
        # Threat Explorer needs the elevated account; opening it in a personal
        # profile produces a sign-in prompt that looks like a permissions fault.
        $url = 'https://security.microsoft.com/threatexplorerv3'
        try {
            $st = Start-UrlInBrowserProfile -Url $url -Purpose 'Admin'
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
            $lblTraceStatus.Text = "Threat Explorer opened ($st). It does not carry your filter -- use the copy buttons and paste into its own filter."
        } catch {
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Red
            $lblTraceStatus.Text = "Could not open Threat Explorer: $($_.Exception.Message)"
        }
    })
    $grdTrace.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $rowKey = $grdTrace.Rows[$e.RowIndex].Cells['Key'].Value
        $rec = Find-TraceResultByKey -RowKey $rowKey
        if (-not $rec) {
            $txtDetail.Text = "Could not match that row back to a trace result -- run Trace again."
            return
        }
        $txtDetail.Text = "Loading..."
        try {
            $details = Invoke-MessageTraceDetail -MessageTraceId $rec.MessageTraceId -RecipientAddress $rec.RecipientAddress `
                -StartDate $script:TraceSearchStart -EndDate $script:TraceSearchEnd
            $script:CurrentTraceDetailRaw = $details
            $tz = Get-SelectedTraceTimeZone -Label $cboTimeZone.SelectedItem
            $txtDetail.Text = Format-TraceDetail -Details $details -TimeZone $tz
        } catch {
            $script:CurrentTraceDetailRaw = $null
            $txtDetail.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Get-MessageTraceDetailV2'
        }
    })
    $btnRelease.Add_Click({
        if ($grdTrace.SelectedRows.Count -eq 0) { return }
        $rowKey = $grdTrace.SelectedRows[0].Cells['Key'].Value
        $rec = Find-TraceResultByKey -RowKey $rowKey
        if (-not $rec) {
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Red
            $lblTraceStatus.Text = "Could not match that row back to a trace result -- run Trace again."
            return
        }
        $confirmSearch = [System.Windows.Forms.MessageBox]::Show(
            "Look up quarantine entries for this message?`nSender: $($rec.SenderAddress)`nRecipient: $($rec.RecipientAddress)`nSubject: $($rec.Subject)",
            "Confirm Quarantine Lookup", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirmSearch -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        $btnRelease.Enabled = $false
        try {
            $matches = @(Find-QuarantineMatch -MessageId $rec.MessageId -RecipientAddress $rec.RecipientAddress `
                -StartReceivedDate $script:TraceSearchStart -EndReceivedDate $script:TraceSearchEnd)
            if ($matches.Count -eq 0) {
                $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
                $lblTraceStatus.Text = "No quarantine entry found for this message (it may not have been quarantined, may already be released, or may have expired -- quarantine data is only retained for a limited time)."
                return
            }
            foreach ($qmsg in $matches) {
                $confirmRelease = [System.Windows.Forms.MessageBox]::Show(
                    "Quarantine match found:`nType: $($qmsg.Type)   Policy: $($qmsg.PolicyName)`nReceived: $($qmsg.ReceivedTime)   Release status: $($qmsg.ReleaseStatus)`n`nRelease this message to ALL original recipients?",
                    "Confirm Release", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($confirmRelease -ne [System.Windows.Forms.DialogResult]::Yes) { continue }
                Invoke-QuarantineRelease -QuarantineIdentity $qmsg.Identity
                Write-CompletedLog -Task "Released QuarantineMessage $($rec.RecipientAddress) ($Environment; sender: $($rec.SenderAddress); subject: $($rec.Subject))"
            }
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
            $lblTraceStatus.Text = "$($matches.Count) quarantine match(es) processed."
        } catch {
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Red
            $lblTraceStatus.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Quarantine release'
        } finally {
            $btnRelease.Enabled = $true
        }
    })
    $btnExportCsv.Add_Click({
        if ($script:CurrentTraceDisplayResults.Count -eq 0) { $lblTraceStatus.Text = "No results to export."; return }
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV files (*.csv)|*.csv"
        $filePart = if ($SmtpAddress) { ($SmtpAddress -replace '[^a-zA-Z0-9]','_') } else { 'All' }
        $sfd.FileName = "MessageTrace_${filePart}_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
        if ($sfd.ShowDialog($traceForm) -ne [System.Windows.Forms.DialogResult]::OK) { return }
        try {
            $tz = Get-SelectedTraceTimeZone -Label $cboTimeZone.SelectedItem
            $script:CurrentTraceDisplayResults |
                Select-Object @{Name = "Received ($($cboTimeZone.SelectedItem))"; Expression = { ConvertTo-TraceDisplayTime -UtcValue $_.Received -TimeZone $tz }},
                              SenderAddress, RecipientAddress, Subject, Status, MessageTraceId |
                Export-Csv -LiteralPath $sfd.FileName -NoTypeInformation -Encoding UTF8
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Black
            $lblTraceStatus.Text = "Exported $($script:CurrentTraceDisplayResults.Count) row(s) to $($sfd.FileName)."
        } catch {
            $lblTraceStatus.ForeColor = [System.Drawing.Color]::Red
            $lblTraceStatus.Text = "Export failed: $($_.Exception.Message)"
        }
    })
    $btnTraceClose.Add_Click({ $traceForm.Close() })
    $traceForm.StartPosition = 'Manual'
    Set-ConsoleTheme -Form $traceForm
    Set-FormScreenCorner -Form $traceForm -Corner 'BottomLeft'
    # Consistent '?' in the top-right of every tool. Added after the theme so
    # it is not restyled, and BringToFront'd inside the helper because
    # WinForms draws EARLIER-added controls on top.
    Add-ToolHelpButton -Form $traceForm -ScriptPath $PSCommandPath -Title 'Message Trace' | Out-Null

    [void]$traceForm.ShowDialog()
}
#endregion

Show-MessageTraceForm -SmtpAddress $SmtpAddress -Environment $Environment
```


<br>

==============================================================================

# ◆ FILE 9 OF 16 — `Unified-Lookup.ps1`

> **2756 lines.** The most complex read tool. Two directories searched in one pass, results classified and routed, a capped result set filtered. NOTE: much of its dual-directory machinery is NOT needed for a single-domain tenant.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```powershell
#Requires -Version 5.1
# TOOLKIT-VARIANT: Dual
#
# THIS TOOL DOES NOT DECLARE THE OwnProcess MARKER, AND THAT IS THE WHOLE TRICK.
#
# DO NOT WRITE THAT MARKER'S FULL NAME ANYWHERE IN THE FIRST 40 LINES OF THIS
# FILE, NOT EVEN TO SAY IT IS ABSENT. THIS EXACT MISTAKE SHIPPED ONCE.
#
# Test-ToolDeclaresOwnProcess in AdminConsole.ps1 reads the first 40 lines and
# regex-matches the marker text. It has no idea whether the line is a
# declaration or a sentence about one. The original version of this header said
# "NO 'TOOLKIT-REQUIRES<colon> OwnProcess', AND THAT IS THE WHOLE TRICK" -- and
# the console dutifully matched it and launched this tool in a DETACHED process
# with no Exchange Online session.
#
# The symptom was a banner reading "ON-PREMISES ONLY -- Exchange Online is not
# connected", and it took a rewrite of the whole connection model before anyone
# looked at line 4. Group-Manager's Test-ToolNeedsOwnProcess docstring warns about
# precisely this and says it reads the HEADER ONLY for the same reason.
#
# Every other hybrid tool here needs its own process because it calls Exchange
# cmdlets UNQUALIFIED -- Get-Recipient resolves to whichever of the two Exchange
# modules was imported last, so on-premises and Exchange Online cannot share a
# window. This tool calls NOTHING unqualified. Every on-premises call goes
# through Invoke-OnPremCommand and every cloud call through Invoke-EXOCommand,
# each of which resolves the name against a specific module.
#
# That is what makes one window able to search both worlds at once, and it only
# became possible when Invoke-OnPremCommand was written -- its own docstring
# calls it "the counterpart nobody wrote, which is why the hybrid scripts
# qualified their cloud calls and left their on-premises ones bare."
<#
    Unified-Lookup.ps1
    One search box. Every object, wherever it lives.

    AD-Lookup answers "what is this, on-premises". Group-Manager answers "which
    side of the hybrid owns this group". Between them sits the question neither
    asks: find the thing, whatever it is and wherever it is, without knowing in
    advance which tool to open.

    Reasoning: ScriptNotes/Unified-Lookup.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-20
#>
# TOOLKIT-HELP-BEGIN
# Type part of a name. Get back everything that matches, from everywhere.
#
# FOUR SOURCES, SEARCHED TOGETHER:
#
#   On-premises AD         ALWAYS. Distribution lists, mail-enabled security
#                          groups, shared / room / equipment mailboxes, remote
#                          mailbox stubs, mail contacts -- read straight from
#                          Active Directory, which needs no Exchange connection
#                          of any kind. Tick the box to add objects with no
#                          address at all (plain security groups).
#   On-premises Exchange   when the Exchange Management Shell is loaded. Adds
#                          Exchange's own recipient type, which knows about
#                          room lists and arbitration mailboxes that no AD
#                          attribute distinguishes.
#   Exchange Online        cloud distribution lists, Microsoft 365 groups,
#                          shared / room mailboxes, contacts, room lists
#   Soft-deleted           mailboxes deleted in the cloud but still recoverable
#
# CONNECTIONS ARE THE CONSOLE'S JOB, NOT THIS TOOL'S. This runs in the console's
# own process, so whatever the console is connected to, this is connected to. If
# the line at the top says Exchange Online is missing, use Connect ExO in the
# console and search again -- no need to close this window.
#
# THE SOURCE COLUMN IS THE ANSWER TO A QUESTION YOU ASK CONSTANTLY:
#
#   On-prem        exists only on-premises -- not synced, or not yet
#   Cloud          created in the cloud; there is no on-premises object at all
#   Synced         found in BOTH. The on-premises object is the authority and
#                  the cloud copy follows it. Editing the cloud side directly
#                  is the change that silently reverts on the next sync.
#   Soft-deleted   deleted, still recoverable, still holding its address
#
# ONE ROW PER OBJECT. A synced object is found twice and shown once, merged on
# its primary address, because two rows for one thing is not an answer -- it is
# the raw material of one.
#
# WHAT IT DOES NOT SEARCH: Entra security groups that are not mail-enabled. No
# Exchange cmdlet returns them and Graph is not available here yet. A cloud-only
# security group will not appear -- that is a gap, not a result.
#
# THIS TOOL WRITES. It did not used to, and the difference matters.
#
#   Attributes   WHAT THE OBJECT IS -- display name, alias, primary address,
#                hide-from-GAL, the fifteen custom attributes. For a cloud-only
#                object it edits here, with Set-Mailbox / Set-DistributionGroup /
#                Set-MailContact / Set-MailUser / Set-UnifiedGroup as the type
#                requires. For anything with a distinguished name it opens
#                AD-Lookup, because Active Directory is the authority.
#   Mailbox      HOW THE MAILBOX BEHAVES -- forwarding, and whether sending as
#                settings     the mailbox leaves a copy in ITS Sent Items.
#                Always Exchange Online, because that is where a mailbox lives.
#   Open         WHAT THE OBJECT DOES -- a group opens its membership, a mailbox
#                opens its permissions.
#
# WHY ATTRIBUTES AND MAILBOX SETTINGS ARE TWO BUTTONS. A RemoteSharedMailbox has
# its directory object on-premises and its MAILBOX in Exchange Online. One row,
# two homes. Folding both into one button left one of them unreachable -- and it
# did: send-as copies could not be set on any synced shared mailbox at all.
#
# THE SOURCE COLUMN DECIDES WHICH SIDE IS WRITTEN, and that is the safety
# property worth having in one tool. A SYNCED object is edited ON-PREMISES,
# always -- Active Directory is the authority and a cloud-side change to a
# synced object reverts on the next sync cycle. Nothing else in the toolkit
# stops you making that mistake in another window.
#
# Every change is confirmed before it is written and recorded in the audit log.
# TOOLKIT-HELP-END
# TOOLKIT-HELP-RUNS: On-premises Exchange, Active Directory and Exchange Online. Reads everywhere; writes only the one attribute you confirm.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
if (-not (Test-ConsoleLaunch)) { return }

#region Environment ------------------------------------------------------------
# THERE IS NO GUARD HERE, AND THAT IS DELIBERATE.
#
# Every other hybrid tool refuses to start when both Exchange worlds are loaded,
# because an unqualified cmdlet would go to the wrong one. This tool WANTS both
# and qualifies every call, so both being present is the ideal case rather than
# the failure case.
#
# WHAT WAS WRONG WITH THE FIRST VERSION, because the mistake is worth keeping.
#
# It probed both worlds ONCE at load time, cached the answer in two booleans,
# and then skipped an entire half of the environment when the answer was no.
# It opened with a banner reading "ON-PREMISES ONLY -- Exchange Online is not
# connected", which was narrowly true (the console process it runs in had the
# Exchange Management Shell loaded and had never been asked to connect to the
# cloud) and completely useless. A tool whose whole reason to exist is "find
# the object wherever it lives" cannot answer half the question and call that
# a result. Reporting the gap honestly is not the same as being usable.
#
# Group-Manager is the pattern this should have started from: it does not report
# that it lacks a connection, it GETS the connection and then answers. Three
# changes follow from that.
#
#   1. Nothing is cached. Both worlds are re-tested at the top of every search,
#      because a connection made after the window opened must count.
#   2. The state shown is the CONSOLE'S state, because this shares the console's
#      process. That is what makes "use Connect ExO in the console" a fix rather
#      than a brush-off -- the old message said the same words while running
#      detached, where the console's session was genuinely unreachable.
#   3. On-premises no longer depends on the Exchange Management Shell at all.
#      See Get-OnPremAdObject: Active Directory holds every fact this tool
#      displays, the ActiveDirectory module is always loadable, and it shares
#      no cmdlet name with either Exchange. EMS, when it happens to be present,
#      upgrades the on-premises half rather than enabling it.
function Test-OnPremLoaded {
    try { return [bool](Test-EmsAvailable).Available } catch { return $false }
}
function Test-CloudLoaded {
    try { return [bool](Test-ExoInProcess).Present } catch { return $false }
}


#endregion

#region Actions ----------------------------------------------------------------
# A CEILING PER SOURCE, ANNOUNCED WHEN HIT. Four sources means four chances to
# return thousands of rows, and a merged grid of five thousand is not a result.
# TWO CEILINGS, BECAUSE A BROAD SEARCH AND A NARROWED ONE DESERVE DIFFERENT ONES.
#
# WHAT ACTUALLY LIMITS THIS IS THE GRID, NOT EXCHANGE. -ResultSize accepts
# Unlimited and Get-ADObject -ResultSetSize likewise, so nothing cmdlet-side
# stops a bigger number. DataGridView rows added one at a time do: a few
# thousand is a visible pause, and past that you are scrolling a haystack.
#
# 250 WAS TOO LOW AND HID ITS OWN FAILURE. Searching a common word across four
# sources hit the cap on several of them, so the tool truncated silently and the
# operator filtered a set that was already incomplete -- forty tidy rows that
# looked like the whole answer.
#
# A NARROWED SEARCH CAN AFFORD MORE because it is returning only the type and
# side that were asked for. Trading a broad ceiling of 500 for a targeted one of
# 2000 puts the budget where the relevance is.
$script:MaxResults       = 500
$script:MaxResultsNarrow = 2000

# THE RESULT SET, HELD SO IT CAN BE FILTERED WITHOUT RE-QUERYING.
#
# The grid used to be the only copy of the answer, which made every filter a new
# round trip. Keeping the merged rows here means ticking a type is instant and
# costs Exchange nothing.
$script:AllRows      = @()
$script:WasCapped    = $false
$script:LastSearch   = ''

# EMPTY MEANS EVERYTHING. Not "nothing" -- a filter nobody has touched must not
# hide rows, and a null check at every use site is how that goes wrong once.
$script:FilterSource = @()
$script:FilterType   = @()

# WHAT A NARROWED RE-RUN ASKED FOR. Set by the Narrow button, consumed and
# cleared by the next search so it can never leak into a later plain one.
$script:NarrowSource = @()
$script:NarrowType   = @()

# THE CEILING IN FORCE FOR THE CURRENT SEARCH, read by all four fetches.
#
# A variable rather than a -Cap parameter on each of them: four signatures
# changed in four places is four chances to pass the wrong one, and every fetch
# in a given search wants the same number by definition.
$script:ActiveCap = $script:MaxResults

function ConvertTo-OuPath {
    <#
        The OU a distinguished name sits in, in the form ADUC shows it.

        'CN=TS-Shared,OU=Shared Mailboxes,OU=Sites,DC=corp,DC=x,DC=com'
            ->  corp.x.com/Sites/Shared Mailboxes

        SPLIT ON UNESCAPED COMMAS ONLY. A display name containing a comma --
        'CN=Armbruster\, Glenn' -- is escaped with a backslash in the DN, and
        splitting on every comma would tear it in half and put a fragment of
        somebody's surname in the OU column.

        The OU parts are reversed because a DN reads leaf-first and a person
        reads root-first.
    #>
    param([string]$DistinguishedName)

    $dn = "$DistinguishedName"
    if (-not $dn) { return '' }

    $parts = [regex]::Split($dn, '(?<!\\),')
    $ous = New-Object System.Collections.ArrayList
    $dcs = New-Object System.Collections.ArrayList
    foreach ($p in $parts) {
        $t = "$p".Trim()
        # The inner parentheses are load-bearing: inside a method call a comma
        # separates ARGUMENTS before it is anything else. See ARG001.
        if ($t -match '^(?i)OU=(.+)$')     { [void]$ous.Add(($Matches[1] -replace '\\,', ',')) }
        elseif ($t -match '^(?i)DC=(.+)$') { [void]$dcs.Add($Matches[1]) }
    }

    $domain = ($dcs -join '.')
    if ($ous.Count -eq 0) { return $domain }
    [array]::Reverse($ous)
    return ($domain + '/' + ($ous -join '/'))
}

function New-LookupRow {
    <#
        The one shape every source is flattened into.

        FOUR SOURCES RETURN FOUR DIFFERENT OBJECTS -- an on-premises recipient,
        an AD object, a cloud recipient and a soft-deleted mailbox all describe
        the same kinds of thing with different property names and different
        gaps. Normalising once, here, is what lets the merge and the grid stay
        simple; the alternative is every downstream function knowing which of
        four shapes it holds.
    #>
    param(
        [string]$Name, [string]$DisplayName, [string]$Alias, [string]$Primary,
        [string]$Type, [string]$Source, [string]$Dn, [string]$Directory,
        [string]$Server, $Addresses, $HiddenFromGal, $WhenSoftDeleted
    )
    return [PSCustomObject]@{
        Name            = "$Name"
        DisplayName     = "$DisplayName"
        Alias           = "$Alias"
        Primary         = "$Primary"
        OnPremType      = $(if ($Source -eq 'On-prem') { "$Type" } else { '' })
        CloudType       = $(if ($Source -ne 'On-prem') { "$Type" } else { '' })
        Source          = "$Source"
        Dn              = "$Dn"
        OuPath          = (ConvertTo-OuPath -DistinguishedName $Dn)
        Directory       = "$Directory"
        Server          = "$Server"
        Addresses       = @($Addresses | ForEach-Object { "$_" })
        HiddenFromGal   = $HiddenFromGal
        WhenSoftDeleted = $WhenSoftDeleted
    }
}

function Get-OnPremRecipient {
    <#
        Mail-enabled objects from on-premises Exchange, across every directory
        in Tenant.json.

        Invoke-OnPremCommand, NOT a bare Get-Recipient. That is the difference
        between this tool and AD-Lookup, and the reason this one can share a
        process with Exchange Online.

        EVERY DIRECTORY IS SEARCHED, not the first that answers. A name existing
        in two directories at once is not an ambiguity to resolve; it is the
        answer to some tickets.
    #>
    param([Parameter(Mandatory)][string]$Text)

    $rows = New-Object System.Collections.ArrayList
    $errs = New-Object System.Collections.ArrayList

    $t = ConvertTo-OpathValue -Value $Text
    $filter = "(Name -like '*$t*') -or (DisplayName -like '*$t*') -or " +
              "(Alias -like '*$t*') -or (EmailAddresses -like '*$t*')"

    $dirs = @()
    try { $dirs = @(Get-TenantDirectory) } catch { }
    if ($dirs.Count -eq 0) { $dirs = @([PSCustomObject]@{ Id = '' }) }

    foreach ($d in $dirs) {
        $id = "$($d.Id)"
        $dc = ''
        # Resolve-DirectoryDomainController, not Get-TenantDirectoryServer:
        # ARCHITECTURE 14f -- the latter returns a DOMAIN when a directory names
        # several controllers, and Exchange rejects that with error 0x52.
        try { if ($id) { $dc = Resolve-DirectoryDomainController -DirectoryId $id } } catch { }

        $p = @{ Filter = $filter; ResultSize = ($script:ActiveCap + 1); ErrorAction = 'Stop' }
        if ("$dc".Trim()) { $p['DomainController'] = $dc }

        try {
            foreach ($r in @(Invoke-OnPremCommand -Name 'Get-Recipient' -Parameters $p)) {
                [void]$rows.Add((New-LookupRow -Name $r.Name -DisplayName $r.DisplayName `
                    -Alias $r.Alias -Primary $r.PrimarySmtpAddress `
                    -Type $r.RecipientTypeDetails -Source 'On-prem' `
                    -Dn $r.DistinguishedName -Directory $(if ($id) { $id } else { 'default' }) `
                    -Server $dc -Addresses $r.EmailAddresses `
                    -HiddenFromGal $r.HiddenFromAddressListsEnabled -WhenSoftDeleted $null))
            }
        } catch {
            [void]$errs.Add("on-prem $(if ($id) { $id } else { 'default' }): $($_.Exception.Message)")
        }
    }
    return [PSCustomObject]@{ Rows = @($rows.ToArray()); Errors = @($errs.ToArray()) }
}

function Get-CloudRecipient {
    <#
        Mail-enabled objects from Exchange Online.

        Get-Recipient covers cloud distribution lists, Microsoft 365 groups
        (which come back as GroupMailbox), shared and room mailboxes, contacts
        and room lists in one call -- the same reason AD-Lookup uses it rather
        than Get-Mailbox and Get-DistributionGroup separately.
    #>
    param([Parameter(Mandatory)][string]$Text)

    $rows = New-Object System.Collections.ArrayList
    $errs = New-Object System.Collections.ArrayList

    $t = ConvertTo-OpathValue -Value $Text
    $filter = "(Name -like '*$t*') -or (DisplayName -like '*$t*') -or " +
              "(Alias -like '*$t*') -or (EmailAddresses -like '*$t*')"

    try {
        $found = @(Invoke-EXOCommand -Name 'Get-Recipient' -Parameters @{
            Filter      = $filter
            ResultSize  = ($script:ActiveCap + 1)
            ErrorAction = 'Stop'
        })
        foreach ($r in $found) {
            [void]$rows.Add((New-LookupRow -Name $r.Name -DisplayName $r.DisplayName `
                -Alias $r.Alias -Primary $r.PrimarySmtpAddress `
                -Type $r.RecipientTypeDetails -Source 'Cloud' `
                -Dn '' -Directory 'Exchange Online' -Server '' `
                -Addresses $r.EmailAddresses `
                -HiddenFromGal $r.HiddenFromAddressListsEnabled -WhenSoftDeleted $null))
        }
    } catch {
        [void]$errs.Add("cloud: $($_.Exception.Message)")
    }
    return [PSCustomObject]@{ Rows = @($rows.ToArray()); Errors = @($errs.ToArray()) }
}

function Get-SoftDeletedMailbox {
    <#
        Mailboxes deleted in the cloud and still recoverable.

        A SEPARATE CALL, because -SoftDeletedMailbox is a switch on Get-Mailbox
        and not a recipient type Get-Recipient can be asked for. Soft-deleted
        objects are invisible to every other search in this tool, which is
        exactly why they are worth a call of their own: "that address is in use
        and I cannot find what is using it" is usually one of these.

        FAILS QUIETLY INTO A REPORTED ERROR, not an exception. If the parameter
        set differs on some tenant, or the permission is missing, the other
        three sources still return -- losing the whole search over the least
        important source would be the wrong trade.
    #>
    param([Parameter(Mandatory)][string]$Text)

    $rows = New-Object System.Collections.ArrayList
    $errs = New-Object System.Collections.ArrayList

    $t = ConvertTo-OpathValue -Value $Text
    $filter = "(Name -like '*$t*') -or (DisplayName -like '*$t*') -or (Alias -like '*$t*')"

    try {
        $found = @(Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
            SoftDeletedMailbox = $true
            Filter             = $filter
            ResultSize         = ($script:ActiveCap + 1)
            ErrorAction        = 'Stop'
        })
        foreach ($r in $found) {
            [void]$rows.Add((New-LookupRow -Name $r.Name -DisplayName $r.DisplayName `
                -Alias $r.Alias -Primary $r.PrimarySmtpAddress `
                -Type $r.RecipientTypeDetails -Source 'Soft-deleted' `
                -Dn '' -Directory 'Exchange Online' -Server '' `
                -Addresses $r.EmailAddresses -HiddenFromGal $null `
                -WhenSoftDeleted $r.WhenSoftDeleted))
        }
    } catch {
        [void]$errs.Add("soft-deleted: $($_.Exception.Message)")
    }
    return [PSCustomObject]@{ Rows = @($rows.ToArray()); Errors = @($errs.ToArray()) }
}

# Resolve-AdObjectType MOVED TO Common.ps1 on 2026-08-20. AD-Lookup needed the
# same derivation the moment it stopped calling Get-Recipient unqualified, and
# two copies of a recipient-type rule is precisely the thing that goes subtly
# out of step. Get-OnPremAdObject below still calls it; it just lives elsewhere.

function Get-OnPremAdObject {
    <#
        THE ON-PREMISES HALF, AND IT DOES NOT NEED THE EXCHANGE MANAGEMENT SHELL.

        This function replaced Get-AdOnlyObject, which only ever returned objects
        with NO mailbox and left the real on-premises search to
        Get-OnPremRecipient -- which needs EMS loaded in this process. That was
        the bug the tool shipped with: in a console where EMS had not been
        connected, the on-premises half returned security groups and nothing
        else, and in a console where Exchange Online had not been connected the
        cloud half was skipped outright. Two optional halves and no answer.

        Active Directory is not optional in the same way. It holds every fact
        this tool displays for an on-premises object -- name, alias, addresses,
        OU, and enough Exchange attributes to say what kind of recipient it is
        (Resolve-AdObjectType) -- and Get-ADObject is SAFE IN THIS PROCESS
        regardless of what Exchange sessions are loaded, because it comes from
        the ActiveDirectory module and shares no cmdlet name with either
        Exchange. Unlike Get-Recipient it cannot be captured by the wrong one.

        EMS, when it is present, still runs and still wins: Get-OnPremRecipient
        gives Exchange's own RecipientTypeDetails, which knows about room lists
        and arbitration mailboxes that no AD attribute distinguishes. Pass
        -ExchangeAnswered in that case and the mail-enabled objects are dropped
        here rather than duplicated. EMS upgrades this half; it no longer gates it.

        (!(objectClass=computer)) IS LOAD-BEARING. In the AD schema `computer`
        is a SUBCLASS of `user`, so (objectClass=user) matches every domain
        computer -- an eight-hundred-office directory would bury three real
        matches under a thousand workstations.
    #>
    param(
        [Parameter(Mandatory)][string]$Text,
        [switch]$ExchangeAnswered,
        [switch]$IncludeNonMail
    )

    $rows = New-Object System.Collections.ArrayList
    $errs = New-Object System.Collections.ArrayList

    $esc  = ConvertTo-LdapFilterValue -Value $Text
    $ldap = "(&(|(objectClass=group)(objectClass=user)(objectClass=contact))(!(objectClass=computer))" +
            "(|(name=*$esc*)(displayName=*$esc*)(sAMAccountName=*$esc*)(mailNickname=*$esc*)(proxyAddresses=*$esc*)))"

    $props = @(
        'displayName', 'mail', 'sAMAccountName', 'objectClass', 'mailNickname',
        'proxyAddresses', 'homeMDB', 'targetAddress', 'groupType',
        'userAccountControl', 'msExchRemoteRecipientType',
        'msExchResourceSearchProperties', 'msExchHideFromAddressLists'
    )

    $dirs = @()
    try { $dirs = @(Get-TenantDirectory) } catch { }
    if ($dirs.Count -eq 0) { $dirs = @([PSCustomObject]@{ Id = '' }) }

    foreach ($d in $dirs) {
        $id = "$($d.Id)"
        $server = ''
        try { if ($id) { $server = Resolve-DirectoryDomainController -DirectoryId $id } } catch { }

        $p = @{
            LDAPFilter    = $ldap
            Properties    = $props
            ResultSetSize = ($script:ActiveCap + 1)
            ErrorAction   = 'Stop'
        }
        if ("$server".Trim()) { $p['Server'] = $server }

        try {
            foreach ($o in @(Get-ADObject @p)) {
                $proxies = @($o.proxyAddresses) | Where-Object { "$_".Trim() }
                $mailed  = ($proxies.Count -gt 0)

                # Exchange already returned this one, with better information
                # about it than AD has. Two rows for one object is not an answer.
                if ($mailed -and $ExchangeAnswered) { continue }
                if (-not $mailed -and -not $IncludeNonMail) { continue }

                # THE PRIMARY IS THE UPPERCASE SMTP: ENTRY. Exactly one entry in
                # proxyAddresses is uppercase and that is what makes it primary
                # -- the same rule Set-DistributionGroup enforces from the other
                # direction. -cmatch, not -match: case IS the data here.
                $primary = ''
                foreach ($a in $proxies) {
                    if ("$a" -cmatch '^SMTP:') { $primary = ("$a" -replace '^SMTP:', ''); break }
                }
                if (-not $primary) { $primary = "$($o.mail)".Trim() }

                $hidden = $null
                if ($null -ne $o.msExchHideFromAddressLists) { $hidden = [bool]$o.msExchHideFromAddressLists }

                $alias = "$($o.mailNickname)".Trim()
                if (-not $alias) { $alias = "$($o.sAMAccountName)".Trim() }

                [void]$rows.Add((New-LookupRow -Name $o.Name -DisplayName $o.displayName `
                    -Alias $alias -Primary $primary `
                    -Type (Resolve-AdObjectType -AdObject $o) `
                    -Source $(if ($mailed) { 'On-prem' } else { 'AD only' }) `
                    -Dn $o.DistinguishedName `
                    -Directory $(if ($id) { $id } else { 'default' }) -Server $server `
                    -Addresses $proxies -HiddenFromGal $hidden -WhenSoftDeleted $null))
            }
        } catch {
            [void]$errs.Add("AD $(if ($id) { $id } else { 'default' }): $($_.Exception.Message)")
        }
    }
    return [PSCustomObject]@{ Rows = @($rows.ToArray()); Errors = @($errs.ToArray()) }
}

#region Cloud attribute editing ------------------------------------------------
<#
    EDITING A CLOUD-ONLY OBJECT.

    AD-Lookup edits on-premises objects by writing raw LDAP attributes with
    Set-ADObject. For a cloud-only object there is no directory to write to --
    you go through Exchange with NAMED PARAMETERS, and a different cmdlet per
    recipient type. The two halves share a dialog and nothing else.

    EVERY PARAMETER BELOW WAS VERIFIED AGAINST THE MICROSOFT LEARN REFERENCE
    (and its authoring source) ON 2026-08-20, not recalled. Three findings
    changed the design and are encoded here rather than left as comments:

    1. Set-Mailbox -PrimarySmtpAddress and Set-MailContact -PrimarySmtpAddress
       ARE ON-PREMISES ONLY. Their "Applicable" lines omit Exchange Online
       entirely. Set-DistributionGroup and Set-UnifiedGroup do accept it in the
       cloud. So the primary address is set two different ways depending on
       type -- see Set-CloudPrimaryAddress.

    2. Documented on all four cmdlets, verbatim: "You can't use the
       EmailAddresses parameter and the PrimarySmtpAddress parameter in the same
       command." Never emit both.

    3. Also verbatim: "To replace all existing proxy email addresses with the
       values you specify, use the following syntax..." -- a bare list REPLACES
       THE WHOLE COLLECTION. That is the X500-loss trap from Set-DistributionGroup
       wearing cloud clothes, and it is why the address rebuild below carries
       every existing entry across.

    Set-UnifiedGroup has NO Name, NO WindowsEmailAddress and NO
    SimpleDisplayName, which is why the table is per-cmdlet rather than one
    shared list.
#>
$script:CloudCommonFields = @(
    'DisplayName', 'Alias', 'CustomAttribute1', 'CustomAttribute2',
    'CustomAttribute3', 'CustomAttribute4', 'CustomAttribute5',
    'CustomAttribute6', 'CustomAttribute7', 'CustomAttribute8',
    'CustomAttribute9', 'CustomAttribute10', 'CustomAttribute11',
    'CustomAttribute12', 'CustomAttribute13', 'CustomAttribute14',
    'CustomAttribute15'
)

function Get-CloudEditPlan {
    <#
        Which cmdlet edits this object, and which fields it accepts.

        Returns $null for anything this cannot safely edit, and the CALLER shows
        that as a greyed button with a reason -- an editor that opens on an
        object it cannot write is worse than one that declines.
    #>
    param([Parameter(Mandatory)]$Row)

    $t = "$(Get-RowTypeText -Row $Row)"

    # GroupMailbox is what Get-Recipient calls a Microsoft 365 group.
    if ($t -match '(?i)groupmailbox|unifiedgroup') {
        return [PSCustomObject]@{
            Cmdlet       = 'Set-UnifiedGroup'
            Fields       = @($script:CloudCommonFields + @('Notes'))
            BoolFields   = @('HiddenFromAddressListsEnabled')
            PrimaryParam = 'PrimarySmtpAddress'   # verified: accepted in EXO
        }
    }
    if ($t -match '(?i)distributiongroup|mailuniversal|roomlist') {
        return [PSCustomObject]@{
            Cmdlet       = 'Set-DistributionGroup'
            Fields       = @($script:CloudCommonFields + @('SimpleDisplayName'))
            BoolFields   = @('HiddenFromAddressListsEnabled')
            PrimaryParam = 'PrimarySmtpAddress'   # verified: accepted in EXO
        }
    }
    # GUEST MAIL USERS ARE NOT EDITED HERE, AND THIS IS THE ONE RULE IN THIS
    # FUNCTION THAT MICROSOFT DOES NOT DOCUMENT.
    #
    # The Exchange admin center refuses them with "You cannot edit guest mail
    # user here. For its management use https://portal.azure.com". A full search
    # of the Exchange documentation source finds GuestMailUser in exactly three
    # places, all of them lists of valid RecipientTypeDetails FILTER values on
    # Get- cmdlets. There is no published statement that Set-MailUser fails on
    # one, and no published statement that it works.
    #
    # So this block is EMPIRICAL, not doc-backed -- and it is written to fail
    # towards not writing. A guest is an Entra B2B object that Exchange only
    # shadows; the authority is the directory, and the portal is where it lives.
    if ($t -match '(?i)guestmailuser') { return $null }

    # MailUser AND MailContact ARE DIFFERENT CMDLETS. Routing MailUser to
    # Set-MailContact was a bug in the first version of this table.
    #
    # Microsoft publishes no RecipientTypeDetails-to-cmdlet mapping at all --
    # that table does not exist -- so the evidence is the Identity parameter
    # TYPE on each cmdlet: Set-MailContact takes MailContactIdParameter,
    # Set-MailUser takes MailUserIdParameter. Different .NET types, each scoped
    # to its own noun. Set-MailUser additionally carries 27 parameters
    # Set-MailContact lacks (UserPrincipalName, ExchangeGuid, MailboxRegion and
    # so on), which is exactly what you would expect of a thing that HAS a user
    # account versus a thing that does not.
    if ($t -match '(?i)mailuser') {
        return [PSCustomObject]@{
            Cmdlet       = 'Set-MailUser'
            # NO 'Office'. It does not exist on Set-MailUser -- the documented
            # route is Set-User -Office, which is a different cmdlet and a
            # different object, so it is deliberately not offered here rather
            # than offered and silently failing.
            Fields       = @($script:CloudCommonFields + @('SimpleDisplayName', 'ExternalEmailAddress'))
            BoolFields   = @('HiddenFromAddressListsEnabled')
            # '' NOT 'PrimarySmtpAddress', AND THIS ONE IS A TRAP. Unlike
            # Set-Mailbox and Set-MailContact, Set-MailUser DOES list Exchange
            # Online in the parameter's Applicable line -- so it looks
            # supported. The body of the same entry then says "This parameter is
            # functional only in on-premises Exchange." Accepted and inert is
            # worse than rejected: the command would succeed and change nothing.
            PrimaryParam = ''
        }
    }
    if ($t -match '(?i)mailcontact') {
        return [PSCustomObject]@{
            Cmdlet       = 'Set-MailContact'
            Fields       = @($script:CloudCommonFields + @('SimpleDisplayName', 'ExternalEmailAddress'))
            BoolFields   = @('HiddenFromAddressListsEnabled')
            PrimaryParam = ''                     # verified: on-prem only
        }
    }

    # SchedulingMailbox IS DELIBERATELY NOT MATCHED BY THE MAILBOX RULE BELOW.
    #
    # It is a real RecipientTypeDetails value -- it appears in the valid-values
    # list on Get-Mailbox, so Exchange classifies it as a mailbox subtype -- but
    # Microsoft has never published a single sentence defining what it is or
    # which cmdlet manages it. Set-Mailbox is the plausible answer and plausible
    # is not good enough for a write path. Left unroutable until somebody can
    # say what one is.
    if ($t -match '(?i)schedulingmailbox') { return $null }

    if ($t -match '(?i)mailbox') {
        return [PSCustomObject]@{
            Cmdlet       = 'Set-Mailbox'
            Fields       = @($script:CloudCommonFields + @('SimpleDisplayName', 'Office'))
            DeliveryFields = @()
            BoolFields   = @('HiddenFromAddressListsEnabled')
            PrimaryParam = ''                     # verified: on-prem only
        }
    }
    return $null
}

function Test-HasCloudMailbox {
    <#
        Does this row have a mailbox living in Exchange Online?

        THIS IS NOT THE SAME QUESTION AS "IS THE SOURCE COLUMN CLOUD", and
        conflating the two was a real bug. The Source column says where the
        OBJECT is mastered. For a REMOTE mailbox -- RemoteSharedMailbox,
        RemoteUserMailbox, RemoteRoomMailbox -- the directory stub is
        on-premises and the mailbox itself is in the cloud. One row, two homes.

        WHAT THAT BROKE: MessageCopyForSentAsEnabled is a property of the CLOUD
        mailbox. Routing every Synced row to AD-Lookup, which writes
        Set-ADObject, meant there was no path to it at all for exactly the
        shared mailboxes where send-as copies matter most.

        A plain on-premises UserMailbox or SharedMailbox -- no 'Remote' prefix --
        is genuinely hosted on-premises, so Set-Mailbox against Exchange Online
        would not find it. Those return false.
    #>
    param([Parameter(Mandatory)]$Row)

    if ("$($Row.Source)" -eq 'Soft-deleted') { return $false }

    $t = "$(Get-RowTypeText -Row $Row)"
    if ($t -notmatch '(?i)mailbox') { return $false }

    # 'GroupMailbox' CONTAINS THE WORD MAILBOX AND IS NOT ONE. It is what
    # Get-Recipient calls a Microsoft 365 group, and Get-Mailbox refuses it with
    # "The current operation is not supported on GroupMailbox."
    #
    # This substring collision is the single root of three separate bugs: the
    # forwarding probe erroring on a group, the routing table sending a group to
    # a distribution-list editor, and the licence line asking about a group that
    # cannot hold one. Everything that means MAILBOX has to say so explicitly.
    if ($t -match '(?i)groupmailbox') { return $false }

    # A remote mailbox stub always points at a cloud mailbox, whatever the
    # Source column says about the directory object.
    if ($t -match '(?i)remote.*mailbox') { return $true }

    # Otherwise only a cloud-mastered object has a cloud mailbox.
    return ("$($Row.Source)" -eq 'Cloud')
}

function Get-MailboxSettingsPlan {
    <#
        Exchange Online mailbox BEHAVIOUR, as opposed to directory identity.

        THE SPLIT THIS SITS ON: Attributes edits what an object IS and lives
        where the object is mastered -- Active Directory for anything with a
        distinguished name. Mailbox settings edit how a mailbox BEHAVES, and
        those always live in Exchange Online, even when the directory object
        does not.

        Everything here is Set-Mailbox against the cloud, for any row where
        Test-HasCloudMailbox is true.
    #>
    param([Parameter(Mandatory)]$Row)

    if (-not (Test-HasCloudMailbox -Row $Row)) { return $null }

    return [PSCustomObject]@{
        Cmdlet       = 'Set-Mailbox'
        Fields       = @()
        PrimaryParam = ''

        # FORWARDING. Handled apart from Fields because the smtp: prefix has to
        # be stripped for display, and because clearing one is the common
        # operation rather than setting it.
        #
        # TWO PARAMETERS, NOT ONE, AND THEY INTERACT. Documented verbatim: "If
        # you configure values for both the ForwardingAddress and
        # ForwardingSmtpAddress parameters, the value of ForwardingSmtpAddress
        # is ignored." ForwardingAddress takes an internal recipient;
        # ForwardingSmtpAddress takes any address and is what an external
        # forward uses. Both are shown so that silent precedence is visible.
        DeliveryFields = @('ForwardingAddress', 'ForwardingSmtpAddress')

        BoolFields = @(
            # THE SHARED-MAILBOX COMPLAINT, AND WHY THIS WENT FIRST. Without
            # these, somebody sending as a shared mailbox gets the sent item in
            # their OWN Sent Items and nobody else can see it.
            #
            # Documented note: previously shared-mailbox only, now available for
            # user mailboxes too -- but NOT for linked user mailboxes.
            'MessageCopyForSentAsEnabled',
            'MessageCopyForSendOnBehalfEnabled',

            # THE DANGEROUS ONE. $false means mail goes ONLY to the forward
            # target and does not land in the mailbox at all.
            'DeliverToMailboxAndForward'
        )
    }
}

function Get-CloudObjectProperty {
    <#
        Read the current values back from Exchange Online.

        The corresponding Get-* for each Set-*, because Get-Recipient does not
        return CustomAttribute values or Notes and the editor must show what is
        actually there rather than what the search happened to carry.
    #>
    param([Parameter(Mandatory)]$Plan, [Parameter(Mandatory)][string]$Identity)

    $getter = $Plan.Cmdlet -replace '^Set-', 'Get-'
    return (Invoke-EXOCommand -Name $getter -Parameters @{
        Identity = $Identity; ErrorAction = 'Stop' })
}

function Set-CloudPrimaryAddress {
    <#
        Change the primary SMTP address of a cloud object.

        TWO ROUTES, AND THE TYPE DECIDES WHICH -- see finding 1 above.

        Set-DistributionGroup and Set-UnifiedGroup take -PrimarySmtpAddress
        directly. Set-Mailbox and Set-MailContact do NOT in Exchange Online, so
        the whole EmailAddresses collection is rebuilt: the new address is
        written with an uppercase SMTP: prefix, the old primary is demoted to
        lowercase smtp: so mail to it still delivers, and EVERY OTHER ENTRY IS
        CARRIED ACROSS UNCHANGED.

        That last clause is the entire point. X500 addresses are what let replies
        to old mail resolve, and a bare EmailAddresses list replaces the whole
        collection -- so rebuilding without them is a silent, delayed breakage
        that shows up weeks later as "replies bounce".

        THE OLD PRIMARY IS KEPT, NOT REMOVED. Removing it breaks mail to the old
        address immediately, which is almost never what a rename wants.
    #>
    param(
        [Parameter(Mandatory)]$Plan,
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][string]$NewPrimary,
        [Parameter(Mandatory)]$ExistingAddresses
    )

    if ($Plan.PrimaryParam) {
        Invoke-EXOCommand -Name $Plan.Cmdlet -Parameters @{
            Identity                 = $Identity
            "$($Plan.PrimaryParam)"  = $NewPrimary
            ErrorAction              = 'Stop'
        } | Out-Null
        return
    }

    $rebuilt = New-Object System.Collections.ArrayList
    foreach ($a in @($ExistingAddresses)) {
        $s = "$a".Trim()
        if (-not $s) { continue }

        # Demote the current primary. -creplace, because the ONLY thing
        # distinguishing it from the others is that its prefix is uppercase.
        if ($s -cmatch '^SMTP:') { $s = $s -creplace '^SMTP:', 'smtp:' }

        # Drop any existing copy of the new address in either case, so it is not
        # added twice -- "You can use this value only once in a command".
        if ($s -match "(?i)^smtp:$([regex]::Escape($NewPrimary))$") { continue }
        [void]$rebuilt.Add($s)
    }
    [void]$rebuilt.Insert(0, "SMTP:$NewPrimary")

    Invoke-EXOCommand -Name $Plan.Cmdlet -Parameters @{
        Identity       = $Identity
        EmailAddresses = @($rebuilt.ToArray())
        ErrorAction    = 'Stop'
    } | Out-Null
}
#endregion

function Show-CloudAttributeEditor {
    <#
        The attribute editor for a CLOUD-ONLY object.

        Deliberately the same shape as AD-Lookup's: a grid of name and value,
        click or double-click to edit one field, confirm, write, refresh. An
        operator should not have to learn two editors because the object happens
        to live somewhere else.

        WHAT IS DIFFERENT IS UNDERNEATH. AD-Lookup writes LDAP attributes with
        Set-ADObject -Replace/-Clear. This builds a parameter for a Set-* cmdlet
        and sends it through Invoke-EXOCommand, which resolves the name against
        the Exchange Online module rather than trusting import order -- the same
        discipline the rest of this file uses.

        ONE FIELD PER SAVE, ON PURPOSE. Batching several changes into one
        command makes the audit line ambiguous and makes a partial failure
        impossible to read. One field, one confirmation, one log entry.
    #>
    param(
        [Parameter(Mandatory)]$Row,
        # THE PLAN IS PASSED IN, NOT DERIVED HERE. One dialog now serves two
        # buttons: Attributes hands it a Get-CloudEditPlan (directory identity)
        # and Mailbox settings hands it a Get-MailboxSettingsPlan (Exchange
        # Online behaviour). Same grid, same edit path, same audit line --
        # different set of fields and a different title.
        [Parameter(Mandatory)]$Plan,
        [string]$Title = 'Cloud attributes',
        $Owner
    )

    $plan = $Plan

    $ident = "$($Row.Primary)".Trim()
    if (-not $ident) { $ident = "$($Row.Name)".Trim() }

    $dlg                = New-Object System.Windows.Forms.Form
    $dlg.Text           = "$Title -- $($Row.Name)"
    $dlg.ClientSize     = New-Object System.Drawing.Size(760, 560)
    $dlg.StartPosition  = 'CenterParent'
    $dlg.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $lblTop          = New-Object System.Windows.Forms.Label
    $lblTop.Location = New-Object System.Drawing.Point(12, 10)
    $lblTop.Size     = New-Object System.Drawing.Size(736, 34)
    # THE SECOND LINE CHANGES WITH THE ROW, because "cloud-only" is not true of
    # every object this dialog now opens for. A RemoteSharedMailbox is mastered
    # on-premises and its MAILBOX is in the cloud -- saying it is cloud-only
    # would be wrong, and saying nothing would leave the operator wondering why
    # a Synced object is being edited in Exchange Online.
    $lblTop.Text     = "$ident  --  written with $($plan.Cmdlet) against Exchange Online." +
                       [Environment]::NewLine +
                       $(if ("$($Row.Source)" -eq 'Cloud') {
                           'This object is cloud-only, so there is no on-premises copy and nothing will sync down to it.'
                         } else {
                           'The directory object is on-premises, but the MAILBOX is in Exchange Online -- which is where these settings live.'
                         })
    $dlg.Controls.Add($lblTop)

    $grd = New-Object System.Windows.Forms.DataGridView
    $grd.Location = New-Object System.Drawing.Point(12, 50)
    $grd.Size     = New-Object System.Drawing.Size(736, 420)
    $grd.ReadOnly = $true
    $grd.AllowUserToAddRows    = $false
    $grd.AllowUserToDeleteRows = $false
    $grd.AllowUserToResizeRows = $false
    $grd.RowHeadersVisible     = $false
    $grd.SelectionMode         = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grd.MultiSelect           = $false
    $grd.BackgroundColor       = [System.Drawing.Color]::White
    $grd.Font                  = New-Object System.Drawing.Font('Consolas', 9)
    $grd.Anchor                = 'Top,Left,Right,Bottom'
    foreach ($c in @(
        @{ N = 'Attribute'; W = 230 },
        @{ N = 'Value';     W = 420 },
        @{ N = 'Editable';  W =  70 })) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c.N; $col.HeaderText = $c.N; $col.Width = $c.W
        [void]$grd.Columns.Add($col)
    }
    $dlg.Controls.Add($grd)

    $lblSt          = New-Object System.Windows.Forms.Label
    $lblSt.Location = New-Object System.Drawing.Point(12, 478)
    $lblSt.Size     = New-Object System.Drawing.Size(736, 34)
    $lblSt.Anchor   = 'Left,Right,Bottom'
    $dlg.Controls.Add($lblSt)

    $btnEdit          = New-Object System.Windows.Forms.Button
    $btnEdit.Text     = 'Edit'
    $btnEdit.Location = New-Object System.Drawing.Point(12, 518)
    $btnEdit.Size     = New-Object System.Drawing.Size(110, 30)
    $btnEdit.Enabled  = $false
    $btnEdit.Anchor   = 'Left,Bottom'
    $dlg.Controls.Add($btnEdit)

    $btnRefresh          = New-Object System.Windows.Forms.Button
    $btnRefresh.Text     = 'Refresh'
    $btnRefresh.Location = New-Object System.Drawing.Point(130, 518)
    $btnRefresh.Size     = New-Object System.Drawing.Size(110, 30)
    $btnRefresh.Anchor   = 'Left,Bottom'
    $dlg.Controls.Add($btnRefresh)

    $btnDone          = New-Object System.Windows.Forms.Button
    $btnDone.Text     = 'Done'
    $btnDone.Location = New-Object System.Drawing.Point(638, 518)
    $btnDone.Size     = New-Object System.Drawing.Size(110, 30)
    $btnDone.Anchor   = 'Right,Bottom'
    $dlg.Controls.Add($btnDone)

    function Update-CloudGrid {
        $grd.Rows.Clear()
        $btnEdit.Enabled = $false
        try {
            $o = Get-CloudObjectProperty -Plan $plan -Identity $ident
        } catch {
            $lblSt.ForeColor = [System.Drawing.Color]::Firebrick
            $lblSt.Text = Format-ErrorDetail -ErrorRecord $_ -Context "$($plan.Cmdlet -replace '^Set-', 'Get-')"
            return
        }

        # PRIMARY ADDRESS FIRST -- it is the field most identity edits are about,
        # and the one whose write path differs by type.
        #
        # ONLY WHEN THE PLAN HAS IDENTITY FIELDS. The Mailbox settings plan
        # carries no Fields at all: it is about how a mailbox behaves, not what
        # it is called. Offering an address editor there would put the same
        # field in two dialogs, which is how two dialogs start disagreeing.
        if (@($plan.Fields).Count -gt 0) {
            $i = $grd.Rows.Add('PrimarySmtpAddress', "$($o.PrimarySmtpAddress)", 'Yes')
            $grd.Rows[$i].Tag = [PSCustomObject]@{
                Name = 'PrimarySmtpAddress'; Value = "$($o.PrimarySmtpAddress)"
                Editable = $true; Kind = 'Primary' }
        }

        foreach ($f in @($plan.Fields)) {
            if (-not $o.PSObject.Properties[$f]) { continue }
            $v = "$($o.$f)"
            $i = $grd.Rows.Add($f, $v, 'Yes')
            $grd.Rows[$i].Tag = [PSCustomObject]@{ Name = $f; Value = $v; Editable = $true; Kind = 'String' }
        }

        # IS MAIL ACTUALLY LEAVING THIS MAILBOX? Computed once, before anything
        # is coloured, because two of the rows below only MEAN something in
        # relation to this.
        $hasForward = [bool]("$($o.ForwardingSmtpAddress)".Trim() -or "$($o.ForwardingAddress)".Trim())

        # FORWARDING, DISPLAYED WITHOUT THE PREFIX. Exchange returns
        # ForwardingSmtpAddress as 'smtp:user@domain.com'; the prefix is
        # Exchange's, not the operator's, and it is not what you type back in.
        # Stripping it also makes the unchanged-value comparison work.
        foreach ($f in @($plan.DeliveryFields)) {
            if (-not $o.PSObject.Properties[$f]) { continue }
            $v = "$($o.$f)" -replace '^(?i)smtp:', ''
            $i = $grd.Rows.Add($f, $v, 'Yes')
            $grd.Rows[$i].Tag = [PSCustomObject]@{ Name = $f; Value = $v; Editable = $true; Kind = 'Forward' }
            if ($v) { $grd.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Firebrick }
        }

        foreach ($f in @($plan.BoolFields)) {
            if (-not $o.PSObject.Properties[$f]) { continue }
            $v = "$([bool]$o.$f)"
            $i = $grd.Rows.Add($f, $v, 'Yes')
            $grd.Rows[$i].Tag = [PSCustomObject]@{ Name = $f; Value = $v; Editable = $true; Kind = 'Bool' }

            # MAIL LEAVES AND DOES NOT STAY. Red means "this is doing something
            # you should notice", and the -and $hasForward is the whole rule.
            #
            # THE FIRST VERSION COLOURED ANY False, WHICH WAS WRONG. With no
            # forward configured, DeliverToMailboxAndForward is inert -- mail
            # delivers to the mailbox normally and False means nothing at all.
            # Colouring it red on a perfectly ordinary mailbox is a false alarm,
            # and a tool that cries wolf on the normal case teaches you to
            # ignore the colour on the abnormal one.
            #
            # Worth noting how this got shipped: the comment that used to sit
            # here said "False WITH A FORWARD CONFIGURED means nothing is being
            # kept" -- the right rule, described accurately, next to code that
            # only tested half of it.
            if ($f -eq 'DeliverToMailboxAndForward' -and $v -eq 'False' -and $hasForward) {
                $grd.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Firebrick
            }
        }

        # ADDRESSES ARE SHOWN AND NOT EDITED HERE. The collection is replaced
        # wholesale by Exchange when written, so a single text box over it is
        # data loss dressed as an edit -- the same rule AD-Lookup applies to
        # multi-valued LDAP attributes. Changing the primary rewrites the
        # collection safely; changing the list by hand belongs somewhere that
        # can show every entry at once.
        # Identity dialog only, for the same reason the primary row is.
        $addrs = @()
        if (@($plan.Fields).Count -gt 0) { $addrs = @($o.EmailAddresses | ForEach-Object { "$_" }) }
        if ($addrs.Count) {
            $i = $grd.Rows.Add('EmailAddresses', "($($addrs.Count)) " + ($addrs -join ' | '), '')
            $grd.Rows[$i].Tag = [PSCustomObject]@{ Name = 'EmailAddresses'; Value = ''; Editable = $false; Kind = 'Multi' }
            $grd.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray
        }

        # THE COLOUR IS EXPLAINED RATHER THAN LEFT TO BE DECODED. Red on a grid
        # is a claim that something matters, and a claim nobody can check is
        # just decoration.
        # ONLY THE MAILBOX SETTINGS DIALOG TALKS ABOUT FORWARDING. Saying "no
        # forwarding is configured" under a distribution group's identity fields
        # would be true, irrelevant, and faintly alarming.
        $talksAboutMail = (@($plan.DeliveryFields).Count -gt 0)

        if ($talksAboutMail -and $hasForward) {
            $keep = [bool]$o.DeliverToMailboxAndForward
            $lblSt.ForeColor = [System.Drawing.Color]::Firebrick
            $lblSt.Text = if ($keep) {
                'Red: mail is being FORWARDED. A copy is still kept in this mailbox.'
            } else {
                'Red: mail is FORWARDED AND NOT KEPT here. Nothing is landing in this mailbox.'
            }
        } elseif ($talksAboutMail) {
            $lblSt.ForeColor = [System.Drawing.Color]::DimGray
            $lblSt.Text = "$($grd.Rows.Count) field(s).  No forwarding configured, so DeliverToMailboxAndForward has no effect."
        } else {
            $lblSt.ForeColor = [System.Drawing.Color]::DimGray
            $lblSt.Text = "$($grd.Rows.Count) field(s).  Written with $($plan.Cmdlet)."
        }
    }

    $grd.Add_SelectionChanged({
        $btnEdit.Enabled = $false
        if ($grd.SelectedRows.Count -eq 0) { return }
        $t = $grd.SelectedRows[0].Tag
        if ($t -and $t.Editable) { $btnEdit.Enabled = $true }
    })

    function Invoke-CloudEdit {
        if ($grd.SelectedRows.Count -eq 0) { return }
        $t = $grd.SelectedRows[0].Tag
        if (-not $t -or -not $t.Editable) { return }

        $box                 = New-Object System.Windows.Forms.Form
        $box.Text            = "Edit $($t.Name)"
        $box.ClientSize      = New-Object System.Drawing.Size(520, 210)
        $box.FormBorderStyle = 'FixedDialog'
        $box.MaximizeBox     = $false
        $box.MinimizeBox     = $false
        $box.StartPosition   = 'CenterParent'
        $box.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

        $l1 = New-Object System.Windows.Forms.Label
        $l1.Text = "Current value of $($t.Name):"
        $l1.Location = New-Object System.Drawing.Point(12, 12); $l1.Size = New-Object System.Drawing.Size(480, 18)
        $box.Controls.Add($l1)

        $tc = New-Object System.Windows.Forms.TextBox
        $tc.Location = New-Object System.Drawing.Point(12, 32); $tc.Size = New-Object System.Drawing.Size(496, 22)
        $tc.ReadOnly = $true; $tc.Text = "$($t.Value)"
        $tc.Font = New-Object System.Drawing.Font('Consolas', 9)
        $box.Controls.Add($tc)

        $l2 = New-Object System.Windows.Forms.Label
        $l2.Location = New-Object System.Drawing.Point(12, 66); $l2.Size = New-Object System.Drawing.Size(480, 18)
        $l2.Text = if ($t.Kind -eq 'Bool') { 'New value  (True or False):' }
                   elseif ($t.Kind -eq 'Primary') { 'New primary address  (the old one is kept as a secondary):' }
                   elseif ($t.Kind -eq 'Forward') { 'Forward to  (leave empty to STOP forwarding):' }
                   else { 'New value  (leave empty to clear):' }
        $box.Controls.Add($l2)

        $tn = New-Object System.Windows.Forms.TextBox
        $tn.Location = New-Object System.Drawing.Point(12, 86); $tn.Size = New-Object System.Drawing.Size(496, 22)
        $tn.Text = "$($t.Value)"
        $tn.Font = New-Object System.Drawing.Font('Consolas', 9)
        $box.Controls.Add($tn)

        $ok = New-Object System.Windows.Forms.Button
        $ok.Text = 'Save'; $ok.Location = New-Object System.Drawing.Point(304, 150)
        $ok.Size = New-Object System.Drawing.Size(96, 30)
        $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $box.Controls.Add($ok); $box.AcceptButton = $ok

        $ca = New-Object System.Windows.Forms.Button
        $ca.Text = 'Cancel'; $ca.Location = New-Object System.Drawing.Point(412, 150)
        $ca.Size = New-Object System.Drawing.Size(96, 30)
        $ca.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $box.Controls.Add($ca); $box.CancelButton = $ca

        Set-ConsoleTheme -Form $box
        if ($box.ShowDialog($dlg) -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $newValue = $tn.Text.Trim()
        if ("$newValue" -ceq "$($t.Value)") {
            $lblSt.ForeColor = [System.Drawing.Color]::DimGray
            $lblSt.Text = 'Unchanged -- nothing written.'
            return
        }

        if ($t.Kind -eq 'Bool' -and $newValue -notmatch '^(?i)(true|false)$') {
            $lblSt.ForeColor = [System.Drawing.Color]::Firebrick
            $lblSt.Text = "$($t.Name) takes True or False."
            return
        }
        # A FORWARD MAY BE CLEARED, so empty is valid here -- but a non-empty
        # value must still look like an address. ForwardingAddress technically
        # accepts any recipient identity, and this rejects a bare alias; that is
        # a deliberate trade. A typo'd alias resolving to the WRONG recipient is
        # a worse outcome than being asked to type the full address.
        if ($t.Kind -eq 'Forward' -and $newValue -ne '' -and
            $newValue -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
            $lblSt.ForeColor = [System.Drawing.Color]::Firebrick
            $lblSt.Text = 'Enter a full SMTP address to forward to, or clear the box to stop forwarding.'
            return
        }
        if ($t.Kind -eq 'Primary') {
            if ($newValue -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
                $lblSt.ForeColor = [System.Drawing.Color]::Firebrick
                $lblSt.Text = 'That is not a valid SMTP address -- nothing written.'
                return
            }
        }

        $shown = if ($newValue -eq '') { '(cleared)' } else { $newValue }

        # THE CONFIRMATION SAYS WHAT THE CHANGE MEANS, NOT JUST WHAT IT SETS.
        #
        # 'DeliverToMailboxAndForward: True -> False' is accurate and tells an
        # operator nothing. What it MEANS is that mail stops landing in the
        # mailbox -- which is the difference between a routine change and an
        # incident, and it belongs in front of the person clicking Yes.
        $extra = ''
        if ($t.Kind -eq 'Primary') {
            $extra = [Environment]::NewLine + [Environment]::NewLine +
                     'The old address is kept as a secondary, so mail to it still delivers.'
        }
        elseif ($t.Kind -eq 'Forward') {
            $extra = if ($newValue -eq '') {
                [Environment]::NewLine + [Environment]::NewLine +
                'Forwarding STOPS. Mail will be delivered only to this mailbox.'
            } else {
                [Environment]::NewLine + [Environment]::NewLine +
                "Mail sent to this mailbox will also go to $newValue." + [Environment]::NewLine +
                'Check DeliverToMailboxAndForward as well -- if that is False, mail goes ONLY' + [Environment]::NewLine +
                'to the forward target and does not stay in the mailbox.'
            }
        }
        elseif ($t.Name -eq 'DeliverToMailboxAndForward' -and $newValue -match '^(?i)false$') {
            $extra = [Environment]::NewLine + [Environment]::NewLine +
                     'WARNING: with a forward configured, False means mail is delivered ONLY to' + [Environment]::NewLine +
                     'the forward target. It will NOT be kept in this mailbox.'
        }
        elseif ($t.Name -like 'MessageCopyFor*' -and $newValue -match '^(?i)true$') {
            $extra = [Environment]::NewLine + [Environment]::NewLine +
                     "A copy of anything sent this way will be saved in this mailbox's Sent Items," + [Environment]::NewLine +
                     "as well as the sender's own."
        }

        $confirm = [System.Windows.Forms.MessageBox]::Show(
            ("Change $($t.Name) on $($Row.Name) in Exchange Online?" + [Environment]::NewLine + [Environment]::NewLine +
             "From: $($t.Value)" + [Environment]::NewLine + "To:   $shown" + $extra),
            'Confirm cloud attribute change',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        try {
            if ($t.Kind -eq 'Primary') {
                Set-CloudPrimaryAddress -Plan $plan -Identity $ident `
                    -NewPrimary $newValue -ExistingAddresses $Row.Addresses
            } else {
                # $null CLEARS, '' DOES NOT. Exchange treats an empty string as
                # an invalid value for an address-typed parameter and rejects
                # it; removing the value is $null. Same distinction AD-Lookup
                # makes with Set-ADObject -Clear versus -Replace, arrived at
                # from the other direction.
                $val = if ($t.Kind -eq 'Bool') { [bool]::Parse($newValue) }
                       elseif ($newValue -eq '') { $null }
                       else { $newValue }
                Invoke-EXOCommand -Name $plan.Cmdlet -Parameters @{
                    Identity      = $ident
                    "$($t.Name)"  = $val
                    ErrorAction   = 'Stop'
                } | Out-Null
            }

            # $ident, NOT $Row.Name -- it is the exact identity the write used,
            # so the log names the object the change actually landed on.
            Write-CompletedLog -Action 'EditCloudAttribute' `
                -Target (Get-LogTarget $ident $Row.Primary $Row.Name) -Result 'Success' `
                -Environment 'Exchange Online' `
                -Detail "$($t.Name): '$($t.Value)' -> '$shown' (Unified Lookup, $($plan.Cmdlet))"

            $lblSt.ForeColor = [System.Drawing.Color]::ForestGreen
            $lblSt.Text = "$($t.Name) changed. Grid refreshed from Exchange Online."
            Update-CloudGrid
        } catch {
            Write-FailedLog -Action 'EditCloudAttribute' `
                -Target (Get-LogTarget $ident $Row.Primary $Row.Name) -ErrorRecord $_ `
                -Environment 'Exchange Online' -Detail "$($t.Name)"
            $lblSt.ForeColor = [System.Drawing.Color]::Firebrick
            $lblSt.Text = Format-ErrorDetail -ErrorRecord $_ -Context "$($plan.Cmdlet)"
        }
    }

    $btnEdit.Add_Click({ Invoke-CloudEdit })
    $grd.Add_CellDoubleClick({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        Invoke-CloudEdit
    })
    $btnRefresh.Add_Click({ Update-CloudGrid })
    $btnDone.Add_Click({ $dlg.Close() })

    Update-CloudGrid
    Set-ConsoleTheme -Form $dlg
    [void]$dlg.ShowDialog($Owner)
}
#endregion

#region Routing table ----------------------------------------------------------
<#
    WHERE DOES THIS OBJECT GO TO BE EDITED.

    This is Group-Manager's pattern generalised. Group-Manager reads one property
    (IsDirSynced), decides which side of the hybrid owns a distribution group,
    and hands Group-Manager-Editor the answer rather than letting it guess. The
    same two questions apply to every object type:

        WHICH TOOL  -- decided by what kind of recipient it is
        WHICH SIDE  -- decided by the Source column, which is already on screen

    ORDER MATTERS. The first entry whose Source and Type both match wins, so
    specific rules go above general ones and the last entry is the catch-all.

    TO ADD A TARGET, ADD A ROW. That is the whole extension mechanism and it is
    why this is a table rather than a switch statement buried in a click
    handler.

        Sources -- which values of the Source column this row applies to
        Types   -- regex matched against the recipient type. '.' matches all.
        Script  -- the file to open
        Where   -- 'OnPrem' launches a new process that connects to on-premises
                   Exchange first; 'InProcess' runs it in this window's process,
                   which is where the Exchange Online session lives
        Extra   -- switches handed to the child, e.g. -OnPrem
#>
$script:HandOff = @(
    # DISTRIBUTION GROUPS, EITHER SIDE. The path that already works today, and
    # the reason -OnPrem exists on Group-Manager-Editor: the caller read the
    # authority and tells it, so the editor never has to guess and never falls
    # back to the wrong side when one fails.
    [PSCustomObject]@{
        Label   = 'Edit group'
        Sources = @('On-prem', 'Synced')
        Types   = 'Distribution|MailUniversal|RoomList'
        Script  = 'Group-Manager-Editor.ps1'
        Where   = 'OnPrem'
        Extra   = @('-OnPrem')
    },
    [PSCustomObject]@{
        Label   = 'Edit group'
        Sources = @('Cloud')
        Types   = 'Distribution|MailUniversal|RoomList'
        Script  = 'Group-Manager-Editor.ps1'
        Where   = 'InProcess'
        Extra   = @()
    },

    # EVERYTHING ELSE ON-PREMISES. AD-Lookup is the attribute editor and the
    # ADUC replacement, so it is the right destination for a mailbox, a contact
    # or a security group alike. 'AD only' is included on purpose: an object
    # with no address is exactly the one you most often need to look at
    # properly.
    # MICROSOFT 365 GROUPS GO TO GROUP EDITOR, NOT DL-MANAGER-EDITOR.
    #
    # This row exists because the first version put GroupMailbox in the
    # distribution-group row above, and it failed live with Exchange's own
    # words: "The current operation is not supported on GroupMailbox."
    # Group-Manager-Editor is built on Get-/Set-DistributionGroup; a Microsoft 365
    # group needs Get-/Set-UnifiedGroup, which is what GroupEditor uses.
    #
    # The two are not interchangeable and the name does not warn you --
    # Get-Recipient reports a Microsoft 365 group as 'GroupMailbox', a word that
    # contains neither 'Unified' nor 'M365' and does contain 'Mailbox'.
    [PSCustomObject]@{
        Label   = 'Edit group'
        Sources = @('Cloud')
        Types   = 'GroupMailbox|UnifiedGroup'
        Script  = 'GroupEditor.ps1'
        Where   = 'InProcess'
        Extra   = @()
    },

    # MAILBOXES, EITHER SIDE, GO TO MAILBOX PERMISSIONS.
    #
    # This row is why the Open button was greyed out on a cloud user mailbox.
    # The first version of this table had no entry for mailboxes at all -- only
    # groups and an AD-Lookup catch-all -- so a cloud mailbox matched nothing.
    #
    # AND THE CATCH-ALL WAS THE REAL MISTAKE. It sent every on-premises object
    # to AD-Lookup, which is exactly what the Attributes button already does. So
    # for an on-premises row the two buttons did the SAME THING under different
    # names, and for a cloud mailbox one of them did nothing at all. Two buttons
    # that overlap on half your objects and leave a hole on the other half is
    # worse than either one alone.
    #
    # THE SPLIT NOW, and it is worth stating because it is the rule for adding
    # rows later:
    #
    #   Attributes  edits WHAT THE OBJECT IS -- names, addresses, custom
    #               attributes. Cloud objects here, on-premises ones in
    #               AD-Lookup.
    #   Open        opens WHAT THE OBJECT DOES -- the membership of a group,
    #               the delegations on a mailbox.
    #
    # MailboxPermissions carries no OwnProcess marker and takes -SmtpAddress, so
    # it runs in this process for both sides. Permissions deliberately stayed
    # out of the Attributes editor -- two tools that can grant the same access
    # produce an audit trail with two shapes.
    [PSCustomObject]@{
        Label   = 'Open permissions'
        Sources = @('On-prem', 'Synced', 'Cloud')
        # (?!Group) EXCLUDES GroupMailbox -- see Test-HasCloudMailbox.
        # Get-MailboxPermission is as unsupported on a Microsoft 365 group as
        # Get-Mailbox is.
        Types   = '(?<!Group)Mailbox'
        Script  = 'MailboxPermissions.ps1'
        Where   = 'InProcess'
        Extra   = @()
    }

    # NO CATCH-ALL ROW, DELIBERATELY. A contact or a plain security group has no
    # second tool -- Attributes is the whole story for them -- and a row that
    # opened roughly the right thing would look like the feature works. Where
    # there is nothing to open, the button greys out.
)

function Get-HandOff {
    <#
        The first routing entry that matches this row, or $null.

        SOFT-DELETED NEVER MATCHES. A deleted mailbox cannot be edited by any
        of these tools, and sending someone into one with a soft-deleted
        identity produces a "not found" that reads like the tool is broken.
    #>
    param([Parameter(Mandatory)]$Row)

    if ($Row.Source -eq 'Soft-deleted') { return $null }
    $type = "$(Get-RowTypeText -Row $Row)"

    foreach ($h in $script:HandOff) {
        if ($h.Sources -notcontains "$($Row.Source)") { continue }
        if ("$($h.Types)" -ne '.' -and $type -notmatch $h.Types) { continue }
        if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot $h.Script))) { continue }
        return $h
    }
    return $null
}
#endregion

function Merge-LookupRow {
    <#
        One row per object, however many sources found it.

        MERGED ON PRIMARY SMTP, because that is the only identifier the two
        worlds share. An on-premises remote mailbox stub and its cloud mailbox
        have different distinguished names, different GUIDs in some cases, and
        the same address -- the address is what makes them the same thing.

        Objects with NO address key on their distinguished name instead, so two
        different security groups with the same name in two directories stay two
        rows. They are two objects.

        ON-PREM WINS THE MERGE, and that is not arbitrary. For a synced object
        Active Directory is the authority and the cloud copy follows it, so the
        DN, the OU and the directory shown must be the on-premises ones -- they
        are the ones you would act on. The cloud type is kept alongside because
        the pair is informative: RemoteSharedMailbox on-premises and
        SharedMailbox in the cloud is a correctly synced shared mailbox.

        SOFT-DELETED NEVER MERGES INTO A LIVE ROW. A deleted mailbox sharing an
        address with a live object is not one object in two places -- it is the
        single most interesting thing the search can return, and hiding it
        inside another row would bury it.
    #>
    param([Parameter(Mandatory)]$Rows)

    $byKey = [ordered]@{}
    $out   = New-Object System.Collections.ArrayList

    foreach ($r in $Rows) {
        if ($r.Source -eq 'Soft-deleted') { [void]$out.Add($r); continue }

        $key = if ("$($r.Primary)".Trim()) { "smtp:$("$($r.Primary)".ToLowerInvariant())" }
               else                        { "dn:$("$($r.Dn)".ToLowerInvariant())" }

        if (-not $byKey.Contains($key)) {
            $byKey[$key] = $r
            continue
        }

        $existing = $byKey[$key]
        $onPrem = if ($existing.Source -eq 'On-prem' -or $existing.Source -eq 'AD only') { $existing } else { $r }
        $cloud  = if ($existing.Source -eq 'Cloud') { $existing } else { $r }

        # Rebuilt rather than mutated. $existing may be the cloud row, and
        # editing it in place would leave the merged result carrying cloud
        # identity fields for an object whose authority is on-premises.
        $merged = New-LookupRow -Name $onPrem.Name -DisplayName $onPrem.DisplayName `
            -Alias $onPrem.Alias -Primary $onPrem.Primary -Type '' -Source 'Synced' `
            -Dn $onPrem.Dn -Directory $onPrem.Directory -Server $onPrem.Server `
            -Addresses $onPrem.Addresses -HiddenFromGal $onPrem.HiddenFromGal `
            -WhenSoftDeleted $null
        $merged.OnPremType = "$($onPrem.OnPremType)"
        $merged.CloudType  = "$($cloud.CloudType)"
        $byKey[$key] = $merged
    }

    foreach ($k in $byKey.Keys) { [void]$out.Add($byKey[$k]) }
    return @($out.ToArray())
}

function Get-RowTypeText {
    <#
        The single Type string for the grid.

        A synced object has two, and showing both in a 190px column would show
        neither. The on-premises one is displayed because it is the authority;
        the pair is spelled out in the detail pane, where there is room for it
        to mean something.
    #>
    param([Parameter(Mandatory)]$Row)
    if ($Row.OnPremType) { return $Row.OnPremType }
    return $Row.CloudType
}

function ConvertTo-DnDisplayName {
    <#
        The readable name out of a distinguished name.

        'CN=Miller\, Jennifer L,OU=Users,DC=corp,DC=com'  ->  'Miller, Jennifer L'

        manager and managedBy hold a DN, and a DN is not an answer to "who owns
        this". Same function as AD-Lookup's, kept local rather than shared
        because it is six lines and moving it would mean editing a file that is
        working.

        SPLIT ON UNESCAPED COMMAS ONLY. 'CN=Armbruster\, Glenn' is ONE component
        with a comma in it, and naive splitting tears the name in half.
    #>
    param([string]$Dn)
    if (-not "$Dn".Trim()) { return '' }
    $first = ([regex]::Split("$Dn", '(?<!\\),'))[0]
    return (($first -replace '^\s*CN=', '') -replace '\\,', ',')
}

function Get-UnifiedAdDetail {
    <#
        The AD-sourced facts for one row, read ONCE and cached on the row.

        ONE QUERY, NOT SIX. Account state, title, department, company, manager
        and managedBy live on the same object. The detail pane redraws on every
        selection change, so clicking down a list of twenty results would
        otherwise be a hundred and twenty directory queries for one screen.

        Get-ADObject, so it is safe in this process whatever Exchange sessions
        are loaded -- the same reason the search half uses it.

        CLOUD-ONLY ROWS HAVE NO DN, so this returns empties and the caller
        prints nothing. There is no on-premises object to ask about.
    #>
    param([Parameter(Mandatory)]$Row)

    if ($Row.PSObject.Properties['AdDetail']) { return $Row.AdDetail }

    $d = [PSCustomObject]@{
        Uac = $null; Title = ''; Department = ''; Company = ''
        Manager = ''; ManagedBy = ''; Error = ''
    }

    if ("$($Row.Dn)".Trim()) {
        try {
            $ap = @{
                Identity    = "$($Row.Dn)"
                Properties  = @('userAccountControl', 'title', 'department', 'company', 'manager', 'managedBy')
                ErrorAction = 'Stop'
            }
            if ("$($Row.Server)".Trim()) { $ap['Server'] = "$($Row.Server)" }
            $o = Get-ADObject @ap
            if ($null -ne $o.userAccountControl) { $d.Uac = [int]$o.userAccountControl }
            $d.Title      = "$($o.title)"
            $d.Department = "$($o.department)"
            $d.Company    = "$($o.company)"
            $d.Manager    = ConvertTo-DnDisplayName -Dn "$($o.manager)"
            $d.ManagedBy  = ConvertTo-DnDisplayName -Dn "$($o.managedBy)"
        } catch {
            # NOT FATAL. The address and the type are what most tickets needed.
            $d.Error = $_.Exception.Message
        }
    }

    Add-Member -InputObject $Row -NotePropertyName 'AdDetail' -NotePropertyValue $d -Force
    return $d
}

function Get-RowOuText {
    <#
        The OU column, which for a cloud object is a question with no answer.

        A BLANK CELL IS THE WRONG ANSWER HERE. Blank reads as "this tool did not
        find the OU" -- and the natural next move is to go looking for one, or
        to report the tool as broken. There is no OU to find: organisational
        units are an Active Directory construct, and a cloud-mastered object has
        no Active Directory object at all.

        That distinction is obvious to anyone who has run a hybrid for years and
        is not obvious at all to somebody three weeks in, which is exactly the
        person most likely to be handed a lookup tool. Saying 'N/A (cloud)'
        costs one column and closes the question.

        Soft-deleted rows get the same treatment for the same reason.
    #>
    param([Parameter(Mandatory)]$Row)
    if ("$($Row.OuPath)".Trim()) { return "$($Row.OuPath)" }
    if ("$($Row.Source)" -eq 'Cloud' -or "$($Row.Source)" -eq 'Soft-deleted') { return 'N/A (cloud)' }
    return ''
}

function Format-UnifiedDetail {
    param([Parameter(Mandatory)]$Row)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Name                 : $($Row.Name)")
    [void]$sb.AppendLine("DisplayName          : $($Row.DisplayName)")
    [void]$sb.AppendLine("Alias                : $($Row.Alias)")
    if ($Row.Primary) { [void]$sb.AppendLine("PrimarySmtpAddress   : $($Row.Primary)") }
    [void]$sb.AppendLine("Source               : $($Row.Source)")

    if ($Row.OnPremType) { [void]$sb.AppendLine("Type (on-premises)   : $($Row.OnPremType)") }
    if ($Row.CloudType)  { [void]$sb.AppendLine("Type (cloud)         : $($Row.CloudType)") }

    if ($Row.Dn) {
        [void]$sb.AppendLine("OU                   : $($Row.OuPath)")
        [void]$sb.AppendLine("Directory            : $($Row.Directory)")
        [void]$sb.AppendLine("DistinguishedName    : $($Row.Dn)")
    } else {
        # SAID, NOT OMITTED -- see Get-RowOuText. An absent OU line invites the
        # reader to go looking for one; this line ends the question.
        [void]$sb.AppendLine('OU                   : N/A -- cloud objects do not live in an OU.')
    }
    if ($null -ne $Row.HiddenFromGal) {
        $h = if ($Row.HiddenFromGal) { 'YES -- not in the address book' } else { 'No' }
        [void]$sb.AppendLine("Hidden from GAL      : $h")
    }
    if ($Row.WhenSoftDeleted) {
        [void]$sb.AppendLine("Deleted              : $($Row.WhenSoftDeleted)")
    }

    # ORGANISATION, OWNERSHIP, ACCOUNT STATE AND LICENCE -- parity with
    # AD-Lookup, and for the same reasons its comments give.
    #
    # Printed only when populated. A blank line is a fact about the layout; an
    # absent line is a fact about the object.
    $ad = Get-UnifiedAdDetail -Row $Row
    if ($ad.Title)      { [void]$sb.AppendLine("Title                : $($ad.Title)") }
    if ($ad.Department) { [void]$sb.AppendLine("Department           : $($ad.Department)") }
    if ($ad.Company)    { [void]$sb.AppendLine("Company              : $($ad.Company)") }
    if ($ad.Manager)    { [void]$sb.AppendLine("Manager              : $($ad.Manager)") }
    if ($ad.ManagedBy)  { [void]$sb.AppendLine("ManagedBy            : $($ad.ManagedBy)") }
    if ($ad.Error)      { [void]$sb.AppendLine("Directory read       : failed -- $($ad.Error)") }

    # ACCOUNT STATE, for things that have an account. A distribution list and a
    # contact have none, so the line is absent rather than answered wrongly.
    #
    # 0x0002 is ACCOUNTDISABLE. THE RAW NUMBER IS PRINTED BESIDE IT on purpose:
    # if this decode is ever wrong, the value that proves it is on the same line
    # rather than one tool away.
    $typeText = "$(Get-RowTypeText -Row $Row)"
    if ($typeText -match '(?i)mailbox|mailuser' -and $null -ne $ad.Uac) {
        $acct = if ($ad.Uac -band 2) { "DISABLED  (userAccountControl $($ad.Uac))" }
                else                 { "Enabled  (userAccountControl $($ad.Uac))" }
        [void]$sb.AppendLine("Account              : $acct")
    }

    # LICENCE, FOR THINGS THAT CAN HOLD ONE.
    #
    # Get-UserLicenseState reads AD GROUP MEMBERSHIP, not Graph -- licensing
    # here is group-based, so the licence somebody holds is a fact about the
    # directory and readable with the rights the operator already has. It names
    # the GROUP that drives the assignment rather than the SKU, because the
    # group is the thing that can be acted on.
    #
    # SHARED AND ROOM MAILBOXES INCLUDED DELIBERATELY -- a licensed shared
    # mailbox is usually somebody's mistake, and Mailbox Health flags it for
    # exactly that reason.
    #
    # ASKED FOR EVERY MAILBOX, INCLUDING CLOUD-ONLY ONES -- and the version of
    # this that special-cased them was wrong. Recording why, because the
    # reasoning was plausible and still wrong.
    #
    # The worry was that Get-UserLicenseState reads AD group membership, so a
    # cloud-only object with no AD account would come back as 'not licensed' and
    # read as a finding rather than as a question that does not apply. That
    # cannot happen. Get-UserLicenseState returns Known = $false with a Reason
    # when it cannot find an account, and Format-UserLicenseState renders that
    # as "Unknown -- No on-premises account matches". Its own docstring is
    # explicit: callers MUST render Known = $false differently from
    # Licensed = $false. The distinction was already handled, one layer down.
    #
    # AND SKIPPING WOULD HAVE HIDDEN REAL ANSWERS. The lookup searches mail,
    # userPrincipalName AND proxyAddresses across both forests, precisely
    # because a cloud object's primary SMTP is frequently a proxy address
    # on-premises. A 'cloud' object here often DOES have an on-premises account
    # behind it, and that account's licence is a real answer.
    #
    # THE GENERAL RULE THIS BROKE: when a shared function already distinguishes
    # 'no' from 'cannot say', do not add a second opinion above it. Two places
    # describing one state is exactly what Format-UserLicenseState exists to
    # prevent -- see its docstring.
    if ($typeText -match '(?i)mailbox' -and $typeText -notmatch '(?i)groupmailbox' -and
        "$($Row.Primary)".Trim()) {

        # CACHED ON THE ROW after the first look, for the same reason as
        # Get-UnifiedAdDetail: this pane redraws on every selection change.
        if (-not $Row.PSObject.Properties['LicenseText']) {
            $lic = ''
            try {
                $st  = Get-UserLicenseState -Smtp "$($Row.Primary)"
                $lic = Format-UserLicenseState -State $st
            } catch {
                # A licence lookup that fails is not a reason to lose the rest
                # of the pane.
                $lic = "could not be read: $($_.Exception.Message)"
            }
            Add-Member -InputObject $Row -NotePropertyName 'LicenseText' -NotePropertyValue $lic -Force
        }
        [void]$sb.AppendLine("Licence              : $($Row.LicenseText)")
    }

    [void]$sb.AppendLine('')
    switch ($Row.Source) {
        'Synced' {
            [void]$sb.AppendLine('SYNCED. Active Directory is the authority for this object and the cloud')
            [void]$sb.AppendLine('copy follows it. A change made directly in Exchange Online will be')
            [void]$sb.AppendLine('overwritten on the next sync cycle -- edit the on-premises side.')
        }
        'Cloud' {
            [void]$sb.AppendLine('CLOUD ONLY. There is no on-premises object, so nothing on-premises can')
            [void]$sb.AppendLine('edit it and nothing will sync down to it. Use the cloud tools.')
        }
        'On-prem' {
            [void]$sb.AppendLine('ON-PREMISES ONLY. Either it is not in scope for sync, or it has not')
            [void]$sb.AppendLine('synced yet. If you expected it in the cloud, that is the thing to check.')
        }
        'AD only' {
            [void]$sb.AppendLine('ACTIVE DIRECTORY ONLY -- no mailbox and no address. Nothing can be')
            [void]$sb.AppendLine('delegated to it and mail cannot be sent to it. It is listed because its')
            [void]$sb.AppendLine('name is why the search looked ambiguous.')
        }
        'Soft-deleted' {
            [void]$sb.AppendLine('SOFT-DELETED. Deleted in the cloud, still recoverable, and STILL HOLDING')
            [void]$sb.AppendLine('ITS ADDRESSES -- which is usually why a new object cannot have them.')
            [void]$sb.AppendLine('It is shown as its own row rather than merged into a live one, because a')
            [void]$sb.AppendLine('deleted object sharing an address with a live one is the finding.')
        }
    }

    # FORWARDING, SHOWN WITHOUT BEING ASKED FOR.
    #
    # "Where is this person's mail going" is a question you answer from a
    # lookup, not from an editor -- and a forward is invisible until somebody
    # thinks to look for one. Cached on the row like the rest of the pane.
    #
    # CLOUD MAILBOXES ONLY, because it costs an Exchange Online round trip and
    # answers nothing for a group or a contact.
    if ((Test-HasCloudMailbox -Row $Row) -and "$($Row.Primary)".Trim()) {

        if (-not $Row.PSObject.Properties['ForwardText']) {
            $fwd = ''
            try {
                $mb = Invoke-EXOCommand -Name 'Get-Mailbox' -Parameters @{
                    Identity = "$($Row.Primary)"; ErrorAction = 'Stop' }
                $to = "$($mb.ForwardingSmtpAddress)" -replace '^(?i)smtp:', ''
                if (-not $to) { $to = "$($mb.ForwardingAddress)" }
                if ($to) {
                    $keep = [bool]$mb.DeliverToMailboxAndForward
                    $fwd  = if ($keep) { "$to  (a copy is still kept here)" }
                            else       { "$to  -- AND NOT KEPT HERE" }
                }
            } catch {
                $fwd = "could not be read: $($_.Exception.Message)"
            }
            Add-Member -InputObject $Row -NotePropertyName 'ForwardText' -NotePropertyValue $fwd -Force
        }
        if ($Row.ForwardText) {
            [void]$sb.AppendLine("Forwarding to        : $($Row.ForwardText)")
        }
    }

    # WHAT THE TWO BUTTONS WILL DO WITH THIS OBJECT, INCLUDING WHEN THE ANSWER
    # IS NOTHING.
    #
    # A greyed button is a question the tool refuses to answer. The operator's
    # options are to guess, to read the source, or to conclude it is broken --
    # and on a GuestMailUser, where BOTH buttons behave unusually, all three are
    # bad outcomes. Two lines here cost nothing and remove the guessing.
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Actions:')

    $h = Get-HandOff -Row $Row
    if ($h) {
        [void]$sb.AppendLine("    Open        $($h.Label) -- $($h.Script)")
    } elseif ("$($Row.Source)" -eq 'Soft-deleted') {
        [void]$sb.AppendLine('    Open        nothing -- a deleted mailbox has no tool to open.')
    } else {
        [void]$sb.AppendLine('    Open        nothing for this type. Open is for what an object DOES --')
        [void]$sb.AppendLine('                a group opens its membership, a mailbox opens its')
        [void]$sb.AppendLine('                permissions. This type has neither.')
    }

    if ("$($Row.Source)" -eq 'Soft-deleted') {
        [void]$sb.AppendLine('    Attributes  nothing -- a deleted mailbox cannot be edited.')
    } elseif ("$($Row.Source)" -eq 'Cloud') {
        $plan = Get-CloudEditPlan -Row $Row
        if ($plan) {
            [void]$sb.AppendLine("    Attributes  edits it here, with $($plan.Cmdlet).")
        } elseif ($typeText -match '(?i)guestmailuser') {
            [void]$sb.AppendLine('    Attributes  not offered. A guest is an Entra B2B object that Exchange')
            [void]$sb.AppendLine('                only shadows -- the admin center refuses it too, and')
            [void]$sb.AppendLine('                points at portal.azure.com. Manage it there.')
        } else {
            [void]$sb.AppendLine("    Attributes  not offered -- no verified cloud editor for a $typeText.")
        }
    } elseif ("$($Row.Dn)".Trim()) {
        [void]$sb.AppendLine('    Attributes  opens AD Lookup, which edits the on-premises object.')
        if ("$($Row.Source)" -eq 'Synced') {
            [void]$sb.AppendLine('                On-premises is the RIGHT side for a synced object --')
            [void]$sb.AppendLine('                a cloud-side edit reverts on the next sync.')
        }
    }

    if (Test-HasCloudMailbox -Row $Row) {
        [void]$sb.AppendLine('    Mailbox     forwarding and send-as copies, in Exchange Online.')
        if ("$($Row.Source)" -ne 'Cloud') {
            [void]$sb.AppendLine('                The directory object is on-premises but the MAILBOX is')
            [void]$sb.AppendLine('                in the cloud, so these settings are written there.')
        }
    } elseif ($typeText -match '(?i)mailbox' -and $typeText -notmatch '(?i)groupmailbox' -and
              "$($Row.Source)" -ne 'Soft-deleted') {
        [void]$sb.AppendLine('    Mailbox     nothing -- this mailbox is hosted on-premises, so Exchange')
        [void]$sb.AppendLine('                Online has no copy of it to configure.')
    }

    $addrs = @($Row.Addresses)
    if ($addrs.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("All addresses ($($addrs.Count)):")
        $ordered = $addrs | Sort-Object `
            { if ($_ -cmatch '^SMTP:') { 0 } elseif ($_ -match '^(?i)smtp:') { 1 } else { 2 } },
            { "$_" }
        foreach ($a in $ordered) { [void]$sb.AppendLine("    $a") }
    }
    return $sb.ToString()
}
#endregion

#region Filtering ---------------------------------------------------------------
function Show-ColumnFilter {
    <#
        A checklist of the values PRESENT IN THIS RESULT SET, with counts.

        NOT A FIXED LIST OF EVERY POSSIBLE TYPE. Offering RemoteEquipmentMailbox
        on a search that returned three distribution groups is a menu of things
        that cannot happen. The values come from the rows in hand, so the filter
        can only ever narrow what is actually there.

        COUNTS ARE THE POINT, not decoration. "Cloud (412)" and "Synced (7)" tell
        you where the noise is before you tick anything -- which is usually the
        whole question when a search comes back too big.

        RETURNS AN OBJECT, NOT AN ARRAY, and that is a bug fix rather than a
        style choice.

        This used to return @() for "no filter" and $null for "cancelled".
        POWERSHELL UNROLLS AN EMPTY ARRAY ON RETURN, so @() arrived at the
        caller as $null -- indistinguishable from Cancel. The caller's
        'if ($null -ne $new)' guard then treated "clear this filter" as "the
        operator changed their mind", and the filter could never be removed.
        Ticking All and pressing Apply did nothing at all.

        An object cannot unroll. .Cancelled says which button was pressed and
        .Values carries the selection, empty meaning everything.
    #>
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)]$Values,      # every value in the result set, with repeats
        [Parameter(Mandatory)]$Selected,    # currently ticked, empty = all
        $Owner
    )

    $counts = @{}
    foreach ($v in @($Values)) {
        $k = "$v"
        if (-not $k) { $k = '(none)' }
        if ($counts.ContainsKey($k)) { $counts[$k]++ } else { $counts[$k] = 1 }
    }
    if ($counts.Keys.Count -eq 0) { return $null }

    $dlg                 = New-Object System.Windows.Forms.Form
    $dlg.Text            = "Filter by $Title"
    $dlg.ClientSize      = New-Object System.Drawing.Size(320, 380)
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false
    $dlg.StartPosition   = 'CenterParent'
    $dlg.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Location      = New-Object System.Drawing.Point(12, 12)
    $clb.Size          = New-Object System.Drawing.Size(296, 292)
    $clb.CheckOnClick  = $true
    $clb.IntegralHeight = $false
    $dlg.Controls.Add($clb)

    $keys = @($counts.Keys | Sort-Object)
    foreach ($k in $keys) {
        # TICKED WHEN NOTHING IS FILTERED, so the dialog opens showing the
        # current state truthfully -- "all of these are included" -- rather than
        # empty, which reads as "nothing is included".
        $on = if (@($Selected).Count -eq 0) { $true } else { @($Selected) -contains $k }
        [void]$clb.Items.Add("$k  ($($counts[$k]))", $on)
    }

    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text = 'All'; $btnAll.Location = New-Object System.Drawing.Point(12, 312)
    $btnAll.Size = New-Object System.Drawing.Size(70, 26)
    $btnAll.Add_Click({ for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $true) } })
    $dlg.Controls.Add($btnAll)

    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text = 'None'; $btnNone.Location = New-Object System.Drawing.Point(90, 312)
    $btnNone.Size = New-Object System.Drawing.Size(70, 26)
    $btnNone.Add_Click({ for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $false) } })
    $dlg.Controls.Add($btnNone)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'Apply'; $ok.Location = New-Object System.Drawing.Point(116, 346)
    $ok.Size = New-Object System.Drawing.Size(96, 26)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($ok); $dlg.AcceptButton = $ok

    $ca = New-Object System.Windows.Forms.Button
    $ca.Text = 'Cancel'; $ca.Location = New-Object System.Drawing.Point(216, 346)
    $ca.Size = New-Object System.Drawing.Size(92, 26)
    $ca.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($ca); $dlg.CancelButton = $ca

    Set-ConsoleTheme -Form $dlg
    if ($dlg.ShowDialog($Owner) -ne [System.Windows.Forms.DialogResult]::OK) {
        return [PSCustomObject]@{ Cancelled = $true; Values = @() }
    }

    $picked = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $clb.Items.Count; $i++) {
        if ($clb.GetItemChecked($i)) { [void]$picked.Add($keys[$i]) }
    }

    # EVERYTHING TICKED IS THE SAME AS NO FILTER, and storing it that way keeps
    # the header clean and the "is anything filtered" test a simple count.
    if ($picked.Count -eq $keys.Count) {
        return [PSCustomObject]@{ Cancelled = $false; Values = @() }
    }
    return [PSCustomObject]@{ Cancelled = $false; Values = @($picked.ToArray()) }
}

function Get-FilteredRow {
    <#
        The rows that survive the current filters. Empty filter = everything.
    #>
    $out = @($script:AllRows)
    if (@($script:FilterSource).Count -gt 0) {
        $out = @($out | Where-Object { @($script:FilterSource) -contains "$($_.Source)" })
    }
    if (@($script:FilterType).Count -gt 0) {
        $out = @($out | Where-Object { @($script:FilterType) -contains "$(Get-RowTypeText -Row $_)" })
    }
    return $out
}
#endregion

#region WinForms UI ------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Unified Lookup'
$form.ClientSize    = New-Object System.Drawing.Size(1104, 950)
$form.StartPosition = 'Manual'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblWarn           = New-Object System.Windows.Forms.Label
$lblWarn.Tag       = 'NoTheme'
$lblWarn.Location  = New-Object System.Drawing.Point(12, 12)
$lblWarn.Size      = New-Object System.Drawing.Size(1030, 18)
$lblWarn.Anchor    = 'Top,Left,Right'
$lblWarn.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblWarn)

# NO CONNECT BUTTON HERE. THE CONSOLE OWNS CONNECTIONS.
#
# There was one, briefly, and removing it is the right call. The console already
# has a Connect Scripts section whose whole job is establishing sessions, and
# those sessions live in the console's process -- which is this tool's process
# too, now that the OwnProcess marker is gone from the header. A second connect
# button here would be a second place to do one thing, and the two would drift
# the moment Connect-ExO.ps1 grew an option only one of them passed.
#
# WHAT REPLACES IT IS A SENTENCE THAT SAYS WHERE TO GO. That is not the same as
# the sentence this tool used to show. The old one said "Exchange Online is not
# connected" when the tool had been launched detached and COULD NOT SEE the
# console's session -- it was reporting a problem the operator could not fix
# from where they were standing. Now the state shown here IS the console's
# state, so pointing at the console is pointing at the fix.
function Update-ConnectionBanner {
    <#
        Re-reads BOTH worlds and repaints. Called at startup and at the top of
        every search -- never cached, because connecting in the console while
        this window is open has to count on the next search.
    #>
    $onPrem = Test-OnPremLoaded
    $cloud  = Test-CloudLoaded

    if ($cloud -and $onPrem) {
        $lblWarn.Text      = 'BOTH WORLDS CONNECTED. Active Directory, on-premises Exchange and Exchange Online are all being searched.'
        # FromArgb(80, 220, 130) -- the SAME light green the console header uses
        # for a live connection and for an up-to-date version. Two greens meaning
        # "good" in one toolkit is one green too many, and the dark
        # FromArgb(21, 115, 71) that was here is a BUTTON FILL colour: it is
        # legible as a background behind white text and nearly invisible as text
        # on the dark console background.
        $lblWarn.ForeColor = [System.Drawing.Color]::FromArgb(80, 220, 130)
    } elseif ($cloud) {
        $lblWarn.Text      = 'ACTIVE DIRECTORY + EXCHANGE ONLINE. On-premises objects come from AD directly, so the on-premises half is covered.'
        # FromArgb(80, 220, 130) -- the SAME light green the console header uses
        # for a live connection and for an up-to-date version. Two greens meaning
        # "good" in one toolkit is one green too many, and the dark
        # FromArgb(21, 115, 71) that was here is a BUTTON FILL colour: it is
        # legible as a background behind white text and nearly invisible as text
        # on the dark console background.
        $lblWarn.ForeColor = [System.Drawing.Color]::FromArgb(80, 220, 130)
    } else {
        $lblWarn.Text      = 'ON-PREMISES ONLY -- Active Directory is being searched. Use Connect ExO in the console to add cloud and soft-deleted objects, then search again.'
        $lblWarn.ForeColor = [System.Drawing.Color]::Goldenrod
    }
}

$lblSearch          = New-Object System.Windows.Forms.Label
$lblSearch.Text     = 'Name contains:'
$lblSearch.Location = New-Object System.Drawing.Point(12, 42)
$lblSearch.Size     = New-Object System.Drawing.Size(100, 18)
$form.Controls.Add($lblSearch)

$txtSearch          = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(116, 39)
$txtSearch.Size     = New-Object System.Drawing.Size(260, 24)
$txtSearch.Font     = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtSearch)

$btnSearch          = New-Object System.Windows.Forms.Button
$btnSearch.Text     = 'Search'
$btnSearch.Location = New-Object System.Drawing.Point(388, 38)
$btnSearch.Size     = New-Object System.Drawing.Size(100, 26)
$form.Controls.Add($btnSearch)

# 'Include AD objects with no address' WAS A HIDDEN, ALWAYS-TICKED CHECKBOX
# HERE, AND IT IS GONE. The reason it survived as a control was wrong.
#
# The comment justifying it claimed "six places downstream read $chkAd.Checked,
# and replacing all six with $true is a wider change than hiding one checkbox".
# There was ONE read. I never counted them -- and the hidden control then sat in
# the layout collecting an overlap warning from a button with every right to
# that space.
#
# WHY IT IS ALWAYS ON: the box gated whether objects with NO address were
# searched -- plain security groups, unmailboxed users. Those are precisely the
# things that make a name look ambiguous, and 'HQTS' the security group sitting
# beside 'HQTS-Shared' the mailbox is the reason you opened a lookup tool at
# all. Asking twice for the second half of the answer was the wrong trade; the
# cost is one extra Active Directory query per search.

# SHOWN ONLY WHEN IT WOULD HELP. See Update-ResultGrid: the search must have
# been capped AND a filter must be set, otherwise this re-runs the same query.
$btnNarrow          = New-Object System.Windows.Forms.Button
$btnNarrow.Text     = 'Search again, narrowed'
$btnNarrow.Location = New-Object System.Drawing.Point(496, 38)
$btnNarrow.Size     = New-Object System.Drawing.Size(170, 26)
$btnNarrow.Visible  = $false
$form.Controls.Add($btnNarrow)

# CLEARING A FILTER NEEDS TO BE VISIBLE FROM OUTSIDE THE DIALOG.
#
# The only way out used to be: click the header again, press All, press Apply.
# That is three steps behind a control you have to already know about -- and
# while a bug made it impossible, nothing on the form said the state was even
# reversible. An affordance that only exists inside the thing you are trying to
# undo is not an affordance.
#
# Shown only when something IS filtered, so it doubles as the indicator that a
# filter is on. The header text says which column; this says there is one.
$btnClearFilter          = New-Object System.Windows.Forms.Button
$btnClearFilter.Text     = 'Clear filters'
$btnClearFilter.Location = New-Object System.Drawing.Point(674, 38)
$btnClearFilter.Size     = New-Object System.Drawing.Size(88, 26)
$btnClearFilter.Visible  = $false
$form.Controls.Add($btnClearFilter)

$chkDeleted          = New-Object System.Windows.Forms.CheckBox
$chkDeleted.Text     = 'Include soft-deleted mailboxes'
$chkDeleted.Location = New-Object System.Drawing.Point(766, 42)
$chkDeleted.Size     = New-Object System.Drawing.Size(230, 20)

# OFF BY DEFAULT. It shipped on, and on is the wrong default for the same reason
# the red colouring was: a deleted mailbox is a FINDING, and a finding that
# appears in every routine search stops reading as one. It also costs an extra
# Exchange Online call on every lookup to answer 'no' almost every time.
#
# The case it exists for -- "that address is in use and I cannot find what is
# using it" -- is a question you know you are asking, so ticking the box is part
# of asking it.
$chkDeleted.Checked  = $false
$form.Controls.Add($chkDeleted)

$lblCount           = New-Object System.Windows.Forms.Label
$lblCount.Location  = New-Object System.Drawing.Point(12, 68)
$lblCount.Size      = New-Object System.Drawing.Size(1080, 18)
$lblCount.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblCount)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(12, 90)
$grid.Size     = New-Object System.Drawing.Size(1080, 300)
$grid.ReadOnly = $true
$grid.AllowUserToAddRows    = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.RowHeadersVisible     = $false
$grid.SelectionMode         = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grid.MultiSelect           = $false
$grid.BackgroundColor       = [System.Drawing.Color]::White
$grid.BorderStyle           = [System.Windows.Forms.BorderStyle]::Fixed3D
$grid.Font                  = New-Object System.Drawing.Font('Consolas', 9)
$grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::AliceBlue
$grid.Anchor                = 'Top,Left,Right'
foreach ($c in @(
    @{ N = 'Name';    H = 'Name';          W = 170 },
    @{ N = 'Display'; H = 'Display Name';  W = 170 },
    @{ N = 'Address'; H = 'Email Address'; W = 230 },
    @{ N = 'Type';    H = 'Type';          W = 190 },
    @{ N = 'Src';     H = 'Source';        W = 110 },
    @{ N = 'Ou';      H = 'OU';            W = 202 })) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $c.N; $col.HeaderText = $c.H; $col.Width = $c.W

    # SRC AND TYPE FILTER INSTEAD OF SORTING. Everything else still sorts.
    #
    # A header click cannot do both. Leaving these Automatic meant one click
    # opened the filter dialog AND re-sorted the grid behind it -- and the sort
    # was then thrown away by the next re-paint, since Update-ResultGrid rebuilds
    # from $script:AllRows.
    #
    # Filtering to one source is a stronger version of grouping by it, so
    # nothing is lost: the old comment here argued sorting Source was "the
    # fastest way to see that the object you want is cloud-only", and ticking
    # Cloud answers that better by removing the rest.
    $col.SortMode = if ($c.N -eq 'Src' -or $c.N -eq 'Type') {
        [System.Windows.Forms.DataGridViewColumnSortMode]::NotSortable
    } else {
        [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
    }
    [void]$grid.Columns.Add($col)
}
$form.Controls.Add($grid)

$lblSel           = New-Object System.Windows.Forms.Label
$lblSel.Text      = 'Selected object:'
$lblSel.Location  = New-Object System.Drawing.Point(12, 398)
$lblSel.Size      = New-Object System.Drawing.Size(300, 18)
$form.Controls.Add($lblSel)

$txtDetail            = New-Object System.Windows.Forms.TextBox
$txtDetail.Multiline  = $true
$txtDetail.ReadOnly   = $true
$txtDetail.ScrollBars = 'Vertical'
$txtDetail.Location   = New-Object System.Drawing.Point(12, 418)
$txtDetail.Size       = New-Object System.Drawing.Size(1080, 480)
$txtDetail.Font       = New-Object System.Drawing.Font('Consolas', 9)
$txtDetail.Anchor     = 'Top,Left,Right,Bottom'
$form.Controls.Add($txtDetail)

$btnCopy          = New-Object System.Windows.Forms.Button
$btnCopy.Text     = 'Copy address'
$btnCopy.Location = New-Object System.Drawing.Point(12, 908)
$btnCopy.Size     = New-Object System.Drawing.Size(120, 30)
$btnCopy.Anchor   = 'Left,Bottom'
$form.Controls.Add($btnCopy)

$btnCopyDetail          = New-Object System.Windows.Forms.Button
$btnCopyDetail.Text     = 'Copy all detail'
$btnCopyDetail.Location = New-Object System.Drawing.Point(140, 908)
$btnCopyDetail.Size     = New-Object System.Drawing.Size(120, 30)
$btnCopyDetail.Anchor   = 'Left,Bottom'
$form.Controls.Add($btnCopyDetail)

# THE HAND-OFF. Its caption changes with the selected row, because "Open" tells
# you nothing about where you are about to be sent and this button opens a
# different tool depending on what is highlighted.
$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open'
$btnOpen.Location = New-Object System.Drawing.Point(268, 908)
$btnOpen.Size     = New-Object System.Drawing.Size(160, 30)
$btnOpen.Anchor   = 'Left,Bottom'
$btnOpen.Enabled  = $false
$form.Controls.Add($btnOpen)

# ATTRIBUTES. Cloud objects are edited HERE; on-premises ones are handed to
# AD-Lookup, which already does that job well and is the ADUC replacement.
#
# WHY NOT ONE EDITOR YET: AD-Lookup's Show-AttributeEditor is two hundred lines
# that work and are in daily use. Copying it in to claim "one tool" would leave
# two copies of an editor that WRITES TO PRODUCTION, and they would drift. The
# honest end state is that function moving to Common.ps1 with both callers using
# it -- worth doing when AD-Lookup is parked, not in the same edit that adds the
# cloud half.
$btnAttrs          = New-Object System.Windows.Forms.Button
$btnAttrs.Text     = 'Attributes'
$btnAttrs.Location = New-Object System.Drawing.Point(436, 908)
$btnAttrs.Size     = New-Object System.Drawing.Size(110, 30)
$btnAttrs.Anchor   = 'Left,Bottom'
$btnAttrs.Enabled  = $false
$form.Controls.Add($btnAttrs)

# MAILBOX SETTINGS -- how a mailbox BEHAVES, as opposed to what it is called.
#
# A SEPARATE BUTTON BECAUSE THE TWO LIVE IN DIFFERENT PLACES FOR THE SAME ROW.
# A RemoteSharedMailbox has its directory object on-premises and its mailbox in
# Exchange Online. Attributes edits the first with Set-ADObject; this edits the
# second with Set-Mailbox. Folding them into one button meant one of the two was
# always unreachable -- which is exactly what happened: send-as copies could not
# be set on any synced shared mailbox at all.
#
# Enabled by Test-HasCloudMailbox, NOT by the Source column.
$btnMbx          = New-Object System.Windows.Forms.Button
$btnMbx.Text     = 'Mailbox settings'
$btnMbx.Location = New-Object System.Drawing.Point(554, 908)
$btnMbx.Size     = New-Object System.Drawing.Size(140, 30)
$btnMbx.Anchor   = 'Left,Bottom'
$btnMbx.Enabled  = $false
$form.Controls.Add($btnMbx)

$lblStatus          = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(704, 916)
$lblStatus.Size     = New-Object System.Drawing.Size(268, 18)
$lblStatus.Anchor   = 'Left,Bottom'
$form.Controls.Add($lblStatus)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(982, 908)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Anchor   = 'Right,Bottom'
$form.Controls.Add($btnClose)

function Update-ResultGrid {
    <#
        Paint the grid from $script:AllRows, honouring the filters.

        ONE RENDER PATH FOR SEARCHING AND FILTERING. Ticking a type re-paints
        from rows already in hand and costs Exchange nothing -- which is the
        whole reason the result set is kept.

        SuspendLayout AROUND THE FILL. DataGridView re-lays-out on every Add,
        so a two-thousand-row paint spends most of its time on arithmetic
        nobody sees. This was missing when the ceiling was 250 and it did not
        matter; at 2000 it does.
    #>
    param([string]$Note = '')

    $rows = @(Get-FilteredRow)

    $grid.SuspendLayout()
    try {
        $grid.Rows.Clear()
        foreach ($r in $rows) {
            $i = $grid.Rows.Add("$($r.Name)", "$($r.DisplayName)", "$($r.Primary)",
                                (Get-RowTypeText -Row $r), "$($r.Source)", (Get-RowOuText -Row $r))
            $grid.Rows[$i].Tag = $r
            # Colour carries where a thing lives, and only where it needs to.
            # Soft-deleted is red because it is a finding; AD-only is grey
            # because it is context. Everything else is left alone -- if every
            # row is coloured, none of them is.
            if ($r.Source -eq 'Soft-deleted') {
                $grid.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Firebrick
            } elseif ($r.Source -eq 'AD only') {
                $grid.Rows[$i].DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray
            }
        }
    } finally { $grid.ResumeLayout() }

    # A FILTERED COLUMN SAYS SO IN ITS OWN HEADER. Without this the only clue
    # that rows are hidden is a number you have to remember, and a filter left
    # on from the last search is how you conclude an object does not exist.
    $grid.Columns['Type'].HeaderText = if (@($script:FilterType).Count)   { 'Type  (filtered)' }   else { 'Type' }
    $grid.Columns['Src'].HeaderText  = if (@($script:FilterSource).Count) { 'Source  (filtered)' } else { 'Source' }

    # THREE DIFFERENT CLAIMS, NOT ONE NUMBER.
    #
    # 'shown' is after filtering, 'found' is what the search returned, and
    # 'capped' says the search itself was truncated before either. Collapsing
    # them into "40 object(s)" is how a filtered view of an incomplete set reads
    # as a complete answer.
    $shown = $rows.Count
    $found = @($script:AllRows).Count
    $parts = @("$shown shown")
    if ($shown -ne $found) { $parts += "$found found" }
    if ($script:WasCapped) { $parts += "CAPPED at $($script:ActiveCap) per source -- there may be matches you cannot see" }

    $lblCount.ForeColor = if ($script:WasCapped) { [System.Drawing.Color]::Firebrick }
                          elseif ($shown -eq 0)  { [System.Drawing.Color]::Firebrick }
                          else                   { [System.Drawing.Color]::DimGray }
    $lblCount.Text = ($parts -join '  |  ') + $Note

    # The narrowed re-run only helps when BOTH are true: the search was cut off,
    # and there is a filter to push into the query. Offering it otherwise is a
    # button that re-runs the same search.
    $anyFilter = ((@($script:FilterSource).Count -gt 0) -or (@($script:FilterType).Count -gt 0))
    $btnClearFilter.Visible = $anyFilter
    $btnNarrow.Visible      = ($script:WasCapped -and $anyFilter)

    if ($grid.Rows.Count -gt 0) {
        $grid.ClearSelection()
        $grid.Rows[0].Selected = $true
    }
    Update-SelectedDetail
}

function Get-SelectedRow {
    if ($grid.SelectedRows.Count -eq 0) { return $null }
    # The object rides on the row's Tag, not on an index into a parallel array:
    # the grid is sortable, and sorting moves the Tag with the row.
    return $grid.SelectedRows[0].Tag
}

function Update-SelectedDetail {
    # A FUNCTION, NOT ONLY AN EVENT BODY. Rows.Add() selects the first row as a
    # side effect, so SelectionChanged fires DURING the fill -- before that
    # row's Tag is set. With exactly one match there is no other row to click
    # and the pane would stay empty. Same fault, same fix, as AD-Lookup.
    $r = Get-SelectedRow
    $txtDetail.Text = if ($r) { Format-UnifiedDetail -Row $r } else { '' }
    Update-HandOffButton -Row $r
}

function Update-HandOffButton {
    <#
        Caption and enabled state of the Open button, for the selected row.

        THE CAPTION NAMES THE DESTINATION. A button reading 'Open' that leads to
        Group-Manager-Editor for one row and AD-Lookup for the next is a button you
        have to click to find out what it does.

        WHEN NOTHING MATCHES IT GREYS OUT AND SAYS WHY. The alternative -- open
        something approximate -- is how you end up editing the cloud copy of a
        synced object and watching the change revert overnight.
    #>
    param($Row)

    if (-not $Row) {
        $btnOpen.Enabled = $false
        $btnOpen.Text    = 'Open'
        return
    }
    $h = Get-HandOff -Row $Row
    if ($h) {
        $btnOpen.Enabled = $true
        $btnOpen.Text    = "$($h.Label)"
    } else {
        $btnOpen.Enabled = $false
        $btnOpen.Text    = 'Open'
    }

    # ATTRIBUTES IS ENABLED FOR ANYTHING WITH SOMEWHERE TO WRITE.
    #
    # Cloud rows need a type this tool knows how to edit; on-premises rows go to
    # AD-Lookup, which edits anything with a distinguished name. Soft-deleted
    # rows are neither -- a deleted mailbox cannot be edited at all, and offering
    # it produces a "not found" that reads like the tool is broken.
    if ("$($Row.Source)" -eq 'Soft-deleted') {
        $btnAttrs.Enabled = $false
        $btnMbx.Enabled   = $false
        return
    }

    if ("$($Row.Source)" -eq 'Cloud') {
        $btnAttrs.Enabled = [bool](Get-CloudEditPlan -Row $Row)
    } else {
        $btnAttrs.Enabled = [bool]("$($Row.Dn)".Trim())
    }

    # INDEPENDENT OF Attributes AND OF THE SOURCE COLUMN. A synced
    # RemoteSharedMailbox enables BOTH: its directory object is on-premises and
    # its mailbox is in the cloud.
    $btnMbx.Enabled = [bool](Get-MailboxSettingsPlan -Row $Row)
}

function Copy-SelectedAddress {
    $r = Get-SelectedRow
    if (-not $r) { return }
    if (-not "$($r.Primary)".Trim()) {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text      = 'That object has no address -- nothing to copy.'
        return
    }
    try {
        [System.Windows.Forms.Clipboard]::SetText("$($r.Primary)")
        $lblStatus.ForeColor = [System.Drawing.Color]::ForestGreen
        $lblStatus.Text      = "$($r.Primary) copied."
    } catch {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text      = "Could not copy: $($_.Exception.Message)"
    }
}

$btnSearch.Add_Click({
    $text = $txtSearch.Text.Trim()
    if ($text.Length -lt 2) {
        $lblCount.ForeColor = [System.Drawing.Color]::Firebrick
        $lblCount.Text = 'Type at least two characters.'
        return
    }

    $btnSearch.Enabled = $false
    $form.Cursor       = [System.Windows.Forms.Cursors]::WaitCursor
    Update-ConnectionBanner
    $grid.Rows.Clear()
    $txtDetail.Text = ''
    $lblStatus.Text = ''
    $lblCount.ForeColor = [System.Drawing.Color]::DimGray
    $lblCount.Text  = 'Searching...'
    $form.Refresh()

    try {
        $all    = New-Object System.Collections.ArrayList
        $errs   = New-Object System.Collections.ArrayList
        $used   = New-Object System.Collections.ArrayList
        $capped = $false

        # EACH SOURCE IS INDEPENDENT. One that cannot answer is reported and the
        # rest still return -- a unified search that fails entirely because the
        # least important of four sources threw would be worse than the four
        # separate tools it replaces.
        #
        # NOTHING IS SKIPPED BECAUSE OF A PROBE ANY MORE. Read the Environment
        # region for why that was wrong. What is left is a real dependency
        # (Exchange Online genuinely cannot be searched without a session) and
        # it is now offered as a connection rather than announced as a limit.
        # A NARROWED RE-RUN, IF THE NARROW BUTTON ASKED FOR ONE. Consumed here
        # and cleared immediately, so a later plain Search is never silently
        # constrained by a choice made two searches ago.
        $narrowSource = @($script:NarrowSource)
        $narrowType   = @($script:NarrowType)
        $script:NarrowSource = @()
        $script:NarrowType   = @()
        $isNarrow = (($narrowSource.Count -gt 0) -or ($narrowType.Count -gt 0))
        $script:ActiveCap = if ($isNarrow) { $script:MaxResultsNarrow } else { $script:MaxResults }

        # SOURCE NARROWING SKIPS WHOLE FETCHES -- the cheapest kind. If only
        # 'Cloud' is wanted there is no reason to ask Active Directory at all.
        $wantOnPrem = ($narrowSource.Count -eq 0) -or ($narrowSource -contains 'On-prem') -or ($narrowSource -contains 'Synced') -or ($narrowSource -contains 'AD only')
        $wantCloud  = ($narrowSource.Count -eq 0) -or ($narrowSource -contains 'Cloud')  -or ($narrowSource -contains 'Synced')
        $wantDel    = ($narrowSource.Count -eq 0) -or ($narrowSource -contains 'Soft-deleted')

        $emsHere   = Test-OnPremLoaded
        $cloudHere = Test-CloudLoaded

        # Active Directory ALWAYS runs, and it carries the on-premises half on
        # its own when EMS is not loaded. -ExchangeAnswered tells it to stand
        # down to non-mail objects when EMS did the mail-enabled ones better.
        if ($emsHere -and $wantOnPrem) {
            $r = Get-OnPremRecipient -Text $text
            foreach ($x in $r.Rows) { [void]$all.Add($x) }
            foreach ($e in $r.Errors) { [void]$errs.Add($e) }
            if ($r.Rows.Count -gt $script:ActiveCap) { $capped = $true }
            [void]$used.Add('on-prem Exchange')
        }

        if ($wantOnPrem) {
            # -IncludeNonMail always -- see the note where the old checkbox was.
            $r = Get-OnPremAdObject -Text $text -ExchangeAnswered:$emsHere -IncludeNonMail
            foreach ($x in $r.Rows) { [void]$all.Add($x) }
            foreach ($e in $r.Errors) { [void]$errs.Add($e) }
            if ($r.Rows.Count -gt $script:ActiveCap) { $capped = $true }
            [void]$used.Add($(if ($emsHere) { 'Active Directory' } else { 'Active Directory (carrying the on-premises half)' }))
        }

        if ($cloudHere -and $wantCloud) {
            $r = Get-CloudRecipient -Text $text
            foreach ($x in $r.Rows) { [void]$all.Add($x) }
            foreach ($e in $r.Errors) { [void]$errs.Add($e) }
            if ($r.Rows.Count -gt $script:ActiveCap) { $capped = $true }
            [void]$used.Add('Exchange Online')

            if ($chkDeleted.Checked -and $wantDel) {
                $r = Get-SoftDeletedMailbox -Text $text
                foreach ($x in $r.Rows) { [void]$all.Add($x) }
                foreach ($e in $r.Errors) { [void]$errs.Add($e) }
                [void]$used.Add('soft-deleted')
            }
        } else {
            [void]$errs.Add('Exchange Online -- not connected. Use Connect ExO in the console, then search again.')
        }

        $merged = @(Merge-LookupRow -Rows @($all.ToArray()))

        # TYPE NARROWING IS APPLIED HERE, NOT IN THE QUERY, and that is a real
        # limit worth stating. The Type column is partly DERIVED -- see
        # Resolve-AdObjectType -- so Exchange and AD cannot filter on it. What
        # the re-run buys for a type narrow is the LARGER CEILING, not a cheaper
        # fetch; the source narrow is what actually skips work.
        if ($narrowType.Count -gt 0) {
            $merged = @($merged | Where-Object { $narrowType -contains "$(Get-RowTypeText -Row $_)" })
        }
        $script:AllRows    = @($merged)
        $script:WasCapped  = $capped
        $script:LastSearch = $text

        # FILTERS ARE CLEARED BY A NEW SEARCH, deliberately.
        #
        # A type filter left over from the last lookup is the single easiest way
        # to conclude an object does not exist. The header says '(filtered)'
        # while one is set, but a stale filter surviving a fresh search would
        # mean reading that header before trusting any empty result -- which is
        # a rule nobody remembers at four o'clock.
        $script:FilterSource = @()
        $script:FilterType   = @()

        $note = if ($errs.Count) { '   NOT SEARCHED: ' + ($errs -join '; ') } else { '' }
        if (@($script:AllRows).Count -eq 0) {
            $note = "  Nothing matched '$text'.  Searched: $($used -join ', ').$note"
        } else {
            $note = "  Searched: $($used -join ', ').$note"
        }
        Update-ResultGrid -Note $note
    } catch {
        $txtDetail.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Unified Lookup'
        $lblCount.ForeColor = [System.Drawing.Color]::Firebrick
        $lblCount.Text = 'Search failed -- see below.'
    } finally {
        $btnSearch.Enabled = $true
        $form.Cursor       = [System.Windows.Forms.Cursors]::Default
    }
})

# CLICK THE HEADER TO FILTER. Only Source and Type -- the two columns whose
# values form a short, closed set. A filter on Name or Address would be a search
# box wearing a different hat, and there is one of those at the top already.
$grid.Add_ColumnHeaderMouseClick({
    param($sender, $e)
    # 'Src', NOT 'Source' -- the column's NAME is Src and only its HEADER TEXT
    # is 'Source'. Comparing against the header text made this handler return
    # early on every Source click, so the column silently would not filter,
    # while the matching $grid.Columns['Source'] lookup in Update-ResultGrid
    # returned $null and threw on .HeaderText. One wrong assumption, two bugs,
    # and only the second one announced itself.
    $col = $grid.Columns[$e.ColumnIndex].Name
    if ($col -ne 'Src' -and $col -ne 'Type') { return }
    if (@($script:AllRows).Count -eq 0) { return }

    if ($col -eq 'Src') {
        $vals = @($script:AllRows | ForEach-Object { "$($_.Source)" })
        $new  = Show-ColumnFilter -Title 'Source' -Values $vals -Selected $script:FilterSource -Owner $form
        if ($new -and -not $new.Cancelled) { $script:FilterSource = @($new.Values); Update-ResultGrid }
    } else {
        $vals = @($script:AllRows | ForEach-Object { Get-RowTypeText -Row $_ })
        $new  = Show-ColumnFilter -Title 'Type' -Values $vals -Selected $script:FilterType -Owner $form
        if ($new -and -not $new.Cancelled) { $script:FilterType = @($new.Values); Update-ResultGrid }
    }
})

$grid.Add_SelectionChanged({ Update-SelectedDetail })
$grid.Add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    Copy-SelectedAddress
})
$btnCopy.Add_Click({ Copy-SelectedAddress })
$btnCopyDetail.Add_Click({
    if (-not $txtDetail.Text) { return }
    try {
        [System.Windows.Forms.Clipboard]::SetText($txtDetail.Text)
        $lblStatus.ForeColor = [System.Drawing.Color]::ForestGreen
        $lblStatus.Text      = 'Detail copied.'
    } catch {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text      = "Could not copy: $($_.Exception.Message)"
    }
})
$btnOpen.Add_Click({
    $r = Get-SelectedRow
    if (-not $r) { return }
    $h = Get-HandOff -Row $r
    if (-not $h) {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text      = "Nothing is wired up for a $(Get-RowTypeText -Row $r) in the $($r.Source) column yet."
        return
    }

    # THE IDENTITY IS THE ADDRESS, FALLING BACK TO THE NAME. Every destination
    # tool takes -SmtpAddress, and for an object with no address (an 'AD only'
    # security group) the name is what AD-Lookup can find it by.
    $ident = "$($r.Primary)".Trim()
    if (-not $ident) { $ident = "$($r.Name)".Trim() }
    if (-not $ident) {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text      = 'That row has neither an address nor a name to hand over.'
        return
    }

    $path = Join-Path $PSScriptRoot $h.Script
    try {
        if ($h.Where -eq 'OnPrem') {
            Start-ToolInOnPremProcess -ScriptPath $path -SmtpAddress $ident -Extra $h.Extra
            $lblStatus.ForeColor = [System.Drawing.Color]::ForestGreen
            $lblStatus.Text      = "$($h.Script) opening in its own window -- it connects on-premises first, so give it a moment."
        } else {
            # IN-PROCESS, because this window's process is the one holding the
            # Exchange Online session. A child process would have to sign in
            # again to do the same work -- and could not be given this one's
            # session, which is why Group-Manager says "the cloud side is ALWAYS
            # in-process, whatever the marker says".
            #
            # CALLED DIRECTLY, AS A NESTED DIALOG. This used to defer: record
            # the tool, close this window, and open it after ShowDialog
            # returned. That came from Group-Manager's comment that "a modal form
            # disables the other windows on its thread, so a tool opened while
            # this one is up would draw and refuse input" -- which I took as a
            # constraint on this file without checking whether it applied.
            #
            # IT DOES NOT. Show-CloudAttributeEditor in this same file opens a
            # nested modal with ShowDialog($Owner) while this window is up, and
            # it works. A nested ShowDialog creates its own message loop; the
            # owner is what makes the nesting well-behaved, and ShowDialog()
            # with no argument adopts the active window as owner.
            #
            # The cost of getting it wrong was real and daily: opening
            # permissions on one object closed the lookup, so finding the next
            # object meant starting the whole search again.
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
            $lblStatus.Text      = "Opening $($h.Script)..."
            $form.Refresh()
            try {
                $callArgs = @{ SmtpAddress = $ident }
                foreach ($x in @($h.Extra)) {
                    if ("$x".Trim()) { $callArgs["$x".TrimStart('-')] = $true }
                }
                & $path @callArgs
                $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
                $lblStatus.Text      = "$($h.Script) closed. Search again to refresh the grid."
            } finally {
                $form.Cursor = [System.Windows.Forms.Cursors]::Default
            }
        }
    } catch {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text      = "Could not open $($h.Script): $($_.Exception.Message)"
    }
})

$btnClearFilter.Add_Click({
    $script:FilterSource = @()
    $script:FilterType   = @()
    Update-ResultGrid
})

$btnNarrow.Add_Click({
    # THE FILTER BECOMES THE QUERY.
    #
    # Client-side filtering hides rows; it cannot recover rows the search never
    # returned. When a source hit its ceiling, the filtered view is a tidy
    # subset of an incomplete set -- and it looks complete, which is the danger.
    #
    # This re-runs the SAME text with the current filters applied as real
    # constraints, so the (larger) ceiling is spent only on the type and side
    # that were asked for. 'Show me cloud user mailboxes matching <word>'
    # returns 2000 of those rather than 500 of everything.
    #
    # The filters are carried across the re-run rather than cleared, because the
    # operator chose them one second ago and clearing them here would discard
    # the entire point of the button.
    $keepSource = @($script:FilterSource)
    $keepType   = @($script:FilterType)

    $script:NarrowSource = $keepSource
    $script:NarrowType   = $keepType
    $btnSearch.PerformClick()

    $script:FilterSource = $keepSource
    $script:FilterType   = $keepType
    Update-ResultGrid
})

$btnAttrs.Add_Click({
    $r = Get-SelectedRow
    if (-not $r) { return }

    if ("$($r.Source)" -eq 'Cloud') {
        $plan = Get-CloudEditPlan -Row $r
        if (-not $plan) {
            $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
            $lblStatus.Text      = "No verified cloud editor for a $(Get-RowTypeText -Row $r)."
            return
        }
        Show-CloudAttributeEditor -Row $r -Plan $plan -Title 'Cloud attributes' -Owner $form
        # The grid still shows what the search returned, which may now be stale.
        # Saying so is better than silently displaying an old display name.
        $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
        $lblStatus.Text      = 'Cloud editor closed. Search again to refresh the grid.'
        return
    }

    # ON-PREMISES, SYNCED AND AD-ONLY GO TO AD-LOOKUP. It is the attribute
    # editor for anything with a distinguished name, and for a SYNCED object it
    # is the CORRECT side to edit -- Active Directory is the authority and a
    # cloud-side change would revert on the next sync.
    $ident = "$($r.Primary)".Trim()
    if (-not $ident) { $ident = "$($r.Name)".Trim() }
    $p = Join-Path $PSScriptRoot 'AD-Lookup.ps1'
    if (-not (Test-Path -LiteralPath $p)) {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text      = 'AD-Lookup.ps1 is not in this folder, so on-premises attributes cannot be edited.'
        return
    }

    # NESTED, NOT DEFERRED -- see the Open handler for the full reasoning. This
    # path had the same fault and it was less visible: Attributes on a CLOUD
    # object opens a nested dialog and comes back, while Attributes on an
    # on-premises object closed the whole tool. One button, two behaviours,
    # depending on a property of the object rather than on anything the
    # operator did.
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
    $lblStatus.Text      = 'Opening AD Lookup...'
    $form.Refresh()
    try {
        & $p -SmtpAddress $ident
        $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
        $lblStatus.Text      = 'AD Lookup closed. Search again to refresh the grid.'
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnMbx.Add_Click({
    $r = Get-SelectedRow
    if (-not $r) { return }
    $plan = Get-MailboxSettingsPlan -Row $r
    if (-not $plan) {
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick
        $lblStatus.Text      = 'That object has no mailbox in Exchange Online.'
        return
    }

    # ALWAYS THE CLOUD DIALOG, EVEN FOR A SYNCED ROW. The mailbox is in Exchange
    # Online regardless of where the directory object lives -- that is the whole
    # reason this button is separate from Attributes.
    Show-CloudAttributeEditor -Row $r -Plan $plan -Title 'Mailbox settings' -Owner $form

    # The cached forwarding line on the row is now stale. Dropping it means the
    # next click re-reads it rather than showing what was true before the edit.
    if ($r.PSObject.Properties['ForwardText']) { $r.PSObject.Properties.Remove('ForwardText') }
    Update-SelectedDetail

    $lblStatus.ForeColor = [System.Drawing.Color]::DimGray
    $lblStatus.Text      = 'Mailbox settings closed.'
})

$btnClose.Add_Click({ $form.Close() })


$form.AcceptButton = $btnSearch

Update-ConnectionBanner

Set-ConsoleTheme -Form $form
Set-FormResizable -Form $form
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Unified Lookup' | Out-Null

[void]$form.ShowDialog()

# NOTHING RUNS AFTER THE DIALOG ANY MORE.
#
# There used to be a deferred hand-off block here: the Open and Attributes
# buttons recorded a tool, closed this window, and it was launched once
# ShowDialog returned. Every in-process hand-off is now a nested dialog called
# straight from its click handler, so this window stays open and the operator
# comes back to their search results.
#
# The removed block is worth one line of history because the behaviour it caused
# was easy to mistake for intent: opening permissions on one object ended the
# session, so looking at a second object meant retyping the search.
#endregion
```


<br>

==============================================================================

# ◆ FILE 10 OF 16 — `Group-Manager-Editor.ps1`

> **1246 lines.** THE WRITE REFERENCE. Directory resolution before a write, identity binding on GUID, the on-prem/cloud refusal guard, audit rows on success AND failure.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```powershell
#Requires -Version 5.1
# TOOLKIT-REQUIRES: OwnProcess  -- the -OnPrem side uses Add-/Remove-ADPermission and
#                                 on-premises Get-/Set-DistributionGroup, which share
#                                 cmdlet names with Exchange Online. Group-Manager
#                                 overrides this for the CLOUD side, which must stay in
#                                 the console process to keep its Exchange Online session.
# TOOLKIT-VARIANT: Dual
<#
    Group-Manager-Editor.ps1
    Everything you do to one distribution group, in one window.

    Opened by Group-Manager.ps1, which reads IsDirSynced and passes -OnPrem when
    Active Directory is the authority. Not in Tools.json: it is a helper with a
    front door, not a front door itself.

    Reasoning: ScriptNotes/Group-Manager-Editor.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-08-19
#>
# TOOLKIT-HELP-BEGIN
# One group, every setting, one window.
#
# The title bar says [CLOUD] or [ON-PREMISES]. That is decided by DL Manager
# from the group itself, not by you, and it is the difference between a change
# that takes effect and one that is silently discarded.
#
# FIVE SECTIONS, EACH APPLIED SEPARATELY. Nothing is written until you press the
# Apply or Process button in that section, and each one confirms first.
#
#   Owner            one person at a time. Owners matter more than they look:
#                    a group with no owner can only be managed by an admin, and
#                    nobody is told when the last one leaves.
#
#   Membership       paste names into Add and Remove and process both at once.
#                    Adds run before removals -- see below.
#
#   Delivery         hidden from the address book, whether external senders are
#                    accepted, and the allowed-senders list.
#
#   Send-As          who may send mail that appears to come FROM the list.
#
# THE ONE THING WORTH READING TWICE. An allowed-senders list does nothing on its
# own. It is only enforced when 'Require sender authentication' is ON -- with it
# off, the list is stored and ignored, and anyone can mail the group. This
# window states which of the two situations you are actually in rather than
# leaving you to work it out from two unrelated checkboxes.
#
# There is no undo. Removals are recoverable only by adding people back, and
# the audit log is what tells you who was there.
# TOOLKIT-HELP-END
# TOOLKIT-HELP-RUNS: One distribution group, in the environment named in the title bar.

param(
    [string]$SmtpAddress,

    # WRITE ON-PREMISES INSTEAD OF TO EXCHANGE ONLINE.
    #
    # Set by Group-Manager for a directory-synced group. The caller decides,
    # because the caller is the one that read IsDirSynced. This tool must never
    # guess, and must never fall back to the other side when one fails.
    [switch]$OnPrem
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

. (Join-Path $PSScriptRoot 'Common.ps1')

# Launched from the console, or typed at a prompt? See Test-ConsoleLaunch in
# Common.ps1 for what this is and -- more importantly -- what it is not.
if (-not (Test-ConsoleLaunch)) { return }

#region Environment ------------------------------------------------------------
$script:OnPremMode = [bool]$OnPrem

# Filled by Resolve-DLEnvironment in on-premises mode. Empty in cloud mode, and
# splatting an empty hashtable adds nothing, so every call site stays identical.
$script:DCParam = @{}

# The group's distinguished name, filled by Get-DLState. See Get-SendAsTarget.
$script:GroupDN = ''

function Get-SendAsTarget {
    <#
        WHICH IDENTITY THE SEND-AS CMDLETS WANT. It differs by side, and that is
        not a style choice -- they are different KINDS of cmdlet.

        Get-/Add-/Remove-RecipientPermission are RECIPIENT cmdlets. Their
        -Identity accepts an email address, which is what the rest of this
        window uses throughout.

        Get-/Add-/Remove-ADPermission are AD OBJECT cmdlets. Their -Identity
        takes a name, canonical name, distinguished name or GUID -- and NOT an
        SMTP address. Handed one, Exchange looks for an AD object literally
        named that and reports:

            The operation couldn't be performed because object
            'someone@contoso.com' couldn't be found on '<controller>'

        which reads as a missing group and is nothing of the kind. Every other
        section of the window read the same group successfully at the same
        moment, because every other section calls a recipient cmdlet.

        THE DISTINGUISHED NAME, NOT Identity. ARCHITECTURE 14g: a resolved
        object's Identity comes back as its AD common name -- no domain, no path
        -- and passing that back is a fresh forest-wide NAME LOOKUP rather than
        a bind, which can land on a different object than the one just read. A
        DN is unambiguous.

        Falls back to the address when the DN is somehow empty: failing with the
        old error beats failing with a null identity.
    #>
    if ($script:OnPremMode -and "$script:GroupDN".Trim()) { return "$script:GroupDN" }
    return $script:Identity
}

function Invoke-DLCmd {
    <#
        Runs an Exchange cmdlet against whichever environment this instance is
        for. THE ONLY PLACE THAT KNOWS THE DIFFERENCE.

        Cloud goes through Invoke-EXOCommand, which resolves explicitly to the
        Exchange Online module, so a stray on-premises session in this process
        cannot capture Get-/Set-DistributionGroup.

        On-premises calls the cmdlet directly with the domain controller
        splatted in. It cannot use Invoke-EXOCommand for the obvious reason, and
        does not need the same protection for the less obvious one: the tool
        refuses to start at all when both environments are present.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [hashtable]$Parameters = @{}
    )

    if (-not $script:OnPremMode) {
        return Invoke-EXOCommand -Name $Name -Parameters $Parameters
    }

    # A COPY, not the caller's hashtable. Adding the controller to the original
    # would mutate a table the caller may reuse for a second call -- and against
    # a cmdlet that does not accept -DomainController, that fails confusingly.
    $p = @{}
    foreach ($k in $Parameters.Keys) { $p[$k] = $Parameters[$k] }
    foreach ($k in $script:DCParam.Keys) { $p[$k] = $script:DCParam[$k] }
    return & $Name @p
}

function Get-SamFromUserString {
    # 'CONTOSO\bxrubble' -> 'bxrubble'. Get-ADPermission reports the trustee in
    # DOMAIN\user form, which is not what Add-ADPermission wants back.
    param([string]$UserString)
    $s = "$UserString"
    if ($s -match '\\') { return ($s -split '\\')[-1] }
    return $s
}

function Resolve-DLEnvironment {
    <#
        Confirms the group exists, and on-premises works out WHICH DIRECTORY
        holds it before anything is written.

        Directories come from Tenant.json, so a tenant with one gets one and a
        tenant with three gets three, with no directory names written down here.
        The FIRST directory that resolves the group wins, and its controller is
        pinned into $script:DCParam for every later call -- so the read and the
        write cannot land in different forests.

        On a miss the controller is CLEARED rather than left as it was. A stale
        one would send the next write somewhere nobody asked for.
    #>
    param([Parameter(Mandatory)][string]$Identity)

    $errors = @()

    if (-not $script:OnPremMode) {
        try {
            $g = Invoke-DLCmd -Name 'Get-DistributionGroup' -Parameters @{
                Identity = $Identity; ErrorAction = 'Stop' }
            return [PSCustomObject]@{ Found = $true; Label = 'Exchange Online'; Errors = $errors }
        } catch { $errors += $_.Exception.Message }
        return [PSCustomObject]@{ Found = $false; Label = $null; Errors = $errors }
    }

    # A REAL CONTROLLER, OR NONE AT ALL -- NEVER A DOMAIN.
    #
    # This called Get-TenantDirectoryServer, which DELIBERATELY returns the
    # DOMAIN when a directory names several controllers (ARCHITECTURE 14f, "do
    # not pin"). That is correct for the ActiveDirectory module's -Server and
    # fatal for Exchange's -DomainController, which runs a suitability check and
    # rejects a domain with:
    #
    #     Active Directory error 0x52 ... suitability of server ... A local
    #     error occurred
    #
    # Common.ps1 states the rule outright -- an Exchange cmdlet must be handed a
    # controller from the directory resolver, or nothing at all -- and I crossed
    # exactly the seam it was written to warn about. NewUser hit this same 0x52
    # on 2026-08-12.
    #
    # AND NOTHING AT ALL IS A REAL OPTION. Where no controller can be resolved,
    # the call goes forest-wide instead of being skipped: a wrong pin is worse
    # than no pin, because pinning one directory's controller for a write in
    # another fails as "out of the current user's write scope", which reads as
    # a permissions problem and is not one.
    foreach ($d in @(Get-TenantDirectory)) {
        $srv = ''
        try { $srv = Resolve-DirectoryDomainController -DirectoryId $d.Id } catch { $srv = '' }

        $probe = @{ Identity = $Identity; ErrorAction = 'Stop' }
        if ($srv) { $probe['DomainController'] = $srv }

        try {
            $null = Get-DistributionGroup @probe
            $script:DCParam = if ($srv) { @{ DomainController = $srv } } else { @{} }
            $how = if ($srv) { "$($d.Id) via $srv" } else { "$($d.Id) (AD locates a controller)" }
            return [PSCustomObject]@{ Found = $true; Label = $how; Errors = $errors }
        } catch {
            $where = if ($srv) { "$($d.Id) ($srv)" } else { "$($d.Id) (no controller resolved)" }
            $errors += "${where}: $($_.Exception.Message)"
        }
    }
    $script:DCParam = @{}
    return [PSCustomObject]@{ Found = $false; Label = $null; Errors = $errors }
}
#endregion

#region Parsing ----------------------------------------------------------------
function ConvertTo-EntryList {
    <#
        Splits a text box into a de-duplicated list.

        Newlines, commas and semicolons all separate, because a list pasted out
        of a ticket arrives in whichever the writer happened to use.

        Duplicates are dropped case-insensitively: the same person twice is not
        two operations, and the second would fail with "already a member" and be
        reported as an error the operator has to read and dismiss.
    #>
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    $seen = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
    $kept = New-Object System.Collections.ArrayList
    foreach ($line in ($Text -split '[\r\n,;]+')) {
        $v = "$line".Trim()
        if (-not $v) { continue }
        if ($seen.Add($v)) { [void]$kept.Add($v) }
    }
    return @($kept.ToArray())
}

function Resolve-RecipientKey {
    <#
        Turns whatever the operator typed into the identity the write cmdlets
        want back.

        An alias is completed to a full address first, because Get-Recipient
        will happily resolve 'barney' to something unexpected in a forest with
        several matches.

        RETURNS THE PRIMARY SMTP ADDRESS, not the display name and not Identity.
        Identity for a group comes back as its Active Directory CN -- no domain,
        no path -- and passing that back is a fresh forest-wide name lookup
        rather than a bind, which can land on a different object than the one
        just resolved.
    #>
    param([Parameter(Mandatory)][string]$Entry)

    $tries = @($Entry)
    if ($Entry -notlike '*@*') {
        $completed = Get-TargetSmtp -Text $Entry
        if ($completed -and $completed -ne $Entry) { $tries += $completed }
    }

    $lastErr = $null
    foreach ($t in $tries) {
        try {
            $r = Invoke-DLCmd -Name 'Get-Recipient' -Parameters @{ Identity = $t; ErrorAction = 'Stop' }
            return [PSCustomObject]@{
                Ok      = $true
                Input   = $Entry
                Key     = "$($r.PrimarySmtpAddress)"
                Sam     = "$($r.SamAccountName)"
                Display = "$($r.DisplayName) <$($r.PrimarySmtpAddress)>"
                Error   = $null
            }
        } catch { $lastErr = "$($_.Exception.Message)" }
    }
    return [PSCustomObject]@{
        Ok = $false; Input = $Entry; Key = $null; Sam = $null; Display = $null; Error = $lastErr
    }
}
#endregion

#region State ------------------------------------------------------------------
function Get-DLState {
    <#
        Everything this window shows, read in one pass.

        Errors are COLLECTED, not thrown. A group whose members cannot be listed
        is still worth showing with its delivery settings -- an editor that
        refuses to open teaches nobody anything, and the section that failed is
        named rather than the whole window going blank.
    #>
    param([Parameter(Mandatory)][string]$Identity)

    $state = [PSCustomObject]@{
        Identity     = $Identity
        DisplayName  = ''
        PrimarySmtp  = ''
        GroupType    = ''
        GroupDN      = ''
        Hidden       = $false
        RequireAuth  = $false
        Owners       = @()
        Members      = @()
        Accepted     = @()
        SendAs       = @()
        Errors       = @()
    }

    try {
        $g = Invoke-DLCmd -Name 'Get-DistributionGroup' -Parameters @{
            Identity = $Identity; ErrorAction = 'Stop' }
        $state.DisplayName = "$($g.DisplayName)"
        $state.PrimarySmtp = "$($g.PrimarySmtpAddress)"
        $state.GroupType   = "$($g.RecipientTypeDetails)"
        $state.Hidden      = [bool]$g.HiddenFromAddressListsEnabled
        $state.RequireAuth = [bool]$g.RequireSenderAuthenticationEnabled

        # Captured here because this is the only place the group object exists.
        # Get-SendAsTarget needs it, and re-reading the group to fetch one
        # property would be a second round trip for something already in hand.
        $state.GroupDN     = "$($g.DistinguishedName)"
        $script:GroupDN    = $state.GroupDN
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-DistributionGroup')
        return $state
    }

    foreach ($o in @($g.ManagedBy)) {
        $r = Resolve-RecipientKey -Entry "$o"
        if ($r.Ok) { $state.Owners += $r }
        else {
            # Reported, not dropped. An owner that no longer resolves is usually
            # a deleted account, and it still occupies a slot in ManagedBy --
            # quietly hiding it would make the owner list look shorter than the
            # one any write has to preserve.
            $state.Owners += [PSCustomObject]@{
                Ok = $false; Input = "$o"; Key = "$o"; Sam = $null
                Display = "$o (could not re-resolve -- may be deleted)"; Error = $r.Error }
        }
    }

    try {
        # -ResultSize Unlimited IS LOAD-BEARING. The default is 1000 and it
        # truncates SILENTLY -- a group with 1200 members would display 1000 and
        # look complete.
        $mem = @(Invoke-DLCmd -Name 'Get-DistributionGroupMember' -Parameters @{
            Identity = $Identity; ResultSize = 'Unlimited'; ErrorAction = 'Stop' })
        $state.Members = @($mem | ForEach-Object {
            "$($_.DisplayName) <$($_.PrimarySmtpAddress)>  [$($_.RecipientType)]" })
    } catch {
        $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Get-DistributionGroupMember')
    }

    foreach ($s in @($g.AcceptMessagesOnlyFromSendersOrMembers)) {
        $r = Resolve-RecipientKey -Entry "$s"
        if ($r.Ok) { $state.Accepted += $r }
        else {
            $state.Accepted += [PSCustomObject]@{
                Ok = $false; Input = "$s"; Key = "$s"; Sam = $null
                Display = "$s (could not re-resolve -- may be stale)"; Error = $r.Error }
        }
    }

    # NOT $Identity -- see Get-SendAsTarget. On-premises this must be the DN.
    try { $state.SendAs = @(Get-SendAsGrant -Identity (Get-SendAsTarget)) }
    catch { $state.Errors += (Format-ErrorDetail -ErrorRecord $_ -Context 'Send-As read') }

    return $state
}

function Get-SendAsGrant {
    <#
        Who holds Send-As. The two environments model this differently and the
        shapes do not match, so both are normalised here.

        Exchange Online has Get-RecipientPermission, whose trustee is an
        ADDRESS. On-premises has Get-ADPermission, whose User is DOMAIN\sam and
        whose rights are a collection to be searched rather than a string to
        match.

        NT AUTHORITY\SELF is on every object -- it is the platform's own grant,
        not somebody's delegation -- so it is filtered rather than displayed.

        Deny entries are KEPT and flagged. A Deny outranks an Allow, so hiding
        one would mean reporting access that does not actually work.
    #>
    param([Parameter(Mandatory)][string]$Identity)
    $out = @()

    if ($script:OnPremMode) {
        $perms = @(Invoke-DLCmd -Name 'Get-ADPermission' -Parameters @{
            Identity = $Identity; ErrorAction = 'Stop' })
        foreach ($p in $perms) {
            $who = "$($p.User)"
            if ($who -like '*\SELF') { continue }
            if ($p.IsInherited) { continue }
            if (-not (@($p.ExtendedRights) -match 'Send-As')) { continue }
            $out += [PSCustomObject]@{
                Display = $who; Key = (Get-SamFromUserString -UserString $who); Deny = [bool]$p.Deny }
        }
    } else {
        $perms = @(Invoke-DLCmd -Name 'Get-RecipientPermission' -Parameters @{
            Identity = $Identity; ErrorAction = 'Stop' })
        foreach ($p in $perms) {
            $trustee = "$($p.Trustee)"
            if ($trustee -match 'NT AUTHORITY\\SELF') { continue }
            if ("$($p.AccessRights)" -notmatch 'SendAs') { continue }
            $out += [PSCustomObject]@{
                Display = $trustee; Key = $trustee; Deny = ("$($p.AccessControlType)" -eq 'Deny') }
        }
    }
    return @($out | Sort-Object Display)
}

function Format-DLState {
    param([Parameter(Mandatory)]$State)
    $sb = New-Object System.Text.StringBuilder

    [void]$sb.AppendLine("Group           : $($State.DisplayName)")
    [void]$sb.AppendLine("Address         : $($State.PrimarySmtp)")
    [void]$sb.AppendLine("Type            : $($State.GroupType)")
    [void]$sb.AppendLine("Hidden from GAL : $($State.Hidden)")
    [void]$sb.AppendLine('')

    # THE TWO SETTINGS ARE REPORTED TOGETHER, because one governs the other.
    #
    # An allowed-senders list is only enforced when sender authentication is on.
    # Printing them as two independent lines invites exactly the mistake this
    # tool exists to prevent: carefully building a list that is being ignored.
    if ($State.RequireAuth) {
        [void]$sb.AppendLine("Delivery        : INTERNAL SENDERS ONLY (sender authentication ON)")
        if ($State.Accepted.Count -gt 0) {
            [void]$sb.AppendLine("                  and the allowed-senders list below IS enforced.")
        } else {
            [void]$sb.AppendLine("                  no allowed-senders list, so any internal sender may post.")
        }
    } else {
        [void]$sb.AppendLine("Delivery        : EXTERNAL SENDERS ACCEPTED (sender authentication OFF)")
        if ($State.Accepted.Count -gt 0) {
            [void]$sb.AppendLine("                  WARNING: the allowed-senders list below is NOT enforced.")
            [void]$sb.AppendLine("                  It is stored and ignored -- anyone may mail this group.")
        }
    }
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine("Owners ($($State.Owners.Count)):")
    if ($State.Owners.Count -eq 0) {
        [void]$sb.AppendLine('  (none -- only an administrator can manage this group)')
    } else {
        $State.Owners | ForEach-Object { [void]$sb.AppendLine("  $($_.Display)") }
    }
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine("Allowed senders ($($State.Accepted.Count)):")
    if ($State.Accepted.Count -eq 0) {
        [void]$sb.AppendLine('  (none -- not restricted to specific senders)')
    } else {
        $State.Accepted | ForEach-Object { [void]$sb.AppendLine("  $($_.Display)") }
    }
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine("Send-As grants ($($State.SendAs.Count)), SELF excluded:")
    if ($State.SendAs.Count -eq 0) {
        [void]$sb.AppendLine('  (none)')
    } else {
        $State.SendAs | ForEach-Object {
            $tag = if ($_.Deny) { '  [DENY]' } else { '' }
            [void]$sb.AppendLine("  $($_.Display)$tag")
        }
    }
    [void]$sb.AppendLine('')

    [void]$sb.AppendLine("Members ($($State.Members.Count)):")
    if ($State.Members.Count -eq 0) {
        [void]$sb.AppendLine('  (none)')
    } else {
        $State.Members | ForEach-Object { [void]$sb.AppendLine("  $_") }
    }

    if ($State.Errors.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('--- Field errors (everything above is still valid) ---')
        $State.Errors | ForEach-Object { [void]$sb.AppendLine($_) }
    }
    return $sb.ToString()
}
#endregion

#region Writes -----------------------------------------------------------------

# -BypassSecurityGroupManagerCheck ON EVERY WRITE.
#
# Without it Exchange refuses with OperationRequiresGroupManagerException --
# "this operation can only be performed by a manager of the group" -- which
# reads as a missing admin right and is not one. Exchange checks whether YOU
# are listed in the group's ManagedBy, regardless of what your RBAC role says,
# and an administrator editing somebody else's group almost never is.
#
# The rest of the toolkit already does this: ManageOwnerGroups, ManageMailObject,
# Bulk-DLMembership and six project scripts all pass it. This file was the only
# writer that did not.
#
# It applies to Set-DistributionGroup AND to the member cmdlets, so it goes on
# all five paths -- owner, accepted senders, the two flags, and add/remove
# member. Fixing only the owner call would have moved the same error one button
# to the right.
function Set-DLOwners {
    <#
        ManagedBy IS REPLACED WHOLE, never appended to.

        Set-DistributionGroup -ManagedBy takes the complete list, so adding one
        owner means writing every existing owner back alongside them. Getting
        this wrong does not error -- it silently removes everybody else.
    #>
    param([Parameter(Mandatory)][string]$Identity, [string[]]$Owners)
    Invoke-DLCmd -Name 'Set-DistributionGroup' -Parameters @{
        Identity = $Identity; ManagedBy = @($Owners)
        BypassSecurityGroupManagerCheck = $true
        Confirm = $false; ErrorAction = 'Stop' } | Out-Null
}

function Set-DLAcceptedSenders {
    # Same whole-list replacement as ManagedBy, for the same reason.
    param([Parameter(Mandatory)][string]$Identity, [string[]]$Senders)
    Invoke-DLCmd -Name 'Set-DistributionGroup' -Parameters @{
        Identity = $Identity; AcceptMessagesOnlyFromSendersOrMembers = @($Senders)
        BypassSecurityGroupManagerCheck = $true
        Confirm = $false; ErrorAction = 'Stop' } | Out-Null
}

function Set-DLFlag {
    param(
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][ValidateSet('Hidden', 'RequireAuth')][string]$Flag,
        [Parameter(Mandatory)][bool]$Value
    )
    $p = @{ Identity = $Identity; BypassSecurityGroupManagerCheck = $true
            Confirm = $false; ErrorAction = 'Stop' }
    if ($Flag -eq 'Hidden') { $p['HiddenFromAddressListsEnabled'] = $Value }
    else                    { $p['RequireSenderAuthenticationEnabled'] = $Value }
    Invoke-DLCmd -Name 'Set-DistributionGroup' -Parameters $p | Out-Null
}

function Add-DLMemberEntry {
    param([Parameter(Mandatory)][string]$Identity, [Parameter(Mandatory)][string]$Member)
    Invoke-DLCmd -Name 'Add-DistributionGroupMember' -Parameters @{
        Identity = $Identity; Member = $Member
        BypassSecurityGroupManagerCheck = $true
        Confirm = $false; ErrorAction = 'Stop' } | Out-Null
}

function Remove-DLMemberEntry {
    param([Parameter(Mandatory)][string]$Identity, [Parameter(Mandatory)][string]$Member)
    Invoke-DLCmd -Name 'Remove-DistributionGroupMember' -Parameters @{
        Identity = $Identity; Member = $Member
        BypassSecurityGroupManagerCheck = $true
        Confirm = $false; ErrorAction = 'Stop' } | Out-Null
}

function Add-DLSendAs {
    # On-premises uses -ExtendedRights 'Send-As' WITH THE HYPHEN; the cloud
    # right is 'SendAs' without one. They are not interchangeable and neither
    # errors usefully when given the other's spelling.
    param([Parameter(Mandatory)][string]$Identity, [Parameter(Mandatory)][string]$User)
    if ($script:OnPremMode) {
        Invoke-DLCmd -Name 'Add-ADPermission' -Parameters @{
            Identity = $Identity; User = $User; ExtendedRights = 'Send-As'
            Confirm = $false; ErrorAction = 'Stop' } | Out-Null
    } else {
        Invoke-DLCmd -Name 'Add-RecipientPermission' -Parameters @{
            Identity = $Identity; Trustee = $User; AccessRights = 'SendAs'
            Confirm = $false; ErrorAction = 'Stop' } | Out-Null
    }
}

function Remove-DLSendAs {
    param([Parameter(Mandatory)][string]$Identity, [Parameter(Mandatory)][string]$User)
    if ($script:OnPremMode) {
        Invoke-DLCmd -Name 'Remove-ADPermission' -Parameters @{
            Identity = $Identity; User = $User; ExtendedRights = 'Send-As'
            Confirm = $false; ErrorAction = 'Stop' } | Out-Null
    } else {
        Invoke-DLCmd -Name 'Remove-RecipientPermission' -Parameters @{
            Identity = $Identity; Trustee = $User; AccessRights = 'SendAs'
            Confirm = $false; ErrorAction = 'Stop' } | Out-Null
    }
}
#endregion

#region Target -----------------------------------------------------------------
$script:Identity = Get-TargetSmtp -Text $SmtpAddress
if (-not $script:Identity) {
    [void][System.Windows.Forms.MessageBox]::Show(
        "No distribution group was supplied.`r`n`r`nOpen this from DL Manager, which looks the group up and decides which environment it belongs to.",
        'DL Editor', [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning)
    return
}
$script:State = $null
#endregion

#region WinForms UI ------------------------------------------------------------
$form               = New-Object System.Windows.Forms.Form
$form.Text          = if ($script:OnPremMode) { "DL Editor: $script:Identity  [ON-PREMISES]" }
                      else                    { "DL Editor: $script:Identity  [CLOUD]" }
# CAPPED AT 1100 AND THE INPUT BOXES GAVE UP THE HEIGHT, not the state view.
#
# A form taller than the screen scrolls as one slab, which moves the section
# you are reading as well as the box you are typing in. Each multi-line box
# scrolls on its own instead -- fifty names in the Add box no longer pushes the
# Send-As section off the bottom of the window.
# WIDTH IS FIXED TO THE CONTENT, and the maximum enforces it.
#
# Nothing on this form is anchored, so a window wider than 951 is dead space --
# the group boxes stop at 922 and the boxes inside them stop at 912 no matter
# how far the frame is dragged. MaximumSize used to allow 1400, which is 450px
# of nothing.
#
# 951 rather than 934: AutoScroll below puts a vertical scrollbar in the last
# ~17px of the client area, and content that runs under it earns a HORIZONTAL
# scrollbar as well.
$form.ClientSize    = New-Object System.Drawing.Size(951, 1066)
$form.MaximumSize   = New-Object System.Drawing.Size(967, 1105)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

# AUTOSCROLL, because this form is deliberately taller than some screens.
# Without it the Send-As section is simply unreachable on a laptop, which is a
# worse outcome than a scrollbar.
$form.AutoScroll    = $true

$lblWarn           = New-Object System.Windows.Forms.Label
$lblWarn.Tag       = 'NoTheme'
$lblWarn.Location  = New-Object System.Drawing.Point(12, 10)
$lblWarn.Size      = New-Object System.Drawing.Size(910, 34)
$lblWarn.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblWarn)

# THE GUARD IS ASYMMETRIC, AND I HAD IT SYMMETRIC. That was wrong.
#
# The reasoning was "the cmdlet names are shared, so either side loaded
# alongside the other could answer instead". True for ONE direction only.
#
# The cloud path calls everything through Invoke-EXOCommand, which resolves
# explicitly to the Exchange Online module. That is the entire reason that
# helper exists: a stray on-premises session in the same process CANNOT capture
# Get-/Set-DistributionGroup from it. So cloud mode is safe with EMS loaded --
# and blocking it broke the normal case, because the console connects to both
# on purpose and its header shows EMS and ExO green together.
#
# The on-premises path calls cmdlets DIRECTLY, unqualified, because there is no
# equivalent helper for the management shell. That one really would be captured
# by an Exchange Online session, so it keeps its refusal.
#
#   cloud    needs ExO. EMS alongside is fine.
#   on-prem  needs EMS. ExO alongside is fatal.
$emsState = Test-EmsAvailable
$exoState = Test-ExoInProcess
$script:Blocked = $false

if ($script:OnPremMode) {
    if ($exoState.Present) {
        $lblWarn.Text   = 'BLOCKED: Exchange Online is loaded in this process and would answer instead of on-premises Exchange. This must run in its own window.'
        $script:Blocked = $true
    } elseif (-not $emsState.Available) {
        $lblWarn.Text   = 'NOT connected to on-premises Exchange -- nothing can run. Use Connect ExSE.'
        $script:Blocked = $true
    } else {
        $lblWarn.Text = 'ON-PREMISES. Active Directory is the authority for this group; the directory is resolved from Tenant.json.'
    }
} else {
    if (-not $exoState.Present) {
        $lblWarn.Text   = 'NOT connected to Exchange Online -- nothing can run. Use Connect ExO.'
        $script:Blocked = $true
    } else {
        $lblWarn.Text = 'CLOUD. Every call is routed through the Exchange Online session explicitly.'
    }
}
$lblWarn.ForeColor = if ($script:Blocked) { [System.Drawing.Color]::Red }
                     else                 { [System.Drawing.Color]::FromArgb(80, 220, 130) }

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Location   = New-Object System.Drawing.Point(12, 50)
$txtOut.Size       = New-Object System.Drawing.Size(910, 300)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

# --- Owner ------------------------------------------------------------------
$grpOwner = New-Object System.Windows.Forms.GroupBox
$grpOwner.Text     = 'Owner'
$grpOwner.Location = New-Object System.Drawing.Point(12, 360)
$grpOwner.Size     = New-Object System.Drawing.Size(910, 92)
$form.Controls.Add($grpOwner)

$lblOwner          = New-Object System.Windows.Forms.Label
$lblOwner.Text     = 'Grantee (email or logon):'
$lblOwner.Location = New-Object System.Drawing.Point(10, 24)
$lblOwner.AutoSize = $true
$grpOwner.Controls.Add($lblOwner)

$txtOwner          = New-Object System.Windows.Forms.TextBox
$txtOwner.Location = New-Object System.Drawing.Point(180, 21)
$txtOwner.Size     = New-Object System.Drawing.Size(330, 22)
$txtOwner.Font     = New-Object System.Drawing.Font('Consolas', 9)
$grpOwner.Controls.Add($txtOwner)

$btnOwnerAdd          = New-Object System.Windows.Forms.Button
$btnOwnerAdd.Text     = 'Set/Add Owner'
$btnOwnerAdd.Location = New-Object System.Drawing.Point(10, 54)
$btnOwnerAdd.Size     = New-Object System.Drawing.Size(160, 28)
$grpOwner.Controls.Add($btnOwnerAdd)

$btnOwnerRemove          = New-Object System.Windows.Forms.Button
$btnOwnerRemove.Text     = 'Remove Owner'
$btnOwnerRemove.Location = New-Object System.Drawing.Point(180, 54)
$btnOwnerRemove.Size     = New-Object System.Drawing.Size(160, 28)
$grpOwner.Controls.Add($btnOwnerRemove)

# --- Membership --------------------------------------------------------------
$grpMem = New-Object System.Windows.Forms.GroupBox
$grpMem.Text     = 'Membership'
$grpMem.Location = New-Object System.Drawing.Point(12, 460)
$grpMem.Size     = New-Object System.Drawing.Size(910, 210)
$form.Controls.Add($grpMem)

$lblMemAdd          = New-Object System.Windows.Forms.Label
$lblMemAdd.Text     = 'Members to add  (one per line):'
$lblMemAdd.Location = New-Object System.Drawing.Point(10, 22)
$lblMemAdd.AutoSize = $true
$grpMem.Controls.Add($lblMemAdd)

$txtMemAdd            = New-Object System.Windows.Forms.TextBox
$txtMemAdd.Multiline  = $true
$txtMemAdd.ScrollBars = 'Vertical'
$txtMemAdd.Location   = New-Object System.Drawing.Point(10, 42)
$txtMemAdd.Size       = New-Object System.Drawing.Size(435, 110)
$txtMemAdd.Font       = New-Object System.Drawing.Font('Consolas', 9)
$grpMem.Controls.Add($txtMemAdd)

$lblMemRem          = New-Object System.Windows.Forms.Label
$lblMemRem.Text     = 'Members to remove  (one per line):'
$lblMemRem.Location = New-Object System.Drawing.Point(465, 22)
$lblMemRem.AutoSize = $true
$grpMem.Controls.Add($lblMemRem)

$txtMemRem            = New-Object System.Windows.Forms.TextBox
$txtMemRem.Multiline  = $true
$txtMemRem.ScrollBars = 'Vertical'
$txtMemRem.Location   = New-Object System.Drawing.Point(465, 42)
$txtMemRem.Size       = New-Object System.Drawing.Size(435, 110)
$txtMemRem.Font       = New-Object System.Drawing.Font('Consolas', 9)
$grpMem.Controls.Add($txtMemRem)

$btnMem          = New-Object System.Windows.Forms.Button
$btnMem.Text     = 'Process Membership Change'
$btnMem.Location = New-Object System.Drawing.Point(10, 164)
$btnMem.Size     = New-Object System.Drawing.Size(240, 32)
$grpMem.Controls.Add($btnMem)

$lblMemHint           = New-Object System.Windows.Forms.Label
$lblMemHint.Text      = 'Either box may be left empty. Adds run before removals. Nothing is written until you confirm.'
$lblMemHint.Location  = New-Object System.Drawing.Point(262, 172)
$lblMemHint.AutoSize  = $true
$lblMemHint.ForeColor = [System.Drawing.Color]::DimGray
$grpMem.Controls.Add($lblMemHint)

# --- Delivery restrictions ---------------------------------------------------
$grpDel = New-Object System.Windows.Forms.GroupBox
$grpDel.Text     = 'Delivery restrictions'
$grpDel.Location = New-Object System.Drawing.Point(12, 678)
$grpDel.Size     = New-Object System.Drawing.Size(910, 236)
$form.Controls.Add($grpDel)

$chkHidden          = New-Object System.Windows.Forms.CheckBox
$chkHidden.Text     = 'Hidden from the address book'
$chkHidden.Location = New-Object System.Drawing.Point(10, 24)
$chkHidden.AutoSize = $true
$grpDel.Controls.Add($chkHidden)

# WORDED AS THE OUTCOME, NOT THE PROPERTY.
#
# The underlying setting is RequireSenderAuthenticationEnabled, and every
# operator has at some point reasoned about it backwards -- ON blocks external
# senders. The label states what happens rather than what the flag is called.
$chkExternal          = New-Object System.Windows.Forms.CheckBox
$chkExternal.Text     = 'Accept mail from EXTERNAL senders  (turns sender authentication OFF)'
$chkExternal.Location = New-Object System.Drawing.Point(10, 50)
$chkExternal.AutoSize = $true
$grpDel.Controls.Add($chkExternal)

$btnFlags          = New-Object System.Windows.Forms.Button
$btnFlags.Text     = 'Apply these two settings'
$btnFlags.Location = New-Object System.Drawing.Point(560, 32)
$btnFlags.Size     = New-Object System.Drawing.Size(200, 30)
$grpDel.Controls.Add($btnFlags)

$lblAcc          = New-Object System.Windows.Forms.Label
$lblAcc.Text     = 'Allowed senders -- add  (one per line):'
$lblAcc.Location = New-Object System.Drawing.Point(10, 84)
$lblAcc.AutoSize = $true
$grpDel.Controls.Add($lblAcc)

$txtAccAdd            = New-Object System.Windows.Forms.TextBox
$txtAccAdd.Multiline  = $true
$txtAccAdd.ScrollBars = 'Vertical'
$txtAccAdd.Location   = New-Object System.Drawing.Point(10, 104)
$txtAccAdd.Size       = New-Object System.Drawing.Size(435, 80)
$txtAccAdd.Font       = New-Object System.Drawing.Font('Consolas', 9)
$grpDel.Controls.Add($txtAccAdd)

$lblAccRem          = New-Object System.Windows.Forms.Label
$lblAccRem.Text     = 'Allowed senders -- remove  (one per line):'
$lblAccRem.Location = New-Object System.Drawing.Point(465, 84)
$lblAccRem.AutoSize = $true
$grpDel.Controls.Add($lblAccRem)

$txtAccRem            = New-Object System.Windows.Forms.TextBox
$txtAccRem.Multiline  = $true
$txtAccRem.ScrollBars = 'Vertical'
$txtAccRem.Location   = New-Object System.Drawing.Point(465, 104)
$txtAccRem.Size       = New-Object System.Drawing.Size(435, 80)
$txtAccRem.Font       = New-Object System.Drawing.Font('Consolas', 9)
$grpDel.Controls.Add($txtAccRem)

$btnAcc          = New-Object System.Windows.Forms.Button
$btnAcc.Text     = 'Apply allowed-sender changes'
$btnAcc.Location = New-Object System.Drawing.Point(10, 192)
$btnAcc.Size     = New-Object System.Drawing.Size(240, 30)
$grpDel.Controls.Add($btnAcc)

$lblAccHint           = New-Object System.Windows.Forms.Label
$lblAccHint.Text      = 'The current list is shown above. An empty list means the group is not restricted.'
$lblAccHint.Location  = New-Object System.Drawing.Point(262, 198)
$lblAccHint.AutoSize  = $true
$lblAccHint.ForeColor = [System.Drawing.Color]::DimGray
$grpDel.Controls.Add($lblAccHint)

# --- Send-As -----------------------------------------------------------------
$grpSendAs = New-Object System.Windows.Forms.GroupBox
$grpSendAs.Text     = 'Send-As  (mail appears to come FROM the list)'
$grpSendAs.Location = New-Object System.Drawing.Point(12, 922)
$grpSendAs.Size     = New-Object System.Drawing.Size(910, 92)
$form.Controls.Add($grpSendAs)

$lblSendAs          = New-Object System.Windows.Forms.Label
$lblSendAs.Text     = 'Grantee (email or logon):'
$lblSendAs.Location = New-Object System.Drawing.Point(10, 24)
$lblSendAs.AutoSize = $true
$grpSendAs.Controls.Add($lblSendAs)

$txtSendAs          = New-Object System.Windows.Forms.TextBox
$txtSendAs.Location = New-Object System.Drawing.Point(180, 21)
$txtSendAs.Size     = New-Object System.Drawing.Size(330, 22)
$txtSendAs.Font     = New-Object System.Drawing.Font('Consolas', 9)
$grpSendAs.Controls.Add($txtSendAs)

$btnSendAsAdd          = New-Object System.Windows.Forms.Button
$btnSendAsAdd.Text     = 'Grant Send-As'
$btnSendAsAdd.Location = New-Object System.Drawing.Point(10, 54)
$btnSendAsAdd.Size     = New-Object System.Drawing.Size(160, 28)
$grpSendAs.Controls.Add($btnSendAsAdd)

$btnSendAsRem          = New-Object System.Windows.Forms.Button
$btnSendAsRem.Text     = 'Revoke Send-As'
$btnSendAsRem.Location = New-Object System.Drawing.Point(180, 54)
$btnSendAsRem.Size     = New-Object System.Drawing.Size(160, 28)
$grpSendAs.Controls.Add($btnSendAsRem)

# --- Footer ------------------------------------------------------------------
$btnRefresh          = New-Object System.Windows.Forms.Button
$btnRefresh.Text     = 'Refresh'
$btnRefresh.Location = New-Object System.Drawing.Point(12, 1024)
$btnRefresh.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnRefresh)

$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(812, 1024)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$form.Controls.Add($btnClose)
#endregion

#region Behaviour --------------------------------------------------------------
function Update-Display {
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $script:State  = Get-DLState -Identity $script:Identity
        $txtOut.Text   = Format-DLState -State $script:State

        # THE CHECKBOXES FOLLOW THE GROUP, not the other way round. Re-reading
        # after every change means a failed write cannot leave a tick implying
        # it succeeded.
        $chkHidden.Checked   = [bool]$script:State.Hidden
        $chkExternal.Checked = -not [bool]$script:State.RequireAuth
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'DL Editor'
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

function Test-CanRun {
    # Every write path calls this first. $script:Blocked was computed from the
    # environment probe, and a guard that only colours a label is not a guard.
    if ($script:Blocked) {
        [void][System.Windows.Forms.MessageBox]::Show($lblWarn.Text,
            'Cannot run in this process',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
    return $true
}

function Confirm-Change {
    param([string]$Summary)
    $r = [System.Windows.Forms.MessageBox]::Show(
        "$Summary`r`n`r`nGroup: $($script:State.PrimarySmtp)`r`nEnvironment: $(if ($script:OnPremMode) { 'ON-PREMISES' } else { 'CLOUD' })",
        'Confirm change',
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question)
    return ($r -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Set-Busy {
    param([bool]$Busy)
    foreach ($g in @($grpOwner, $grpMem, $grpDel, $grpSendAs)) { $g.Enabled = -not $Busy }
    $btnRefresh.Enabled = -not $Busy
}

$btnRefresh.Add_Click({ Update-Display })
$btnClose.Add_Click({ $form.Close() })

$btnOwnerAdd.Add_Click({
    if (-not (Test-CanRun)) { return }
    $who = $txtOwner.Text.Trim()
    if (-not $who) { $txtOut.Text = 'Enter a grantee first.'; return }
    $r = Resolve-RecipientKey -Entry $who
    if (-not $r.Ok) { $txtOut.Text = "Could not resolve '$who': $($r.Error)"; return }

    # EVERY EXISTING OWNER IS WRITTEN BACK. ManagedBy is replaced whole, so a
    # list containing only the new owner would silently remove the others.
    $keep = @($script:State.Owners | ForEach-Object { $_.Key })
    if ($keep -contains $r.Key) { $txtOut.Text = "$($r.Display) is already an owner."; return }
    $new = @($keep + $r.Key)

    if (-not (Confirm-Change -Summary "Add owner $($r.Display)?`r`nOwners will be: $($new -join ', ')")) { return }
    Set-Busy -Busy $true
    try {
        Set-DLOwners -Identity $script:Identity -Owners $new
        Write-CompletedLog -Action 'AddDLOwner' -Target $script:State.PrimarySmtp -Result 'Success' -Detail "owner: $($r.Key)"
        $txtOwner.Clear()
        Update-Display
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Add owner'
        Write-FailedLog -Action 'AddDLOwner' -Target $script:State.PrimarySmtp -ErrorRecord $_ -Detail "owner: $($r.Key)"
    } finally { Set-Busy -Busy $false }
})

$btnOwnerRemove.Add_Click({
    if (-not (Test-CanRun)) { return }
    $who = $txtOwner.Text.Trim()
    if (-not $who) { $txtOut.Text = 'Enter a grantee first.'; return }
    $r = Resolve-RecipientKey -Entry $who
    $key = if ($r.Ok) { $r.Key } else { $who }

    $keep = @($script:State.Owners | Where-Object { $_.Key -ne $key } | ForEach-Object { $_.Key })
    if ($keep.Count -eq $script:State.Owners.Count) {
        $txtOut.Text = "'$who' is not currently an owner. Nothing was changed."
        return
    }

    # A GROUP WITH NO OWNER IS A REAL STATE AND A BAD ONE. It is allowed --
    # sometimes it is what the ticket says -- but not silently.
    $warn = if ($keep.Count -eq 0) { "`r`n`r`nWARNING: this removes the LAST owner. Only an administrator will be able to manage this group." } else { '' }
    if (-not (Confirm-Change -Summary "Remove owner '$who'?$warn")) { return }

    Set-Busy -Busy $true
    try {
        Set-DLOwners -Identity $script:Identity -Owners $keep
        Write-CompletedLog -Action 'RemoveDLOwner' -Target $script:State.PrimarySmtp -Result 'Success' -Detail "owner: $key"
        $txtOwner.Clear()
        Update-Display
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Remove owner'
        Write-FailedLog -Action 'RemoveDLOwner' -Target $script:State.PrimarySmtp -ErrorRecord $_ -Detail "owner: $key"
    } finally { Set-Busy -Busy $false }
})

$btnMem.Add_Click({
    if (-not (Test-CanRun)) { return }
    $toAdd = @(ConvertTo-EntryList -Text $txtMemAdd.Text)
    $toRem = @(ConvertTo-EntryList -Text $txtMemRem.Text)
    if ($toAdd.Count -eq 0 -and $toRem.Count -eq 0) {
        $txtOut.Text = 'Nothing to do -- put at least one name in Add or Remove.'
        return
    }

    $lines = @()
    if ($toAdd.Count) { $lines += "ADD ($($toAdd.Count)):";    $lines += ($toAdd | ForEach-Object { "   $_" }) }
    if ($toRem.Count) { $lines += "REMOVE ($($toRem.Count)):"; $lines += ($toRem | ForEach-Object { "   $_" }) }
    if (-not (Confirm-Change -Summary ($lines -join "`r`n"))) { return }

    Set-Busy -Busy $true
    try {
        $out = New-Object System.Text.StringBuilder
        $ok = 0; $bad = 0

        # ADDS FIRST, THEN REMOVALS, AND THE ORDER IS NOT ARBITRARY.
        #
        # A list that both adds and removes is usually a handover. Removing
        # first means that if the run fails in the middle it fails having taken
        # access away and not yet granted it. Adding first means a mid-run
        # failure leaves too many members rather than too few -- the safer way
        # to be wrong.
        foreach ($m in $toAdd) {
            try {
                Add-DLMemberEntry -Identity $script:Identity -Member $m
                [void]$out.AppendLine("  ADD    $m  -- OK")
                Write-CompletedLog -Action 'AddDLMember' -Target $script:State.PrimarySmtp -Result 'Success' -Detail "member: $m"
                $ok++
            } catch {
                [void]$out.AppendLine("  ADD    $m  -- FAILED: $($_.Exception.Message)")
                Write-FailedLog -Action 'AddDLMember' -Target $script:State.PrimarySmtp -ErrorRecord $_ -Detail "member: $m"
                $bad++
            }
        }
        foreach ($m in $toRem) {
            try {
                Remove-DLMemberEntry -Identity $script:Identity -Member $m
                [void]$out.AppendLine("  REMOVE $m  -- OK")
                Write-CompletedLog -Action 'RemoveDLMember' -Target $script:State.PrimarySmtp -Result 'Success' -Detail "member: $m"
                $ok++
            } catch {
                [void]$out.AppendLine("  REMOVE $m  -- FAILED: $($_.Exception.Message)")
                Write-FailedLog -Action 'RemoveDLMember' -Target $script:State.PrimarySmtp -ErrorRecord $_ -Detail "member: $m"
                $bad++
            }
        }

        # CLEARED ONLY WHEN EVERYTHING WORKED. A box emptied after a partial run
        # destroys the record of which names still need attention.
        if ($bad -eq 0) { $txtMemAdd.Clear(); $txtMemRem.Clear() }

        Update-Display
        $txtOut.Text = "$ok succeeded, $bad failed.`r`n`r`n" + $out.ToString() + "`r`n" + $txtOut.Text
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Membership change'
    } finally { Set-Busy -Busy $false }
})

$btnFlags.Add_Click({
    if (-not (Test-CanRun)) { return }
    $wantHidden = [bool]$chkHidden.Checked
    $wantAuth   = -not [bool]$chkExternal.Checked

    $changes = @()
    if ($wantHidden -ne [bool]$script:State.Hidden) {
        $changes += "Hidden from address book: $($script:State.Hidden) -> $wantHidden"
    }
    if ($wantAuth -ne [bool]$script:State.RequireAuth) {
        $changes += "Sender authentication: $($script:State.RequireAuth) -> $wantAuth"
        if (-not $wantAuth -and $script:State.Accepted.Count -gt 0) {
            # THE INTERACTION IS SPELLED OUT AT THE MOMENT IT MATTERS.
            #
            # Turning authentication off does not clear the allowed-senders
            # list, it stops the list being enforced. The group goes from
            # restricted to open while still LOOKING restricted in the portal.
            $changes += "  WARNING: this group has $($script:State.Accepted.Count) allowed sender(s)."
            $changes += "  Turning sender authentication off stops that list being enforced --"
            $changes += "  the entries remain but anyone will be able to mail this group."
        }
    }
    if ($changes.Count -eq 0) { $txtOut.Text = 'Neither setting differs from the group. Nothing to apply.'; return }
    if (-not (Confirm-Change -Summary ($changes -join "`r`n"))) { return }

    Set-Busy -Busy $true
    try {
        if ($wantHidden -ne [bool]$script:State.Hidden) {
            Set-DLFlag -Identity $script:Identity -Flag 'Hidden' -Value $wantHidden
            Write-CompletedLog -Action 'SetDLHidden' -Target $script:State.PrimarySmtp -Result 'Success' -Detail "hidden: $wantHidden"
        }
        if ($wantAuth -ne [bool]$script:State.RequireAuth) {
            Set-DLFlag -Identity $script:Identity -Flag 'RequireAuth' -Value $wantAuth
            Write-CompletedLog -Action 'SetDLSenderAuth' -Target $script:State.PrimarySmtp -Result 'Success' -Detail "requireAuth: $wantAuth"
        }
        Update-Display
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Delivery settings'
        Write-FailedLog -Action 'SetDLDelivery' -Target $script:State.PrimarySmtp -ErrorRecord $_ -Detail 'flags'
        Update-Display
    } finally { Set-Busy -Busy $false }
})

$btnAcc.Add_Click({
    if (-not (Test-CanRun)) { return }
    $addRaw = @(ConvertTo-EntryList -Text $txtAccAdd.Text)
    $remRaw = @(ConvertTo-EntryList -Text $txtAccRem.Text)
    if ($addRaw.Count -eq 0 -and $remRaw.Count -eq 0) {
        $txtOut.Text = 'Nothing to do -- put at least one name in an allowed-senders box.'
        return
    }

    # THE WHOLE LIST IS REBUILT AND WRITTEN ONCE.
    #
    # AcceptMessagesOnlyFromSendersOrMembers is replaced entirely by
    # Set-DistributionGroup, so the current entries have to be carried across.
    # Everything is keyed on the resolved primary SMTP address, because
    # comparing display names would miss a match that is really there and add a
    # duplicate.
    $current = @($script:State.Accepted | ForEach-Object { $_.Key })
    $notes   = @()

    $addKeys = @()
    foreach ($a in $addRaw) {
        $r = Resolve-RecipientKey -Entry $a
        if ($r.Ok) { $addKeys += $r.Key } else { $notes += "  could not resolve (skipped): $a" }
    }
    $remKeys = @()
    foreach ($a in $remRaw) {
        $r = Resolve-RecipientKey -Entry $a
        $remKeys += $(if ($r.Ok) { $r.Key } else { $a })
    }

    $new = @($current | Where-Object { $remKeys -notcontains $_ })
    foreach ($k in $addKeys) { if ($new -notcontains $k) { $new += $k } }

    $summary = @("Allowed senders: $($current.Count) -> $($new.Count)")
    if ($notes.Count) { $summary += $notes }
    if (-not $script:State.RequireAuth) {
        $summary += ''
        $summary += 'NOTE: sender authentication is OFF, so this list is not currently'
        $summary += 'enforced. Tick "Accept mail from EXTERNAL senders" OFF and apply the'
        $summary += 'settings above for it to take effect.'
    }
    if (-not (Confirm-Change -Summary ($summary -join "`r`n"))) { return }

    Set-Busy -Busy $true
    try {
        Set-DLAcceptedSenders -Identity $script:Identity -Senders $new
        Write-CompletedLog -Action 'SetDLAcceptedSenders' -Target $script:State.PrimarySmtp -Result 'Success' `
            -Detail "count: $($current.Count) -> $($new.Count)"
        $txtAccAdd.Clear(); $txtAccRem.Clear()
        Update-Display
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Allowed senders'
        Write-FailedLog -Action 'SetDLAcceptedSenders' -Target $script:State.PrimarySmtp -ErrorRecord $_ -Detail 'accepted senders'
    } finally { Set-Busy -Busy $false }
})

$btnSendAsAdd.Add_Click({
    if (-not (Test-CanRun)) { return }
    $who = $txtSendAs.Text.Trim()
    if (-not $who) { $txtOut.Text = 'Enter a grantee first.'; return }
    $r = Resolve-RecipientKey -Entry $who
    if (-not $r.Ok) { $txtOut.Text = "Could not resolve '$who': $($r.Error)"; return }

    # THE IDENTITY DIFFERS BY SIDE. Add-ADPermission wants a sAMAccountName;
    # Add-RecipientPermission wants an address. Passing the wrong one fails in a
    # way that reads as a permissions problem.
    $key = if ($script:OnPremMode -and $r.Sam) { $r.Sam } else { $r.Key }

    if (-not (Confirm-Change -Summary "Grant Send-As to $($r.Display)?")) { return }
    Set-Busy -Busy $true
    try {
        Add-DLSendAs -Identity (Get-SendAsTarget) -User $key
        Write-CompletedLog -Action 'AddDLSendAs' -Target $script:State.PrimarySmtp -Result 'Success' -Detail "grantee: $key"
        $txtSendAs.Clear()
        Update-Display
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Grant Send-As'
        Write-FailedLog -Action 'AddDLSendAs' -Target $script:State.PrimarySmtp -ErrorRecord $_ -Detail "grantee: $key"
    } finally { Set-Busy -Busy $false }
})

$btnSendAsRem.Add_Click({
    if (-not (Test-CanRun)) { return }
    $who = $txtSendAs.Text.Trim()
    if (-not $who) { $txtOut.Text = 'Enter a grantee first.'; return }
    $r = Resolve-RecipientKey -Entry $who
    $key = if ($r.Ok) { if ($script:OnPremMode -and $r.Sam) { $r.Sam } else { $r.Key } } else { $who }

    if (-not (Confirm-Change -Summary "Revoke Send-As from '$who'?")) { return }
    Set-Busy -Busy $true
    try {
        Remove-DLSendAs -Identity (Get-SendAsTarget) -User $key
        Write-CompletedLog -Action 'RemoveDLSendAs' -Target $script:State.PrimarySmtp -Result 'Success' -Detail "grantee: $key"
        $txtSendAs.Clear()
        Update-Display
    } catch {
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Revoke Send-As'
        Write-FailedLog -Action 'RemoveDLSendAs' -Target $script:State.PrimarySmtp -ErrorRecord $_ -Detail "grantee: $key"
    } finally { Set-Busy -Busy $false }
})
#endregion

#region Start ------------------------------------------------------------------
# THE DIRECTORY IS RESOLVED BEFORE ANYTHING IS READ, so every later call uses
# the same controller. A group that cannot be found is fatal here rather than
# producing five separate failures once the window is open.
$env = Resolve-DLEnvironment -Identity $script:Identity
if (-not $env.Found) {
    [void][System.Windows.Forms.MessageBox]::Show(
        ("'$script:Identity' could not be found in $(if ($script:OnPremMode) { 'any configured on-premises directory' } else { 'Exchange Online' }).`r`n`r`n" +
         ($env.Errors -join "`r`n")),
        'DL Editor -- not found',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error)
    return
}
if ($env.Label) { $lblWarn.Text = "$($lblWarn.Text)   ($($env.Label))" }

Update-Display

Set-ConsoleTheme -Form $form
Set-FormResizable -Form $form

# BOTTOM-LEFT, like every other tool in the toolkit.
#
# StartPosition was CenterScreen, which drops a 1100px window over whatever you
# were reading -- usually the ticket you are working from.
#
# AFTER Set-FormResizable, because that sets MinimumSize and can move the form;
# cornering first would be undone by it.
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'DL Editor' | Out-Null

[void]$form.ShowDialog()
#endregion
```


<br>

==============================================================================

# ◆ FILE 11 OF 16 — `Protocols.ps1`

> **404 lines.** The tool that taught C15 (slice a large query; never retry an oversized one) and C13 (ask the cmdlet what it declares). The best single artefact on bulk queries.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```powershell
#Requires -Version 5.1
<#
    Protocols.ps1
    Reports client-access protocol state for every mailbox and exports to CSV. Author : Glenn Armbruster Revised : 2026-07-28

    Reasoning: ScriptNotes/Protocols.ps1.notes.md

    Author   : Glenn Armbruster
    Revised  : 2026-07-31
#>
# TOOLKIT-HELP-BEGIN
# Reports which client protocols are switched on for every mailbox -- POP, IMAP,
# ActiveSync, Outlook on the web and the rest -- and writes a CSV.
#
# Useful for security review: legacy protocols are the ones most often used in
# password-spray attacks, and this shows where they are still enabled.
#
# Read-only. It changes nothing.
# TOOLKIT-HELP-END

param(
    [string]$SmtpAddress,
    [string]$Environment,
    [string]$DomainController
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
$Protocols = @('OWA', 'ActiveSync', 'POP', 'IMAP', 'MAPI', 'EWS')
#endregion

#region Data -----------------------------------------------------------------
function Get-EwsOrgDefault {
    try {
        $org = Invoke-EXOCommand -Name 'Get-OrganizationConfig' -Parameters @{ ErrorAction = 'Stop' }
        if ($null -ne $org.EwsEnabled) { return [bool]$org.EwsEnabled }
    } catch { }
    return $true
}

# Get-CmdletParameterNames comes from Common.ps1. A local copy lived here
# and shadowed it -- byte-identical, so harmless, but it was a second place
# to keep in step. See ARCHITECTURE 14n.
function Invoke-CasQuery {
    param(
        [string]$Cmdlet,
        [string[]]$Props,
        [string]$RecipientType,
        [string]$Filter,
        [scriptblock]$Progress
    )
    $supported = Get-CmdletParameterNames -Name $Cmdlet

    $remaining = @($Props)
    for ($attempt = 1; $attempt -le 5; $attempt++) {
    # Only pass parameters the installed cmdlet declares -- the module surface
    # varies by version and a guessed parameter fails the whole call.
        $params = @{ ErrorAction = 'Stop' }
        if ($supported -contains 'ResultSize') { $params['ResultSize'] = 'Unlimited' }
        if ($supported -contains 'Properties') { $params['Properties'] = $remaining }
        if ($supported -contains 'PageSize')   { $params['PageSize']   = 500 }
        if ($RecipientType -and $supported -contains 'RecipientTypeDetails') {
            $params['RecipientTypeDetails'] = $RecipientType
        }
        if ($Filter -and $supported -contains 'Filter') { $params['Filter'] = $Filter }
        try {
            return @(Invoke-EXOCommand -Name $Cmdlet -Parameters $params)
        } catch {
            $msg = "$($_.Exception.Message)"

            $m = [regex]::Match($msg, 'InvalidProperties\s*=\s*(.+?)\s*$')
            if ($m.Success) {
                $bad = $m.Groups[1].Value -split '\s*,\s*' | Where-Object { $_ }
                $remaining = @($remaining | Where-Object { $bad -notcontains $_ })
                if (-not $remaining) { throw }
                & $Progress ("Dropped unsupported propert{0}: {1}. Retrying..." -f
                             $(if ($bad.Count -eq 1) { 'y' } else { 'ies' }), ($bad -join ', '))
                continue
            }

            $transientPattern = 'underlying connection was closed|operation has timed out|' +
                                'unexpected error occurred on a receive|task was canceled|' +
                                'Unable to read data from the transport|forcibly closed'
            if (($msg -match $transientPattern) -and $attempt -lt 5) {
                $wait = [math]::Pow(2, $attempt)
                & $Progress "Connection dropped. Retrying in $wait s (attempt $($attempt + 1) of 5)..."
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }
    throw "Could not query $Cmdlet after $attempt attempts."
}

function Get-ProtocolReport {
    param([scriptblock]$Progress)

    $cmdName = 'Get-EXOCASMailbox'
    if (-not (Get-Command $cmdName -ErrorAction SilentlyContinue)) {
        $cmdName = 'Get-CASMailbox'
    }

    # RecipientTypeDetails is a FILTER parameter on Get-EXOCasMailbox, not a
    # returnable property -- asking for it fails the call outright.
    $props = @('PrimarySmtpAddress', 'DisplayName',
               'OWAEnabled', 'ActiveSyncEnabled', 'PopEnabled', 'ImapEnabled',
               'MapiEnabled', 'EwsEnabled')

    $types = @('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox',
               'LinkedMailbox', 'TeamMailbox', 'DiscoveryMailbox')

    $collected = New-Object System.Collections.Generic.List[object]
    $skipped   = @()
    $script:Subdivided = @()

    # Slice ONLY if the cmdlet accepts the filter. Otherwise it is dropped and
    # every pass returns every mailbox -- duplicates stamped with the wrong type.
    $supported = Get-CmdletParameterNames -Name $cmdName
    $useTypes  = ($supported -contains 'RecipientTypeDetails')

    if ($useTypes) {
        $canFilter = ($supported -contains 'Filter')
        $letters   = @([char[]](97..122) | ForEach-Object { "$_" }) + @('0','1','2','3','4','5','6','7','8','9')

        $i = 0
        foreach ($type in $types) {
            $i++
            & $Progress "Querying $type ($i of $($types.Count))..." ([int](100 * ($i - 1) / $types.Count))
            $got = $null
            try {
                $got = @(Invoke-CasQuery -Cmdlet $cmdName -Props $props `
                                         -RecipientType $type -Progress $Progress)
            } catch {
                $firstErr = "$($_.Exception.Message)"

                if (-not $canFilter) {
                    $skipped += "$type ($firstErr)"
                    & $Progress "$type failed and cannot be subdivided -- skipping."
                    continue
                }

                & $Progress "$type too large -- splitting by first letter..."
                $script:Subdivided += $type
                $got  = New-Object System.Collections.Generic.List[object]
                $lost = @()
                $n = 0
                foreach ($ch in $letters) {
                    $n++
                    $pctBase = 100 * ($i - 1) / $types.Count
                    $pctIn    = (100.0 / $types.Count) * ($n / $letters.Count)
                    & $Progress "$type : slice $ch ($n of $($letters.Count)) -- $($collected.Count + $got.Count) so far..." ([int]($pctBase + $pctIn))
                    try {
                        $part = @(Invoke-CasQuery -Cmdlet $cmdName -Props $props `
                                                  -RecipientType $type `
                                                  -Filter "Alias -like '$ch*'" -Progress $Progress)
                        foreach ($p in $part) { $got.Add($p) }
                    } catch {
                        $lost += $ch
                    }
                }
                if ($lost.Count -gt 0) {
                    $skipped += "$type -- slices [$($lost -join ',')] failed"
                }
                if ($got.Count -eq 0) { $skipped += "$type ($firstErr)"; continue }
            }

            foreach ($s in $got) {
                Add-Member -InputObject $s -NotePropertyName '_Type' `
                           -NotePropertyValue $type -Force
                $collected.Add($s)
            }
            & $Progress "$type : $($got.Count). Running total $($collected.Count)."
        }
    } else {
        & $Progress 'Querying all mailboxes...'
        $all = @(Invoke-CasQuery -Cmdlet $cmdName -Props $props -Progress $Progress)
        foreach ($s in $all) { $collected.Add($s) }
    }

    & $Progress 'Reading organisation EWS default...'
    $ewsDefault = Get-EwsOrgDefault

    & $Progress 'Removing any duplicates...'
    $seen  = @{}
    $uniq  = New-Object System.Collections.Generic.List[object]
    $dupes = 0
    foreach ($m in $collected) {
        $k = "$($m.PrimarySmtpAddress)".ToLower()
        if (-not $k) { continue }
        if ($seen.ContainsKey($k)) { $dupes++; continue }
        $seen[$k] = $true
        $uniq.Add($m)
    }
    $collected = $uniq

    & $Progress 'Building report...'
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($m in $collected) {
        $ews = if ($null -eq $m.EwsEnabled) { $ewsDefault } else { [bool]$m.EwsEnabled }
        $rows.Add([PSCustomObject][ordered]@{
            PrimarySmtpAddress = "$($m.PrimarySmtpAddress)"
            OWA                = [bool]$m.OWAEnabled
            ActiveSync         = [bool]$m.ActiveSyncEnabled
            POP                = [bool]$m.PopEnabled
            IMAP               = [bool]$m.ImapEnabled
            MAPI               = [bool]$m.MapiEnabled
            EWS                = $ews
            RecipientType      = "$($m._Type)"
            DisplayName        = "$($m.DisplayName)"
        })
    }
    return [PSCustomObject]@{
        Rows       = $rows
        Cmdlet     = $cmdName
        EwsDefault = $ewsDefault
        Skipped    = $skipped
        Duplicates = $dupes
        Subdivided = $script:Subdivided
    }
}

function Format-ReportSummary {
    param($Result)
    $rows = $Result.Rows
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Mailboxes  : $($rows.Count)")
    [void]$sb.AppendLine("Source     : $($Result.Cmdlet)")
    [void]$sb.AppendLine("EWS default: $($Result.EwsDefault)  (used where the mailbox has no explicit setting)")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Enabled counts:')
    foreach ($p in $Protocols) {
        $on = @($rows | Where-Object { $_.$p }).Count
        $pct = if ($rows.Count) { [math]::Round(100 * $on / $rows.Count) } else { 0 }
        [void]$sb.AppendLine(("  {0,-11}{1,6} of {2}  ({3}%)" -f $p, $on, $rows.Count, $pct))
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('By recipient type:')
    $rows | Group-Object RecipientType | Sort-Object Count -Descending | ForEach-Object {
        [void]$sb.AppendLine(("  {0,-22}{1,6}" -f $_.Name, $_.Count))
    }
    if ($Result.Subdivided -and $Result.Subdivided.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("Subdivided by alias letter (too large for one call): $($Result.Subdivided -join ', ')")
        [void]$sb.AppendLine('  Slicing keys on Alias a-z/0-9. Sanity-check the total against the tenant;')
        [void]$sb.AppendLine('  an alias starting with another character would not be picked up.')
    }
    if ($Result.Duplicates -gt 0) {
        [void]$sb.AppendLine("Duplicate rows removed: $($Result.Duplicates)")
    }
    if ($Result.Skipped -and $Result.Skipped.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('*** INCOMPLETE -- these types could not be read: ***')
        foreach ($s in $Result.Skipped) { [void]$sb.AppendLine("  $s") }
    }
    return $sb.ToString()
}
#endregion

#region WinForms UI ----------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Report -- Mailbox Protocols'
$form.Size            = New-Object System.Drawing.Size(760, 656)
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif', 9)

$lblInfo          = New-Object System.Windows.Forms.Label
$lblInfo.Text     = 'Reports OWA, ActiveSync, POP, IMAP, MAPI and EWS for every mailbox in the connected tenant.'
$lblInfo.Location = New-Object System.Drawing.Point(12, 12)
$lblInfo.Size     = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblInfo)

$lblTenant           = New-Object System.Windows.Forms.Label
$lblTenant.Tag       = 'NoTheme'
$lblTenant.Location  = New-Object System.Drawing.Point(12, 34)
$lblTenant.Size      = New-Object System.Drawing.Size(720, 18)
$lblTenant.Font      = New-Object System.Drawing.Font('Microsoft Sans Serif', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblTenant)

$btnRun          = New-Object System.Windows.Forms.Button
$btnRun.Text     = 'Run Report'
$btnRun.Location = New-Object System.Drawing.Point(12, 60)
$btnRun.Size     = New-Object System.Drawing.Size(120, 30)
$form.Controls.Add($btnRun)

$btnOpen          = New-Object System.Windows.Forms.Button
$btnOpen.Text     = 'Open Reports Folder'
$btnOpen.Location = New-Object System.Drawing.Point(140, 60)
$btnOpen.Size     = New-Object System.Drawing.Size(150, 30)
$form.Controls.Add($btnOpen)


$btnClose          = New-Object System.Windows.Forms.Button
$btnClose.Text     = 'Close'
$btnClose.Location = New-Object System.Drawing.Point(622, 60)
$btnClose.Size     = New-Object System.Drawing.Size(110, 30)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$lblStatus           = New-Object System.Windows.Forms.Label
$lblStatus.Text      = 'Idle.'
$lblStatus.Location  = New-Object System.Drawing.Point(12, 134)
$lblStatus.Size      = New-Object System.Drawing.Size(720, 18)
$form.Controls.Add($lblStatus)

$pbar = Add-ReportProgressBar -Form $form -X 12 -Y 154 -Width 720 -Height 12

$txtOut            = New-Object System.Windows.Forms.TextBox
$txtOut.Multiline  = $true
$txtOut.ReadOnly   = $true
$txtOut.ScrollBars = 'Vertical'
$txtOut.Location   = New-Object System.Drawing.Point(12, 172)
$txtOut.Size       = New-Object System.Drawing.Size(720, 424)
$txtOut.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtOut)

$files = Add-ReportFileControls -Form $form -ReportDir $ReportDir -Prefix 'Protocols' `
                -X 12 -Y 96 -Width 720 -StatusLabel $lblStatus


$btnOpen.Add_Click({ Open-ReportFolder -Path $ReportDir })

$btnRun.Add_Click({
    $btnRun.Enabled = $false
    $prev = $form.Cursor
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $txtOut.Text = ''
    try {
        $progress = New-ProgressReporter -Label $lblStatus -Bar $pbar `
                        -Activity 'Mailbox protocol report'
        $report = Get-ProtocolReport -Progress $progress

        if (-not (Test-Path -LiteralPath $ReportDir)) {
            New-Item -Path $ReportDir -ItemType Directory -Force | Out-Null
        }
        if ($report.Rows.Count -eq 0) {
            $lblStatus.Text = 'No mailboxes returned -- nothing written.'
            $txtOut.Text = (Format-ReportSummary -Result $report) +
                           [Environment]::NewLine +
                           'No CSV was written because the query returned no rows.'
            return
        }

        $stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
        $csv   = Join-Path $ReportDir "Protocols_$stamp.csv"
        $report.Rows | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8

        $files.SelectPath($csv)
        Complete-ProgressReporter -Activity 'Mailbox protocol report' -Bar $pbar
        $lblStatus.Text  = "Done. $($report.Rows.Count) row(s) written."

        $txtOut.Text = (Format-ReportSummary -Result $report) +
                       [Environment]::NewLine +
                       "Saved to:" + [Environment]::NewLine + "  $csv"

        $logNote = if ($report.Skipped -and $report.Skipped.Count) { ' -- PARTIAL' } else { '' }
        # PARTIAL is the keyword the log is scanned for; Result is the same
        # fact in a column a filter can see.
        if ($logNote) { $logResult = 'Partial' } else { $logResult = 'Success' }
        Write-CompletedLog -Task "Protocol report -- $($report.Rows.Count) mailboxes$logNote" -Result $logResult
    } catch {
        $lblStatus.Text = 'Failed.'
        $txtOut.Text = Format-ErrorDetail -ErrorRecord $_ -Context 'Protocol report'
    } finally {
        Complete-ProgressReporter -Activity 'Mailbox protocol report'
        $form.Cursor    = $prev
        $btnRun.Enabled = $true
    }
})

$d = Get-ConnectedTenantDomain
if ($d) {
    $lblTenant.Text      = "Tenant: $d"
    $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(0, 128, 0)
} else {
    $lblTenant.Text      = 'Tenant: not connected -- connect to Exchange Online first.'
    $lblTenant.ForeColor = [System.Drawing.Color]::FromArgb(200, 0, 0)
}

$form.StartPosition = 'Manual'
Set-FormScreenCorner -Form $form -Corner 'BottomLeft'
Set-ConsoleTheme -Form $form
# Consistent '?' in the top-right of every tool. Added after the theme so it
# is not restyled, and BringToFront'd inside the helper because WinForms
# draws EARLIER-added controls on top.
Add-ToolHelpButton -Form $form -ScriptPath $PSCommandPath -Title 'Protocols' | Out-Null

[void]$form.ShowDialog()
#endregion
```


<br>

==============================================================================

# ◆ FILE 12 OF 16 — `Config/Tenant.TEMPLATE.json`

> **133 lines.** The ONLY file permitted to name the organisation. Hybrid, ONE AD domain. Pim is ON with DurationHours 9 to match tenant policy; SharePointOnline is deliberately OFF until a tool uses it.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```json
{
  "_comment": [
    "TENANT FACTS. This is the only file in the toolkit permitted to contain a",
    "domain name, a server name, or an organisation's internal vocabulary.",
    "Everything else reads from here. If a tenant name appears in a .ps1 file,",
    "that is a bug -- Test-Toolkit checks for it.",
    "",
    "Engineer preferences (paths, browser profiles) live in Settings.json, not",
    "here. This file is the same on every machine at a tenant; that one is not.",
    "",
    "SHAPED FOR A HYBRID TENANT WITH ONE AD DOMAIN. Directories is a LIST and",
    "the toolkit does not care how long it is -- but a template that ships as",
    "an empty array reads as 'fill in something' rather than 'fill in this',",
    "and an operator's first run then greys out every AD tool with no hint as",
    "to why. One populated entry is the shape to copy, not a limit.",
    "",
    "Rename the Id to whatever this organisation calls its directory. It is an",
    "identifier, not a label: it is written into the Environment column of the",
    "audit log and passed as -DirectoryId between tools, so short and stable",
    "beats descriptive. Label is the descriptive one and can be changed freely."
  ],

  "TenantName": "",
  "PrimaryDomain": "",
  "OnMicrosoftDomain": "",

  "_comment_Capabilities": [
    "What this tenant HAS. A false here greys out every tool that needs it,",
    "which is the intended way to hide capability rather than deleting tools.",
    "Set for hybrid: mailboxes both on-premises and in Exchange Online, with a",
    "synchronised AD. Pim is ON: rights here are eligible, not standing.",
    "",
    "SharePointOnline is OFF deliberately. The tenant has it; the first pass has",
    "no tool that uses it, and a capability flag that is true with nothing behind",
    "it puts a permanently red indicator on the console. Turn it on in the same",
    "change that adds the first SharePoint tool."
  ],
  "Capabilities": {
    "ExchangeOnline": true,
    "SharePointOnline": false,
    "OnPremAD": true,
    "OnPremExchange": true,
    "Pim": true,
    "PimTimer": true
  },

  "_comment_Directories": [
    "ONE ENTRY PER AD DOMAIN. IsDefault decides which one is assumed when a",
    "tool is not told; with a single directory it is assumed either way, but",
    "set it anyway so adding a second later does not silently change behaviour.",
    "",
    "DomainControllers is OPTIONAL and empty is the better answer unless there",
    "is a reason. Naming none lets AD locate a controller, which is what you",
    "want; naming one pins writes to that box and owns the consequences when it",
    "reboots. See ARCHITECTURE 14f -- a wrong pin is worse than no pin.",
    "",
    "The two OU values say where NEW objects go, in the canonical form ADUC",
    "shows: contoso.com/Accounts/Resource Mailbox. Blank means the tool asks",
    "each time, which is correct until someone decides the answer."
  ],
  "Directories": [
    {
      "Id": "Corp",
      "Label": "Corporate directory",
      "Domain": "",
      "DomainControllers": [],
      "SharedMailboxOu": "",
      "DistributionGroupOu": "",
      "IsDefault": true
    }
  ],

  "_comment_OnPremExchange": [
    "ServerUri is the PowerShell endpoint, e.g. http://exch01.contoso.com/PowerShell/",
    "PreferredDomainController is a fallback for tools that need one and were",
    "given no directory; leaving it blank is safe and usually right."
  ],
  "OnPremExchange": {
    "ServerUri": "",
    "PreferredDomainController": ""
  },

  "_comment_Pim": [
    "DurationHours MUST MATCH THIS TENANT'S PIM POLICY. It is 9 here.",
    "",
    "This is not a preference and not a timeout the toolkit enforces -- it is",
    "how long Entra actually grants access for, and the console header counts",
    "down from it. Set it wrong and the countdown is confidently wrong, which",
    "is worse than no countdown: an operator who trusts it will start a long",
    "job believing they have hours left.",
    "",
    "Change it here and nowhere else. No script may hardcode a duration."
  ],
  "Pim": {
    "DurationHours": 9,
    "GroupsEligibleUrl": "https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadgroup",
    "RolesEligibleUrl": "https://portal.azure.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadmigratedroles",
    "GroupsActiveUrl": "",
    "RolesActiveUrl": ""
  },

  "_comment_LicenseGroups": [
    "Group-based licensing only. Name is the group's CN, Sku identifies which",
    "licence it grants, Directory is the Id above.",
    "",
    "Purpose is what a reader needs and nothing used to record: 'Termination'",
    "means a leaver actually holds it and it is safe to remove; 'Report' means",
    "it exists so a report can find people and must NOT be removed. Getting",
    "this wrong removes someone from a group they were never in, reports the",
    "step done, and leaves the licence assigned.",
    "",
    "Example shape, delete if licensing here is not group-based:",
    "  { \"Name\": \"E5 License - Full\", \"Sku\": \"E5\", \"Directory\": \"Corp\", \"Purpose\": \"Termination\" }"
  ],
  "LicenseGroups": [],

  "FeatureRequest": {
    "_comment": [
      "Where the Feature Request tool sends requests. The operator never sees",
      "this address -- it goes on BCC -- but it is not a secret: it is in a",
      "config file anyone can open. It lives here rather than in the script",
      "because Test-Toolkit rejects a mail domain written into a .ps1."
    ],
    "Recipient": ""
  },

  "SharePoint": {
    "AdminUrl": "",
    "MySiteHost": ""
  },

  "Theme": ""
}
```


<br>

==============================================================================

# ◆ FILE 13 OF 16 — `Config/Settings.json`

> **12 lines.** Engineer preferences -- paths, browser profiles. Differs per machine; must never hold tenant facts.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```json
{
  "ReportsPath": "",
  "LogPath": "",
  "RetryQueuePath": "",
  "OperatorEmail": "",
  "SpoAdminUrl": "",
  "ExchangeServerUri": "",
  "BrowserProfiles": {
    "Admin": "",
    "Standard": ""
  }
}
```


<br>

==============================================================================

# ◆ FILE 14 OF 16 — `Config/Sections.json`

> **21 lines.** Console layout. A section holds Rows x ButtonsPerRow buttons; register more and the surplus has nowhere to draw.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```json
{
    "ButtonsPerRow": 5,
    "Sections": [
        {
            "Name": "Admin Tools",
            "Rows": 4
        },
        {
            "Name": "Daily Scripts",
            "Rows": 2
        },
        {
            "Name": "Reporting",
            "Rows": 3
        },
        {
            "Name": "Connect",
            "Rows": 1
        }
    ]
}
```


<br>

==============================================================================

# ◆ FILE 15 OF 16 — `Config/Tools.json`

> **373 lines.** Button registry. Reference for SCHEMA, NOT CONTENTS -- its entries are tools from the source toolkit, most out of scope here.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```json
[
  {
    "Name": "Dashboard",
    "Script": ".\\Dashboard.ps1",
    "Section": "Admin Tools",
    "Order": 0,
    "ObjectTypes": []
  },
  {
    "Order": 1,
    "Name": "Group Editor",
    "Script": ".\\GroupEditor.ps1",
    "ObjectTypes": [
      "UnifiedGroup",
      "OnpremGroup"
    ],
    "Section": "Admin Tools",
    "Tier": "Tier3",
    "NewProcess": false
  },
  {
    "Order": 2,
    "Name": "Mailbox Health",
    "Script": ".\\MailboxHealth.ps1",
    "ObjectTypes": [
      "Mailbox"
    ],
    "Section": "Admin Tools",
    "Tier": "Tier1",
    "NewProcess": false
  },
  {
    "Order": 3,
    "Name": "Message Trace",
    "Script": ".\\MessageTrace.ps1",
    "ObjectTypes": [
      "Mailbox"
    ],
    "Section": "Admin Tools",
    "Tier": "Tier1",
    "NewProcess": false
  },
  {
    "Order": 4,
    "Name": "Header Analyzer",
    "Script": ".\\headeranalyzer.ps1",
    "ObjectTypes": [],
    "Section": "Admin Tools",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 5,
    "Name": "Inbox Rules",
    "Script": ".\\InboxRules.ps1",
    "ObjectTypes": [
      "Mailbox"
    ],
    "Section": "Admin Tools",
    "Tier": "Tier1",
    "NewProcess": false
  },
  {
    "Order": 6,
    "Name": "Calendar Permissions",
    "Script": ".\\CalendarPermissions.ps1",
    "ObjectTypes": [
      "Mailbox"
    ],
    "Section": "Admin Tools",
    "Tier": "Tier2",
    "NewProcess": false
  },
  {
    "Order": 7,
    "Name": "Out of Office",
    "Script": ".\\OOFEditor.ps1",
    "ObjectTypes": [
      "Mailbox"
    ],
    "Section": "Admin Tools",
    "Tier": "Tier2",
    "NewProcess": false
  },
  {
    "Order": 8,
    "Name": "DL Memberships",
    "Script": ".\\DLMemberships.ps1",
    "ObjectTypes": [
      "Mailbox"
    ],
    "Section": "Admin Tools",
    "Tier": "Tier2",
    "NewProcess": false
  },
  {
    "Order": 9,
    "Name": "UnifiedGroup Editor",
    "Script": ".\\unifiedgroup.ps1",
    "ObjectTypes": [
      "UnifiedGroup"
    ],
    "Section": "Admin Tools",
    "Tier": "Tier1",
    "NewProcess": false
  },
  {
    "Order": 10,
    "Name": "Mailbox Permissions",
    "Script": ".\\MailboxPermissions.ps1",
    "ObjectTypes": [
      "Mailbox"
    ],
    "Section": "Admin Tools",
    "Tier": "Tier2",
    "NewProcess": false
  },
  {
    "Order": 11,
    "Name": "Mail Attributes (Rename)",
    "Script": ".\\MailAttributes.ps1",
    "ObjectTypes": [
      "Mailbox"
    ],
    "Section": "Admin Tools",
    "Tier": "Tier3",
    "NewProcess": false
  },
  {
    "Order": 12,
    "Name": "All Attributes",
    "Script": ".\\AllAttributes.ps1",
    "ObjectTypes": [
      "Mailbox"
    ],
    "Section": "Admin Tools",
    "Tier": "Tier3",
    "NewProcess": false
  },
  {
    "Order": 13,
    "Name": "Audit UserMB",
    "Script": ".\\audituser.ps1",
    "ObjectTypes": [
      "Mailbox"
    ],
    "Section": "Admin Tools",
    "Tier": "Tier2",
    "NewProcess": false
  },
  {
    "Order": 14,
    "Name": "Audit SMB",
    "Script": ".\\auditSMB.ps1",
    "ObjectTypes": [
      "Mailbox"
    ],
    "Section": "Admin Tools",
    "Tier": "Tier2",
    "NewProcess": false
  },
  {
    "Order": 17,
    "Name": "New User",
    "Script": ".\\newuser.ps1",
    "ObjectTypes": [
      "Mailbox"
    ],
    "Section": "Daily Scripts",
    "Tier": "Tier1",
    "NewProcess": false
  },
  {
    "Order": 19,
    "Name": "Console",
    "Script": ".\\console.ps1",
    "ObjectTypes": [
      "Mailbox"
    ],
    "Section": "Admin Tools",
    "Tier": "Tier3",
    "NewProcess": false
  },
  {
    "Order": 20,
    "Name": "Mailbox / Onedrive Perms",
    "Script": ".\\mailboxonedriveaccess.ps1",
    "ObjectTypes": [],
    "Section": "Daily Scripts",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 23,
    "Name": "ResourceDuplicator",
    "Script": "./resourceduplicator.ps1",
    "ObjectTypes": [],
    "Section": "Admin Tools",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 28,
    "Name": "Protocols",
    "Script": ".\\Protocols.ps1",
    "ObjectTypes": [],
    "Section": "Reporting",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 29,
    "Name": "Disabled/Licensed",
    "Script": ".\\disabledlicensed.ps1",
    "ObjectTypes": [],
    "Section": "Reporting",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 30,
    "Name": "MB w No activity",
    "Script": ".\\staleaccounts.ps1",
    "ObjectTypes": [],
    "Section": "Reporting",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 31,
    "Name": "DL Inventory",
    "Script": ".\\dlinventory.ps1",
    "ObjectTypes": [],
    "Section": "Reporting",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 32,
    "Name": "External Forwarding",
    "Script": ".\\externalForwarding.ps1",
    "ObjectTypes": [],
    "Section": "Reporting",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 33,
    "Name": "Permissions Audit",
    "Script": "./permissionsaudit.ps1",
    "ObjectTypes": [],
    "Section": "Reporting",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 34,
    "Name": "Orphaned OneDrive",
    "Script": "./orphanedOneDrive.ps1",
    "ObjectTypes": [],
    "Section": "Reporting",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 35,
    "Name": "External Sharing",
    "Script": "./externalSharing.ps1",
    "ObjectTypes": [],
    "Section": "Reporting",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 36,
    "Name": "Ownerless Groups",
    "Script": "./ownerlessgroups.ps1",
    "ObjectTypes": [],
    "Section": "Reporting",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 37,
    "Name": "Report Browser",
    "Script": "./reportBrowser.ps1",
    "ObjectTypes": [],
    "Section": "Reporting",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 38,
    "Name": "M365 Reports",
    "Script": ".\\m365Reports.ps1",
    "ObjectTypes": [],
    "Section": "Reporting",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 39,
    "Name": "Log Viewer",
    "Script": ".\\logviewer.ps1",
    "ObjectTypes": [],
    "Section": "Reporting",
    "Tier": "tier3",
    "NewProcess": false
  },
  {
    "Order": 40,
    "Name": "Activate",
    "Script": ".\\activate.ps1",
    "ObjectTypes": [],
    "Section": "Connect",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 41,
    "Name": "OnPrem",
    "Script": ".\\connect-exse.ps1",
    "ObjectTypes": [],
    "Section": "Connect",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 42,
    "Name": "ExO",
    "Script": ".\\connect-exo.ps1",
    "ObjectTypes": [],
    "Section": "Connect",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 43,
    "Name": "SPO",
    "Script": ".\\connect-spo.ps1",
    "ObjectTypes": [],
    "Section": "Connect",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 44,
    "Name": "Cert Check",
    "Script": ".\\certificateexpiry.ps1",
    "ObjectTypes": [],
    "Section": "Admin Tools",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 46,
    "Name": "Tester",
    "Script": ".\\test-toolkit.ps1",
    "ObjectTypes": [],
    "Section": "Admin Tools",
    "Tier": "",
    "NewProcess": false
  },
  {
    "Order": 48,
    "Name": "Quick Links",
    "Script": ".\\quicklaunch.ps1",
    "ObjectTypes": [],
    "Section": "Daily Scripts",
    "Tier": "",
    "NewProcess": false
  }
]
```


<br>

==============================================================================

# ◆ FILE 16 OF 16 — `Config/Themes.TEMPLATE.json`

> **579 lines.** Palettes. TAKE AS-IS -- one working theme is in scope; the ThemeManager editor is not. SelectedText is not decoration: a radio/checkbox glyph cannot be coloured, so the theme carries the legible "this one is selected" colour.
>
> *Standalone file. Reproduce under this exact name.*

==============================================================================

```json
{
  "ActiveTheme": "Default",
  "FollowTenant": false,
  "TenantThemes": {},
  "Themes": {
    "Default": {
      "Description": "Windows default. No colours applied.",
      "Enabled": false,
      "FormBack": "",
      "FormText": "",
      "PanelBack": "",
      "GroupText": "",
      "ButtonBack": "",
      "ButtonText": "",
      "ButtonBorder": "",
      "ButtonHover": "",
      "ButtonDown": "",
      "InputBack": "",
      "InputText": "",
      "GradientTo": "",
      "Logo": "",
      "SelectedText": "",
      "ServiceNowBack": "#0A5C42",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#0E7856",
      "ServiceNowBorder": "#16A072",
      "QuickLaunchAdminBack": "#122D5C",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#1B448A",
      "QuickLaunchAdminBorder": "#407AD6"
    },
    "Midnight": {
      "Description": "Dark grey. Easy on the eyes for long sessions.",
      "Enabled": true,
      "FormBack": "#1F1F23",
      "FormText": "#E4E4E7",
      "PanelBack": "#27272E",
      "GroupText": "#A1A1AA",
      "ButtonBack": "#33333C",
      "ButtonText": "#F4F4F5",
      "ButtonBorder": "#52525B",
      "ButtonHover": "#3F3F4A",
      "ButtonDown": "#2A2A31",
      "InputBack": "#161619",
      "InputText": "#E4E4E7",
      "GradientTo": "#141417",
      "Logo": "",
      "SelectedText": "#7FD99A",
      "ServiceNowBack": "#0A5C42",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#0E7856",
      "ServiceNowBorder": "#16A072",
      "QuickLaunchAdminBack": "#122D5C",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#1B448A",
      "QuickLaunchAdminBorder": "#407AD6"
    },
    "Christmas": {
      "Description": "Seasonal. Deep green with red and gold.",
      "Enabled": true,
      "FormBack": "#0B3D20",
      "FormText": "#F5E6C8",
      "PanelBack": "#0E4A28",
      "GroupText": "#D4B483",
      "ButtonBack": "#C8102E",
      "ButtonText": "#FFFFFF",
      "ButtonBorder": "#D4AF37",
      "ButtonHover": "#E03048",
      "ButtonDown": "#A00D25",
      "InputBack": "#06210F",
      "InputText": "#F5E6C8",
      "GradientTo": "#06210F",
      "Logo": "",
      "SelectedText": "#FFD166",
      "ServiceNowBack": "#0A5C42",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#0E7856",
      "ServiceNowBorder": "#16A072",
      "QuickLaunchAdminBack": "#122D5C",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#1B448A",
      "QuickLaunchAdminBorder": "#407AD6"
    },
    "Slate": {
      "Description": "Neutral blue-grey. Calm and low-contrast for long sessions.",
      "Enabled": true,
      "FormBack": "#1E2530",
      "FormText": "#E6EAF0",
      "PanelBack": "#263041",
      "GroupText": "#9FB3C8",
      "ButtonBack": "#3B4A61",
      "ButtonText": "#F0F4F8",
      "ButtonBorder": "#55677F",
      "ButtonHover": "#4C5F7A",
      "ButtonDown": "#2E3A4D",
      "InputBack": "#161C25",
      "InputText": "#E6EAF0",
      "GradientTo": "#141A22",
      "Logo": "",
      "SelectedText": "#7FD1FF",
      "ServiceNowBack": "#0A5C42",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#0E7856",
      "ServiceNowBorder": "#16A072",
      "QuickLaunchAdminBack": "#122D5C",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#1B448A",
      "QuickLaunchAdminBorder": "#407AD6"
    },
    "Nordic": {
      "Description": "Cool arctic palette. Soft contrast, easy on tired eyes.",
      "Enabled": true,
      "FormBack": "#2E3440",
      "FormText": "#ECEFF4",
      "PanelBack": "#3B4252",
      "GroupText": "#88C0D0",
      "ButtonBack": "#434C5E",
      "ButtonText": "#ECEFF4",
      "ButtonBorder": "#5E81AC",
      "ButtonHover": "#4C566A",
      "ButtonDown": "#353C4A",
      "InputBack": "#292E39",
      "InputText": "#ECEFF4",
      "GradientTo": "#272C36",
      "Logo": "",
      "SelectedText": "#88C0D0",
      "ServiceNowBack": "#0A5C42",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#0E7856",
      "ServiceNowBorder": "#16A072",
      "QuickLaunchAdminBack": "#122D5C",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#1B448A",
      "QuickLaunchAdminBorder": "#407AD6"
    },
    "Solarized": {
      "Description": "The classic Solarized dark palette. Warm and readable.",
      "Enabled": true,
      "FormBack": "#002B36",
      "FormText": "#EEE8D5",
      "PanelBack": "#073642",
      "GroupText": "#93A1A1",
      "ButtonBack": "#0E4A58",
      "ButtonText": "#EEE8D5",
      "ButtonBorder": "#268BD2",
      "ButtonHover": "#14606F",
      "ButtonDown": "#062B33",
      "InputBack": "#00212B",
      "InputText": "#EEE8D5",
      "GradientTo": "#001E26",
      "Logo": "",
      "SelectedText": "#2AA198",
      "ServiceNowBack": "#4F6A00",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#647F00",
      "ServiceNowBorder": "#859900",
      "QuickLaunchAdminBack": "#122D5C",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#1B448A",
      "QuickLaunchAdminBorder": "#407AD6"
    },
    "Graphite": {
      "Description": "Near-black with amber accents. Minimal and sharp.",
      "Enabled": true,
      "FormBack": "#1A1A1A",
      "FormText": "#ECECEC",
      "PanelBack": "#242424",
      "GroupText": "#B0B0B0",
      "ButtonBack": "#333333",
      "ButtonText": "#F5F5F5",
      "ButtonBorder": "#4D4D4D",
      "ButtonHover": "#454545",
      "ButtonDown": "#262626",
      "InputBack": "#131313",
      "InputText": "#ECECEC",
      "GradientTo": "#0F0F0F",
      "Logo": "",
      "SelectedText": "#FFB300",
      "ServiceNowBack": "#7A4F00",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#96620A",
      "ServiceNowBorder": "#C08400",
      "QuickLaunchAdminBack": "#122D5C",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#1B448A",
      "QuickLaunchAdminBorder": "#407AD6"
    },
    "Forest": {
      "Description": "Mid green. Distinct from the blues without going dark.",
      "Enabled": true,
      "FormBack": "#22402F",
      "FormText": "#EAF3ED",
      "PanelBack": "#2E5741",
      "GroupText": "#A8CDB6",
      "ButtonBack": "#3E7057",
      "ButtonText": "#F2F8F4",
      "ButtonBorder": "#54916F",
      "ButtonHover": "#4A8465",
      "ButtonDown": "#2F5843",
      "InputBack": "#1B3325",
      "InputText": "#EAF3ED",
      "GradientTo": "#1A3325",
      "Logo": "",
      "SelectedText": "#8FE3AF",
      "ServiceNowBack": "#0A5C42",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#0E7856",
      "ServiceNowBorder": "#16A072",
      "QuickLaunchAdminBack": "#122D5C",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#1B448A",
      "QuickLaunchAdminBorder": "#407AD6"
    },
    "Plum": {
      "Description": "Muted purple. Useful when several consoles are open at once.",
      "Enabled": true,
      "FormBack": "#221A2E",
      "FormText": "#EFE8F5",
      "PanelBack": "#2E2340",
      "GroupText": "#B9A5CC",
      "ButtonBack": "#46345E",
      "ButtonText": "#F3EDF8",
      "ButtonBorder": "#5D4880",
      "ButtonHover": "#574073",
      "ButtonDown": "#33264A",
      "InputBack": "#1A1424",
      "InputText": "#EFE8F5",
      "GradientTo": "#171021",
      "Logo": "",
      "SelectedText": "#C89BFF",
      "ServiceNowBack": "#4A3A72",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#5D4A8E",
      "ServiceNowBorder": "#7B63B8",
      "QuickLaunchAdminBack": "#122D5C",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#1B448A",
      "QuickLaunchAdminBorder": "#407AD6"
    },
    "Daylight": {
      "Description": "Light theme. For bright rooms and projectors.",
      "Enabled": true,
      "FormBack": "#F4F6F9",
      "FormText": "#1F2933",
      "PanelBack": "#FFFFFF",
      "GroupText": "#52606D",
      "ButtonBack": "#2A6FB5",
      "ButtonText": "#FFFFFF",
      "ButtonBorder": "#1E5A96",
      "ButtonHover": "#3A82CC",
      "ButtonDown": "#1B5188",
      "InputBack": "#FFFFFF",
      "InputText": "#1F2933",
      "GradientTo": "",
      "Logo": "",
      "SelectedText": "#12507E",
      "ServiceNowBack": "#1E7A46",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#269556",
      "ServiceNowBorder": "#2FB46A",
      "QuickLaunchAdminBack": "#2A6FB5",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#3A82CC",
      "QuickLaunchAdminBorder": "#1E5A96"
    },
    "Contrast": {
      "Description": "Maximum contrast. Accessibility and difficult screens.",
      "Enabled": true,
      "FormBack": "#000000",
      "FormText": "#FFFFFF",
      "PanelBack": "#0A0A0A",
      "GroupText": "#FFFF00",
      "ButtonBack": "#000000",
      "ButtonText": "#FFFFFF",
      "ButtonBorder": "#FFFFFF",
      "ButtonHover": "#333333",
      "ButtonDown": "#1A1A1A",
      "InputBack": "#FFFFFF",
      "InputText": "#000000",
      "GradientTo": "",
      "Logo": "",
      "SelectedText": "#00FF00",
      "ServiceNowBack": "#004D2E",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#006B40",
      "ServiceNowBorder": "#00A060",
      "QuickLaunchAdminBack": "#000000",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#333333",
      "QuickLaunchAdminBorder": "#FFFFFF"
    },
    "Sapphire": {
      "Description": "Deep corporate navy. The classic dark blue admin console.",
      "Enabled": true,
      "FormBack": "#0F2B5B",
      "FormText": "#DCE6F5",
      "PanelBack": "#14356E",
      "GroupText": "#9DB8E0",
      "ButtonBack": "#1E4D94",
      "ButtonText": "#FFFFFF",
      "ButtonBorder": "#3D7BD4",
      "ButtonHover": "#2A63B8",
      "ButtonDown": "#123A73",
      "InputBack": "#0A1F42",
      "InputText": "#E8F0FC",
      "GradientTo": "#071A3C",
      "Logo": "",
      "SelectedText": "#7FD4FF",
      "ServiceNowBack": "#0A5C42",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#0E7856",
      "ServiceNowBorder": "#16A072",
      "QuickLaunchAdminBack": "#122D5C",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#1B448A",
      "QuickLaunchAdminBorder": "#407AD6"
    },
    "Burgundy": {
      "Description": "Muted wine red. Warm without being loud.",
      "Enabled": true,
      "FormBack": "#2B1A20",
      "FormText": "#F2E6EA",
      "PanelBack": "#3A242C",
      "GroupText": "#C9A5B0",
      "ButtonBack": "#5E3644",
      "ButtonText": "#F7EDF0",
      "ButtonBorder": "#7E4A5C",
      "ButtonHover": "#6E4152",
      "ButtonDown": "#472A34",
      "InputBack": "#211318",
      "InputText": "#F2E6EA",
      "GradientTo": "#1E1216",
      "Logo": "",
      "SelectedText": "#E8A0B4",
      "ServiceNowBack": "#5E3020",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#763E2A",
      "ServiceNowBorder": "#9E5438",
      "QuickLaunchAdminBack": "#122D5C",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#1B448A",
      "QuickLaunchAdminBorder": "#407AD6"
    },
    "Teal": {
      "Description": "Blue-green. Reads as calm rather than corporate.",
      "Enabled": true,
      "FormBack": "#122E30",
      "FormText": "#E4F2F2",
      "PanelBack": "#194042",
      "GroupText": "#93C4C6",
      "ButtonBack": "#256A6D",
      "ButtonText": "#EFF9F9",
      "ButtonBorder": "#31898D",
      "ButtonHover": "#2E7C80",
      "ButtonDown": "#1B4C4E",
      "InputBack": "#0D2426",
      "InputText": "#E4F2F2",
      "GradientTo": "#0C2224",
      "Logo": "",
      "SelectedText": "#7FE0DF",
      "ServiceNowBack": "#0A5C42",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#0E7856",
      "ServiceNowBorder": "#16A072",
      "QuickLaunchAdminBack": "#122D5C",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#1B448A",
      "QuickLaunchAdminBorder": "#407AD6"
    },
    "Copper": {
      "Description": "Warm rust. Stands out sharply beside any blue.",
      "Enabled": true,
      "FormBack": "#2B1E16",
      "FormText": "#F5EAE1",
      "PanelBack": "#3A2A1F",
      "GroupText": "#CBA98F",
      "ButtonBack": "#6B4429",
      "ButtonText": "#FAF2EA",
      "ButtonBorder": "#8F5C36",
      "ButtonHover": "#7E5231",
      "ButtonDown": "#503322",
      "InputBack": "#211710",
      "InputText": "#F5EAE1",
      "GradientTo": "#1E1610",
      "Logo": "",
      "SelectedText": "#F0AE72",
      "ServiceNowBack": "#6B4A18",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#856022",
      "ServiceNowBorder": "#A87C2E",
      "QuickLaunchAdminBack": "#122D5C",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#1B448A",
      "QuickLaunchAdminBorder": "#407AD6"
    },
    "Indigo": {
      "Description": "Deep blue-violet. Darker and cooler than the default blue.",
      "Enabled": true,
      "FormBack": "#1C1F3A",
      "FormText": "#E7E9F7",
      "PanelBack": "#262A4D",
      "GroupText": "#A5AAD6",
      "ButtonBack": "#3A4177",
      "ButtonText": "#F0F1FA",
      "ButtonBorder": "#4E5799",
      "ButtonHover": "#454D89",
      "ButtonDown": "#2C3159",
      "InputBack": "#15182C",
      "InputText": "#E7E9F7",
      "GradientTo": "#13162A",
      "Logo": "",
      "SelectedText": "#9AA7FF",
      "ServiceNowBack": "#0A5C42",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#0E7856",
      "ServiceNowBorder": "#16A072",
      "QuickLaunchAdminBack": "#122D5C",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#1B448A",
      "QuickLaunchAdminBorder": "#407AD6"
    },
    "Steel": {
      "Description": "Mid blue-grey. Lighter than Slate, still a dark theme.",
      "Enabled": true,
      "FormBack": "#2C3540",
      "FormText": "#EAEEF3",
      "PanelBack": "#3A4653",
      "GroupText": "#A9B8C7",
      "ButtonBack": "#4E5F72",
      "ButtonText": "#F2F6FA",
      "ButtonBorder": "#6C8199",
      "ButtonHover": "#5D7185",
      "ButtonDown": "#3E4B5A",
      "InputBack": "#232B34",
      "InputText": "#EAEEF3",
      "GradientTo": "#222A33",
      "Logo": "",
      "SelectedText": "#8FCBEF",
      "ServiceNowBack": "#0A5C42",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#0E7856",
      "ServiceNowBorder": "#16A072",
      "QuickLaunchAdminBack": "#122D5C",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#1B448A",
      "QuickLaunchAdminBorder": "#407AD6"
    },
    "Ocean": {
      "Description": "Deep blue-teal. Cooler and quieter than the default blue.",
      "Enabled": true,
      "FormBack": "#0E2733",
      "FormText": "#E2EFF5",
      "PanelBack": "#153545",
      "GroupText": "#93B8C9",
      "ButtonBack": "#1F5570",
      "ButtonText": "#EEF7FB",
      "ButtonBorder": "#2C7395",
      "ButtonHover": "#276685",
      "ButtonDown": "#193F53",
      "InputBack": "#0B1E28",
      "InputText": "#E2EFF5",
      "GradientTo": "#0A1B23",
      "Logo": "",
      "SelectedText": "#7FD2F0",
      "ServiceNowBack": "#0A5C42",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#0E7856",
      "ServiceNowBorder": "#16A072",
      "QuickLaunchAdminBack": "#122D5C",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#1B448A",
      "QuickLaunchAdminBorder": "#407AD6"
    },
    "Cobalt": {
      "Description": "Vivid blue. Brighter and more saturated than the default.",
      "Enabled": true,
      "FormBack": "#0D1B3E",
      "FormText": "#E4EAF8",
      "PanelBack": "#132552",
      "GroupText": "#9BAEDC",
      "ButtonBack": "#1E3E8C",
      "ButtonText": "#EEF2FC",
      "ButtonBorder": "#3A5FBF",
      "ButtonHover": "#2A4EA6",
      "ButtonDown": "#16295C",
      "InputBack": "#0A1531",
      "InputText": "#E4EAF8",
      "GradientTo": "#09122A",
      "Logo": "",
      "SelectedText": "#8FB4FF",
      "ServiceNowBack": "#0A5C42",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#0E7856",
      "ServiceNowBorder": "#16A072",
      "QuickLaunchAdminBack": "#122D5C",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#1B448A",
      "QuickLaunchAdminBorder": "#407AD6"
    },
    "Olive": {
      "Description": "Muted green-brown. Low saturation, easy for long sessions.",
      "Enabled": true,
      "FormBack": "#25281A",
      "FormText": "#EDEFE2",
      "PanelBack": "#323622",
      "GroupText": "#B7BF97",
      "ButtonBack": "#4C5330",
      "ButtonText": "#F4F6EB",
      "ButtonBorder": "#67713F",
      "ButtonHover": "#5B653A",
      "ButtonDown": "#383D25",
      "InputBack": "#1D1F14",
      "InputText": "#EDEFE2",
      "GradientTo": "#1B1D12",
      "Logo": "",
      "SelectedText": "#C4D98A",
      "ServiceNowBack": "#0A5C42",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#0E7856",
      "ServiceNowBorder": "#16A072",
      "QuickLaunchAdminBack": "#122D5C",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#1B448A",
      "QuickLaunchAdminBorder": "#407AD6"
    },
    "Mocha": {
      "Description": "Warm dark brown. Softer than a pure grey at night.",
      "Enabled": true,
      "FormBack": "#251E1A",
      "FormText": "#F0E8E2",
      "PanelBack": "#332924",
      "GroupText": "#C0AA9C",
      "ButtonBack": "#4E3C33",
      "ButtonText": "#F7F0EA",
      "ButtonBorder": "#6B5348",
      "ButtonHover": "#5F4940",
      "ButtonDown": "#3A2D27",
      "InputBack": "#1D1714",
      "InputText": "#F0E8E2",
      "GradientTo": "#1A1512",
      "Logo": "",
      "SelectedText": "#E0B48F",
      "ServiceNowBack": "#4E3A1C",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#644B26",
      "ServiceNowBorder": "#836333",
      "QuickLaunchAdminBack": "#122D5C",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#1B448A",
      "QuickLaunchAdminBorder": "#407AD6"
    },
    "Sandstone": {
      "Description": "Warm light theme. Softer than Daylight under bright light.",
      "Enabled": true,
      "FormBack": "#F5F1E8",
      "FormText": "#2B2620",
      "PanelBack": "#FFFDF8",
      "GroupText": "#5E564A",
      "ButtonBack": "#8A6A3E",
      "ButtonText": "#FFFFFF",
      "ButtonBorder": "#6E5330",
      "ButtonHover": "#A07E4C",
      "ButtonDown": "#5E4728",
      "InputBack": "#FFFFFF",
      "InputText": "#2B2620",
      "GradientTo": "",
      "Logo": "",
      "SelectedText": "#6B4E22",
      "ServiceNowBack": "#3F6B47",
      "ServiceNowText": "#FFFFFF",
      "ServiceNowHover": "#4F8459",
      "ServiceNowBorder": "#66A472",
      "QuickLaunchAdminBack": "#8A6A3E",
      "QuickLaunchAdminText": "#FFFFFF",
      "QuickLaunchAdminHover": "#A07E4C",
      "QuickLaunchAdminBorder": "#6E5330"
    }
  }
}
```


<br>

==============================================================================

# END OF REFERENCE CODE

Sixteen files. The instructions that govern their use are in `DEVIN-PACK.md`.
