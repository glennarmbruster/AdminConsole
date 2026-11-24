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
